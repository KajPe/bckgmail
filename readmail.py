#!/env/python3
# -*- coding: utf-8 -*-

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64, email
import sys
import glob

# Scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.labels'
]


# =================================================================================
class bckgmail:
    _base    = ''
    _labels  = {}
    _service = None
    _creds   = None

    # =============================================================================
    # Initialize class
    def __init__(self, path):
        self._labels = {}
        self._base   = os.path.dirname(os.path.abspath(path))

    # =============================================================================
    # Return basepath
    def getBasePath(self):
        return self._base

    # =============================================================================
    # Create folder under base path
    def createDir(self, path):
        folder = os.path.join(self._base, path)
        if not os.path.exists(folder):
            os.makedirs(folder)

    # =============================================================================
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    def getToken(self, reset=False):
        self._creds = None

        f = os.path.join(self._base, 'token.pickle')

        if reset and os.path.exists(f):
            # Remove token.pickle
            os.remove(f)

        if os.path.exists(f):
            # Load existing tocken.pickle
            with open(f, 'rb') as token:
                self._creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not self._creds or not self._creds.valid:
            if self._creds and self._creds.expired and self._creds.refresh_token:
                self._creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(os.path.join(self._base, 'credentials.json'), SCOPES)
                self._creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(f, 'wb') as token:
                pickle.dump(self._creds, token)

    # =============================================================================
    # Get service for GMail
    def getService(self):
        self._service = build('gmail', 'v1', credentials = self._creds)

    # =============================================================================
    # Get GMail labels
    def getGMailLabels(self):
        self._labels = {}
        results = self._service.users().labels().list(
            userId = 'me'
        ).execute()
        labels = results.get('labels', [])
        if labels:
            for label in labels:
                self._labels[label['name']] = label['id']

    # =============================================================================
    # Return labels (first call getGMailLabels)
    def getLabels(self):
        return self._labels

    # =============================================================================
    # Get ID for label (or -1 if not found)
    def getLabelId(self, label):
        return self._labels.get(label, -1)

    # =============================================================================
    # Get GMail message as "full"
    def getMessageAsFull(self, id):
        return self._service.users().messages().get(
            userId = 'me',
            id     = id,
            format = "full"
        ).execute()

    # =============================================================================
    # Get GMail message as "raw"
    def getMessageAsRaw(self, id):
        return self._service.users().messages().get(
            userId = "me",
            id     = id,
            format = "raw",
            metadataHeaders = None
        ).execute()

    # =============================================================================
    # Save GMail message as eml file
    def saveMessageAsEML(self, id, path, force=False):
        outfile_name = os.path.join(self._base, path, f'{id}.eml')

        b = True if force else not os.path.exists(outfile_name)
        if b:
            message  = self.getMessageAsRaw(id)
            msg_str  = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            mime_msg = email.message_from_string(msg_str.decode('cp1252'))

            with open(outfile_name, 'w') as outfile:
                gen = email.generator.Generator(outfile)
                gen.flatten(mime_msg)
            print("Message saved:", outfile_name)
        else:
            print("Message already exists:", outfile_name)

    # =============================================================================
    # Get all GMail message id's
    def getAllMessages(self, label):
        labelId = self.getLabelId(label)
        if labelId == -1:
            return False

        all_message_in_label = []

        response = self._service.users().messages().list(
            userId           = "me",
            labelIds         = [ labelId ],
            q                = None,
            pageToken        = None,
            maxResults       = None,
            includeSpamTrash = None
        ).execute()
        if 'messages' in response:
            all_message_in_label.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = self._service.users().messages().list(
                userId           = "me",
                labelIds         = [ labelId ],
                q                = None,
                pageToken        = page_token,
                maxResults       = None,
                includeSpamTrash = None
            ).execute()
            if 'messages' in response:
                all_message_in_label.extend(response['messages'])

        return all_message_in_label


# =============================================================================
if __name__ == '__main__':
    gm = bckgmail(__file__)
    # Authenticate
    gm.getToken()
    # Get GMail service
    gm.getService()

    # Get all GMail labels
    gm.getGMailLabels()

    # Go through labels and drop not wanted
    labels = []
    all_labels = gm.getLabels()
    # These labels are skipped
    skiplabels = [
        'CHAT', 'SENT', 'IMPORTANT', 'TRASH', 'DRAFT', 'SPAM',
        'CATEGORY_FORUMS', 'CATEGORY_UPDATES', 'CATEGORY_PERSONAL', 'CATEGORY_PROMOTIONS',
        'CATEGORY_SOCIAL', 'STARRED', 'UNREAD', 'Calendar', 'Unwanted', 'Deleted Items', 'Junk E-mail', 'Junk'
    ]
    for label in all_labels:
        if label not in skiplabels:
            labels.append(label)
    labels.sort()

    # Get all GMail messages for each label
    # Note that if a message has been tagged with multiple labels, it will appear in each one of them
    for label in labels:
        # The label might have parents, as MyMails/Work/important.
        # Split it
        label_list = label.split('/')

        # Use the label as a path 
        label_path = os.path.join('GMAIL', *label_list) 
        gm.createDir(label_path)

        # Get all GMail messages for label
        msgs = gm.getAllMessages(label)
        if not msgs:
            print('No email messages found.')
        else:
            # Loop through all messages and save as eml
            for emails in msgs:
                gm.saveMessageAsEML(emails['id'], label_path)
