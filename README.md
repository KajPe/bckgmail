# bckgmail
Downloads GMail messages to local drive as eml files with Google Service API

No IMAP needed!

This is a working code, but more of a proof-of-concept for now.

Note! You have to get your own oauth file from Google console. Store in same folder as the script is as credentials.json.
During first run it should open the browser and ask to give permissions to the script. This will store a token.pickle file.
All is done according to Google documentation and this authentication sequence can be found under class function getToken().

All of your INBOX and any folder is downloaded under a GMAIL folder where the script is running.
eml files can be opened with Outlook, Thunderbird, etc. Also they can be massimported with various tools.
