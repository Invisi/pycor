# Sets folders containing subjects over which will be iterated
FOLDERS = [
    'test\\folder 1',
    'test\\folder 2'
]

# Sets console output level, default is INFO. Disables flagging mails as read
DEBUG = False
# Disable mail sending
DISABLE_OUTGOING_MAIL = True
# Make excel visible during processing
SHOW_EXCEL = False

# Delay (in s) between mailbox logins
DELAY_MAILBOXES = 15

# Passphrase for corrector pws files
PSW_PASSPHRASE = 'password'

# Ignored folder names
FOLDER_IGNORE = ['_postprocessing', 'PyCor_documentation']
