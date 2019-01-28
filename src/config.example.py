# Sets console output level, default is INFO.
DEBUG = False

# Sets folders containing subjects over which will be iterated
FOLDERS = [
    'test\\folder 1',
    'test\\folder 2'
]

# Username
MAIL_USER = "root@localhost"
# Password
MAIL_PASS = "root"
# IMAP server
MAIL_IMAP = "mail.fh-aachen.de"
# SMTP server
MAIL_SMTP = "mail.fh-aachen.de:465"

# Disable mail sending
DISABLE_OUTGOING_MAIL = True

# Whether to flag mails as read
MARK_MAILS_AS_READ = False

# Delay (in s) between mailbox logins
DELAY_MAILBOXES = 15

# Delay (in m) between main loop
DELAY_SLEEP = 2

# Passphrase for corrector pws files (Base64)
PSW_PASSPHRASE = 'gC9VGy09lEk7zK1257Pzj5-mDPclX_FScqLC2RLObyU='

# Ignored folder names
FOLDER_IGNORE = ['_postprocessing', 'PyCor_documentation']

# Make excel visible during processing
SHOW_EXCEL = False

# Sentry DSN
SENTRY_DSN = None
