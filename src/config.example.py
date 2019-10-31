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
MAIL_SMTP = "relay.fh-aachen.de"
# Spoof From header
MAIL_FROM = "uebungen.bau@fh-aachen.de"

# Disable mail sending
DISABLE_OUTGOING_MAIL = True

# Whether to flag mails as read
MARK_MAILS_AS_READ = False

# Minute to run at (multiple of n). E.g.: 5 results in 8:05, 8:10, 8:15...
DELAY_SLEEP = 10

# Passphrase for corrector pws files (Base64)
PSW_PASSPHRASE = 'gC9VGy09lEk7zK1257Pzj5-mDPclX_FScqLC2RLObyU='

# Ignored folder names
FOLDER_IGNORE = ['_postprocessing', 'PyCor_documentation']

# Make excel visible during processing
SHOW_EXCEL = False

# Sentry DSN
SENTRY_DSN = None

# Where to send mails with "PROBLEM" in the subject
ADMIN_CONTACT = "uebungen.bau@fh-aachen.de"
