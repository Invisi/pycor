# TODO: Switch over to pydantic's env and python-dotenv

# Sets console output level to DEBUG, default is INFO.
DEBUG = False

# Sets folders containing subjects over which will be iterated
FOLDERS = ["test_folder\\folder 1", "test_folder\\folder 2"]

# Username
MAIL_USER = "root@example.com"
# Password
MAIL_PASS = "root"
# IMAP server
MAIL_IMAP = "imap.example.com"
# SMTP server
MAIL_SMTP = "smtp.example.com"
# Spoof From header
MAIL_FROM = "pycor@example.com"

# Disable mail sending
DISABLE_OUTGOING_MAIL = True

# Whether to flag downloaded mails as read for debugging purposes
MARK_MAILS_AS_READ = False

# Minute to run at (multiple of n). E.g.: 5 results in 8:05, 8:10, 8:15...
DELAY_SLEEP = 10

# Passphrase for corrector pws files (Base64-encoded Fernet key)
PSW_PASSPHRASE = "gC9VGy09lEk7zK1257Pzj5-mDPclX_FScqLC2RLObyU="

# Ignored folder names
FOLDER_IGNORE = ["_postprocessing"]

# Make excel visible during processing
SHOW_EXCEL = False

# Sentry DSN
SENTRY_DSN = None

# Healthcheck address (called after every run)
HEALTHCHECK_PING = None

# Where to send mails with "PROBLEM" in the subject
ADMIN_CONTACT = "root@example.com"

# Domains from which mails should be accepted (student@domain.tld)
ACCEPTED_DOMAINS = ["alumni.fh-aachen.de", "fh-aachen.de"]

# Downloads mails from gmail and renames attachments
# This should only be used as a last resort, takes corrector codename as key
MAIL_FORWARDS = {
    "1234_1": {
        "username": "",
        "password": "",
    }
}
