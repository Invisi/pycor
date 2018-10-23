import imaplib
import logging
import traceback


def choose_server(usn):
    """Given an e-mail 'usn', detects the e-mail termination and provides with the
    correct server information (both IMAP and SMTP). Available options are:
    gmail and FH Aachen e-mail accounts. It is missing the information for FH
    Aachen SMTP server."""

    domain = usn.split('@')[-1]
    imap = None
    smtp = None

    if domain == 'gmail.com':
        imap = 'imap.gmail.com'
        smtp = 'smtp.gmail.com:465'
    elif domain == 'ad.fh-aachen.de':
        # This server should be selected from FH info
        imap = 'mail.fh-aachen.de'
        smtp = 'mail.fh-aachen.de:587'
    elif domain == '0x0f.net':
        imap = 'mail.0x0f.net'
        smtp = 'mail.0x0f.net:465'
    else:
        raise Exception('Could not detect mail server')

    return imap, smtp


class LoginException(BaseException):
    pass


class Mail:
    def __init__(self, username, password):
        self.log = logging.getLogger('PyCor').getChild('Mail')

        self.username = username
        self.password = password

        self.con = None
        # Login
        imap, smtp = choose_server(self.username)
        try:
            self.con = imaplib.IMAP4_SSL(imap)
            self.con.login(self.username, self.password)
        except imaplib.IMAP4.error:
            self.log.error(traceback.format_exc())
            raise LoginException

        self.log.info('Successfully logged in to %s', username)

    def logout(self):
        if self.con and self.con.state == 'SELECTED':
            self.con.close()
        if self.con:
            self.con.logout()

    def download_new_mail(self):
        return {}
