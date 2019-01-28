import datetime
import email
import email.mime.multipart
import email.mime.text
import imaplib
import logging
import os
import smtplib
import traceback
from pathlib import Path
from typing import Optional, List

import utils

config = utils.import_config()


def choose_server(usn):
    """Given an e-mail 'usn', detects the e-mail termination and provides with the
    correct server information (both IMAP and SMTP). Available options are:
    gmail and FH Aachen e-mail accounts. It is missing the information for FH
    Aachen SMTP server."""

    domain = usn.split('@')[-1]

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
    def __init__(self, username: str, password: str, subject_folder: Path, subject_name: str):
        self.log = logging.getLogger('PyCor').getChild('Mail')

        self.username = username
        self.password = password
        self.subject_folder = subject_folder
        self.subject_name = subject_name

        self.imap = None
        self.smtp = None

        # Login
        self.imap_server, self.smtp_server = choose_server(self.username)
        try:
            self.imap = imaplib.IMAP4_SSL(self.imap_server)
            self.imap.login(self.username, self.password)
            self.imap.select('INBOX')
        except imaplib.IMAP4.error:
            utils.write_error(self.subject_folder,
                              'Fehler beim Zugriff auf das E-Mail-Konto: \n' + traceback.format_exc())
            self.log.error(traceback.format_exc())
            raise LoginException

        self.log.info('%s - Successfully logged in', username)

    def smtp_login(self):
        try:
            self.smtp = smtplib.SMTP_SSL(self.smtp_server)
            self.smtp.login(self.username, self.password)
        except smtplib.SMTPException:
            self.log.error(traceback.format_exc())
            raise LoginException

    def smtp_logout(self):
        if self.smtp:
            self.smtp.quit()

    def check_inbox(self) -> Optional[list]:
        ret, message_str = self.imap.search(None, '(UNSEEN)')

        if ret == 'OK':
            corr_files = []
            message_ids = message_str[0].split()
            for message_id in message_ids:
                # In theory this could fail IF someone deletes the message before it is fetched.
                # This should just result in an empty mail however.
                _, data = self.imap.fetch(message_id, '(RFC822)')

                # Keep mail as unread if in debug mode
                if config.DEBUG:
                    self.imap.store(message_id, '-FLAGS', '\\Seen')

                msg = email.message_from_bytes(data[0][1])

                self.log.info('%s - Downloading message %s from %s (%s)', self.username, message_id.decode('utf-8'),
                              msg['From'], msg['Subject'])

                # Ignore mailer daemon or no-reply
                student_email = email.utils.parseaddr(msg['From'])[1]
                if any(x in student_email for x in ['noreply', 'no-reply', 'mailer-daemon']):
                    continue
                elif student_email.endswith('fh-aachen.de'):
                    downloaded_file = self.download_attachment(msg, student_email)

                    if downloaded_file:
                        corr_files.append(downloaded_file)

                    # No file, multiple files or invalid file
                    if not downloaded_file:
                        # Notify student about missing file
                        self.send(student_email, *Generator.invalid_attachment(self.subject_name))
                        pass
                else:
                    # Notify sender about wrong email address
                    self.send(student_email, *Generator.wrong_address(self.subject_name))
                    pass
            if len(message_ids) == 0:
                self.log.info('%s - No new mail', self.username)

            return corr_files

    def download_attachment(self, msg: email.message.Message, student_email: str) -> Optional[Path]:
        """
        Downloads attachment and returns the full path to it.
        If there are multiple or no attachments None is returned.

        :param msg: Message
        :param student_email: Student's email address
        :return: Full path to downloaded file OR None
        """
        user_dir = self.subject_folder / student_email

        # Create folder
        try:
            user_dir.mkdir(exist_ok=True)
        except OSError:
            self.log.error('Failed to create user folder %s', user_dir)
            self.log.error(traceback.format_exc())
            return

        '''
        Valid Mime Types:
        xls: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        xlsx: application/vnd.ms-excel = 
        xlsm: application/vnd.ms-excel.sheet.macroenabled.12
        '''
        # Get valid excel attachments
        valid_files = []
        for attachment in msg.get_payload():
            if attachment.get_content_type().startswith('application/vnd'):
                filename = attachment.get_filename()
                basename, extension = os.path.splitext(filename)
                if extension not in ['.xls', '.xlsx', '.xlsm']:
                    continue
                valid_files.append(attachment)

        if len(valid_files) == 1:
            # Finally save file in proper folder
            valid_file = valid_files[0]
            basename, extension = os.path.splitext(valid_file.get_filename())

            # TODO: Use msg['Date'] if available? Might falsify real submit date/time
            dt = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d %H.%M.%S')
            file_path = user_dir / '{}_{}{}'.format(dt, utils.random_string(), extension)

            with file_path.open('wb') as fp:
                fp.write(valid_file.get_payload(decode=True))

            self.log.debug('Saved file to %s', os.sep.join(file_path.parts[-3:]))

            return file_path
        else:
            self.log.warning('Got more than one or no attachment(s) at all!')
        return None

    def send(self, recipient: str, subject: str, content: str):
        msg = email.mime.multipart.MIMEMultipart('alternative')
        msg['From'] = 'PyCor <{}>'.format(self.username)
        msg['To'] = recipient
        msg['Subject'] = subject

        # Don't send emails if in debug mode
        if config.DISABLE_OUTGOING_MAIL:
            self.log.debug('sending mail: %s', content)
            return
        try:
            self.smtp_login()
            msg.attach(email.mime.text.MIMEText(content, 'html', 'utf-8'))
            self.smtp.sendmail(self.username, recipient, msg.as_bytes())
            self.smtp_logout()
        except LoginException:
            self.log.error('Failed to send mail to %s', recipient)
            raise


class Generator:
    @staticmethod
    def wrong_address() -> tuple:
        return (
            "Falscher Account!",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br> <br>
                die Hausübung ist <b>zwingend mit Ihrem FH-Aachen-Account ('*****@alumni.fh-aachen.de') einzusenden</b>.
                Bitte senden Sie Ihre Lösungen erneut ein. Die erfolgte Abgabe wird nicht weiter verabeitet und nicht 
                gewertet.
                </p>
            
                <p>
                Mit freundlichen Grüßen<br>
                PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def invalid_attachment() -> tuple:
        return (
            "Keine Excel-Datei im Anhang!",
            """
            <html>
                <p>
                Liebe(r) Studierende(r),<br><br>
                Sie haben <b>keine Excel-Datei eingereicht</b>. Das gültige Format ist .xlsx. <br>
                Sollten Sie Schwierigkeiten mit dem Einreichen Ihrer Lösungen haben, senden Sie bitte eine Email mit dem
                Betreff 'PROBLEM'. Wir nehmen schnellstmöglich Kontakt mit Ihnen auf. Allgemeine Fragen zur Bearbeitung 
                der Hausübung werden nicht beantwortet!
                </p>
            
                <p>
                Mit freundlichen Grüßen<br>
                PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def unknown_attachment(submitted_name: str) -> tuple:
        return (
            "Unbekanntes Modul.",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br><br>
                Sie haben eine Excel-Datei eingereicht, die keinem Modul zuzuordnen war. Möglicherweise wurde die Datei 
                beim Herunterladen umbenannt. Bitte vergleichen Sie den Namen der eingeschickten Datei ({submitted_name}) 
                mit Ihren Aufgabenblatt/der Datei auf Ilias.
                Sollte der Fehler nicht ersichtlich sein, melden Sie sich bitte bei Ihren Professor oder der 
                PyCor-Administration.
                </p>
    
                <p>
                Mit freundlichen Grüßen<br>
                PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def error_processing(subject_name: str) -> tuple:
        return (
            "Fehler bei der Verarbeitung!",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br><br>
                Beim Verarbeiten Ihrer eingesendeten Datei sind Fehler aufgetreten. Möglicherweise ist sie defekt oder 
                enthält nicht alle notwendingen Informationen wie beispielsweise die Matrikelnummer.
                Sollte dies mehrmals auftreten, nehmen Sie bitte Kontakt mit Ihrem Professor oder der 
                PyCor-Administration auf.
                </p>

                <p>
                Mit freundlichen Grüßen<br>
                <b>{subject_name}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_passed(subject_name: str, exercises: List[int], mat_num: int) -> tuple:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Teil(e) erfolgreich gelöst!  Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br> <br>
                Sie haben <b>Aufgabe(n) {exercise_no} erfolgreich gelöst</b>! Herzlichen Glückwunsch! 
                Bitte drucken Sie diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit Ihrer 
                handschriftlichen Lösung vor.
                </p>

                <p>
                Mit freundlichen Grüßen<br>
                <b>{subject_name}</b> und PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def exercise_blocked(
        subject_name: str, exercises: List[int], max_tries: int
    ) -> tuple:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Sie wurden gesperrt!",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br> <br>
                Sie haben <b>{max_tries} fehlerhafte Lösungen zu Aufgabe(n) {exercise_no} eingereicht</b> und wurden 
                daher <b> vorübergehend gesperrt</b>. Weitere Lösungesversuche werden nur <b>nach persönlicher(!) 
                Rücksprache</b> mit Ihrem Professor/den zuständigen HiWis zugelassen.
                </p>

                <p>
                Mit freundlichen Grüßen<br>
                <b>{subject_name}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_congrats(subject_name: str, mat_num: int) -> tuple:
        return (
            f"Hausübung vollständig gelöst! Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                Liebe(r) Studierende(r),<br> <br>
                Sie haben <b>die gesamte Hausübung erfolgreich gelöst</b>! Herzlichen Glückwunsch! Bitte drucken Sie 
                diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit den handschriftlichen Lösungen 
                aller Aufgabenteile vor.
                </p>
            
                <p>
                Mit freundlichen Grüßen<br>
                <b>{subject_name}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_details(solution: dict) -> str:
        ret = ""
        for idx, var in enumerate(solution["var_names"]):
            solved = solution["correct"][idx]
            ret += (
                f"<tr><td>{var}</td><td>"
                f'<font color="{"green" if solved else "red"}">{"Richtig" if solved else "Falsch"}</font>'
                "</td></tr>"
            )

        return f"""
                Detaillierte Ergebnisse: <b>Teilaufgabe {solution['exercise'] + 1}</b><br/><br/>
                <table>
                    <tr>
                        <th>Variable</th>
                        <th>Korrektur</th>
                    </tr>
                    {ret}
                </table>
                <hr/>
                """
