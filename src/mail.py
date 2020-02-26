import datetime
import email
import email.mime.multipart
import email.mime.text
import imaplib
import logging
import os
import smtplib
import time
from email.utils import formatdate
from pathlib import Path
from typing import Optional, List

import typing

import excel
import utils

config = utils.import_config()
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class LoginException(BaseException):
    pass


class Mail:
    def __init__(self):
        self.log = logging.getLogger("PyCor").getChild("Mail")

        self.username = config.MAIL_USER
        self.password = config.MAIL_PASS

        self.imap = None
        self.smtp = None

        # Login
        self.imap_login()

        self.log.info("%s - Successfully logged in", self.username)

    def imap_login(self):
        try:
            self.imap = imaplib.IMAP4_SSL(config.MAIL_IMAP)
            self.imap.login(self.username, self.password)
            self.imap.select("INBOX")
        except (imaplib.IMAP4.error, ConnectionError):
            self.log.exception("Failed to login to IMAP server.")
            raise LoginException

    def smtp_login(self):
        try:
            self.smtp = smtplib.SMTP(config.MAIL_SMTP, 587)
            self.smtp.ehlo()
            self.smtp.starttls()
            self.smtp.login(config.MAIL_USER, config.MAIL_PASS)
        except smtplib.SMTPException:
            self.log.exception("Failed to login to SMTP server.")
            raise LoginException


    def smtp_logout(self):
        if self.smtp:
            try:
                self.smtp.quit()
            except smtplib.SMTPServerDisconnected:
                # Ignore disconnect, that's kinda what we want to do anyway
                pass

    def check_inbox(self, valid_filenames: dict) -> Optional[List[dict]]:
        ret, message_str = self.imap.search(None, "(UNSEEN)")

        if ret == "OK":
            corr_files = []
            message_ids = message_str[0].split()
            for message_id in message_ids:
                # In theory this could fail IF someone deletes the message before it is fetched.
                # This should just result in an empty mail however.
                _, data = self.imap.fetch(message_id, "(RFC822)")

                # Keep mail as unread if in debug mode
                if not (
                    hasattr(config, "MARK_MAILS_AS_READ") and config.MARK_MAILS_AS_READ
                ):
                    self.imap.store(message_id, "-FLAGS", "\\Seen")

                msg: email.message.Message = email.message_from_bytes(data[0][1])

                self.log.info(
                    "%s - Downloading message from %s (%s)",
                    self.username,
                    msg["From"],
                    msg["Subject"],
                )

                # Forward mails to admin if subject contains "problem"
                if (
                    msg["Subject"]
                    and "problem" in msg["Subject"].lower()
                    and config.ADMIN_CONTACT
                ):
                    msg.replace_header(
                        "Subject", f"PyCor: {msg['Subject']} from {msg['From']}"
                    )
                    msg.replace_header("From", "PyCor <{}>".format(config.MAIL_FROM))
                    msg.replace_header("To", config.ADMIN_CONTACT)
                    msg.replace_header("Date", formatdate(localtime=True))
                    self.send(config.ADMIN_CONTACT, "", msg)
                    continue

                student_email = email.utils.parseaddr(msg["From"])[1]

                if (
                    any(
                        _ in student_email
                        for _ in ["noreply", "no-reply", "mailer-daemon"]
                    )
                    or student_email == self.username
                ):
                    # Ignore mailer-daemon, no-reply, or own account
                    continue
                elif student_email.endswith("fh-aachen.de"):
                    # Find files
                    possible_files = [
                        _
                        for _ in msg.get_payload()
                        if not isinstance(_, str) and _.get_content_type() == EXCEL_MIME
                    ]
                    if len(possible_files) != 1:
                        # No file, multiple files or invalid file. Notify student
                        self.log.warn("Student submitted multiple files.")
                        self.send(student_email, *Generator.invalid_attachment())
                        continue

                    stripped_filename = possible_files[0].get_filename().lower().strip()
                    subject_corrector = None
                    for valid_filename, corr in valid_filenames.items():
                        if f"{valid_filename}.xlsx" == stripped_filename:
                            subject_corrector = corr

                    if not subject_corrector:
                        # Unknown subject. Notify student
                        self.log.warn("Student submitted unknown subject.")

                        filename = possible_files[0].get_filename()
                        if "=?" in filename:
                            try:
                                header = email.header.decode_header(filename)
                                if len(header) > 0:
                                    content, encoding = header[0]
                                    filename = content.decode(encoding)
                            except email.errors.HeaderParseError:
                                self.log.error("Failed to parse header for filename")

                        self.send(
                            student_email, *Generator.unknown_attachment(filename)
                        )
                        continue

                    downloaded_file = self.download_attachment(
                        possible_files[0], student_email, subject_corrector
                    )

                    if downloaded_file:
                        corr_files.append(
                            {"student": downloaded_file, "corrector": subject_corrector}
                        )
                        self.log.info("Accepted submitted file")

                else:
                    # Notify sender about wrong email address
                    self.log.debug("Wrong address")
                    self.send(student_email, *Generator.wrong_address())
            return corr_files

    def download_attachment(
        self, _file: email.message.Message, student_email: str, subject: excel.Corrector
    ) -> Optional[Path]:
        """
        Downloads attachment and returns the full path to it.
        If there are multiple or no attachments None is returned.

        :param _file: Message
        :param student_email: Student's email address
        :param subject: :class:`excel.Corrector` instance that contains necessary paths
        :return: Full path to downloaded file OR None
        """

        user_dir = subject.parent_path / student_email

        # Create folder
        try:
            user_dir.mkdir(exist_ok=True)
        except OSError:
            self.log.exception("Failed to create user folder.")
            return

        # Save file in proper folder
        basename, extension = os.path.splitext(_file.get_filename())

        file_path = user_dir / "{}_{}{}".format(
            datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H.%M.%S"),
            utils.random_string(),
            extension,
        )

        with file_path.open("wb") as fp:
            fp.write(_file.get_payload(decode=True))

        self.log.debug("Saved file to %s", os.sep.join(file_path.parts[-3:]))

        return file_path

    def send(
        self,
        recipient: str,
        subject: str,
        content: typing.Union[str, email.message.Message],
    ):
        # Mail is forwarded
        if isinstance(content, email.message.Message):
            msg = content
        else:
            # Mail should be generated
            msg = email.mime.multipart.MIMEMultipart("alternative")
            msg["From"] = "PyCor <{}>".format(config.MAIL_FROM)
            msg["To"] = recipient
            msg["Subject"] = subject
            msg["Date"] = formatdate(localtime=True)

        # Don't send emails if in debug mode
        if hasattr(config, "DISABLE_OUTGOING_MAIL") and config.DISABLE_OUTGOING_MAIL:
            self.log.debug("Sending mail: %s", content)
            return
        try:
            self.smtp_login()

            # Attach html content if it's not a forwarded mail
            if isinstance(content, str):
                msg.attach(email.mime.text.MIMEText(content, "html", "utf-8"))

            self.smtp.sendmail(config.MAIL_FROM, recipient, msg.as_bytes())
            self.smtp_logout()
            self.log.info("Sent mail to %s", recipient)

            # Save mail to Sent
            try:
                self.imap.append(
                    "Sent",
                    "\\Seen",
                    imaplib.Time2Internaldate(time.time()),
                    str(msg).encode("utf-8"),
                )
            except imaplib.IMAP4.abort:
                # Retry saving the mail
                self.imap_login()
                self.imap.append(
                    "Sent",
                    "\\Seen",
                    imaplib.Time2Internaldate(time.time()),
                    str(msg).encode("utf-8"),
                )

        except LoginException:
            self.log.error("Failed to send mail to %s", recipient)
            raise


class Generator:
    @staticmethod
    def wrong_address() -> tuple:
        return (
            "Falscher Account!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    die Hausübung ist <b>zwingend mit Ihrem FH-Aachen-Account ('*****@alumni.fh-aachen.de') einzusenden.
                    </b> Bitte senden Sie Ihre Lösungen erneut ein. Die erfolgte Abgabe wird nicht weiter verarbeitet 
                    und nicht gewertet.
                </p>
                <p>
                    Mit freundlichen Grüßen<br>
                    PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def malformed_attachment() -> tuple:
        return (
            "Ungültige Excel-Datei im Anhang!",
            """
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben eine scheinbar <b>ungültige</b> Excel-Datei eingereicht. Bitte stellen Sie sicher, dass 
                    Sie eine aktuelle Datei als Vorlage genutzt haben.<br>
                    Sollte dies bereits die aktuellste Version sein, wenden Sie sich bitte an Ihre(n) Dozenten/-in oder 
                    Professor/-in, oder senden Sie uns eine Mail mit dem Betreff 'Problem'. 
                    Wir nehmen schnellstmöglich Kontakt mit Ihnen auf.
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
                    Sie haben <b>keine Excel-Datei eingereicht</b>. Das gültige Format ist .xlsx.<br>
                    Sollten Sie Schwierigkeiten mit dem Einreichen Ihrer Lösungen haben, senden Sie bitte eine Email 
                    mit dem Betreff 'Problem'. Wir nehmen schnellstmöglich Kontakt mit Ihnen auf. Allgemeine Fragen 
                    zur Bearbeitung der Hausübung werden nicht beantwortet!
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
                    Sie haben eine Excel-Datei eingereicht, die keinem Modul zuzuordnen war. Möglicherweise wurde die 
                    Datei beim Herunterladen umbenannt. Bitte vergleichen Sie den Namen der eingeschickten Datei 
                    ({submitted_name}) mit Ihren Aufgabenblatt/der Datei auf Ilias.
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
    def error_processing(corrector_title: str) -> tuple:
        return (
            "Fehler bei der Verarbeitung!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Beim Verarbeiten Ihrer eingesendeten Datei sind Fehler aufgetreten. Möglicherweise ist sie defekt 
                    oder enthält nicht alle notwendingen Informationen wie beispielsweise die Matrikelnummer.
                    Sollte dies mehrmals auftreten, nehmen Sie bitte Kontakt mit Ihrem Professor oder der 
                    PyCor-Administration auf.
                </p>
                <p>
                    Mit freundlichen Grüßen<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_passed(
        corrector_title: str, exercises: List[int], mat_num: int
    ) -> tuple:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Teil(e) erfolgreich gelöst!  Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>Aufgabe(n) {exercise_no} erfolgreich gelöst</b>! Herzlichen Glückwunsch! 
                    Bitte drucken Sie diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit Ihrer 
                    handschriftlichen Lösung vor.
                </p>
                <p>
                    Mit freundlichen Grüßen<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def exercise_blocked(
        corrector_title: str, exercises: List[int], max_tries: int
    ) -> tuple:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Sie wurden gesperrt!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>{max_tries} fehlerhafte Lösungen zu Aufgabe(n) {exercise_no} eingereicht</b> und 
                    wurden daher <b>vorübergehend gesperrt</b>. Weitere Lösungsversuche werden nur <b>nach 
                    persönlicher(!) Rücksprache</b> mit Ihrem Professor/den zuständigen HiWis zugelassen.
                </p>
                <p>
                    Mit freundlichen Grüßen<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_congrats(corrector_title: str, mat_num: int) -> tuple:
        return (
            f"Hausübung vollständig gelöst! Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>die gesamte Hausübung erfolgreich gelöst</b>! Herzlichen Glückwunsch! Bitte drucken 
                    Sie diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit den handschriftlichen 
                    Lösungen aller Aufgabenteile vor.
                </p>
                <p>
                    Mit freundlichen Grüßen<br>
                    <b>{corrector_title}</b> und PyCor
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
                <table><tr><th>Variable</th><th>Korrektur</th></tr>{ret}</table>
                <hr/>
                """
