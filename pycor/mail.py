import datetime
import email.errors
import email.header
import email.mime.multipart
import email.mime.text
import imaplib
import logging
import os
import smtplib
import time
import typing
from email.utils import formatdate
from pathlib import Path

from pycor import config, excel, utils

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class LoginException(BaseException):
    pass


def _encode_name(x: str) -> str:
    try:
        file_name, charset = email.header.decode_header(x)[0]
        if charset:
            return file_name.decode(charset)
        return str(file_name)
    except (AttributeError, KeyError, email.errors.MessageError):
        logging.exception("Failed to decode filename")
        return ""


def _filter(
    msg: email.message.Message,
) -> typing.Tuple[email.message.Message, typing.Union[bool, str]]:
    """
    Finds files and filters out non-xlsx, recursively
    """

    # Throw out strings
    if isinstance(msg, str):
        return typing.cast(email.message.Message, msg), False

    # This might be an S/MIME mail
    if msg.get_content_type() == "multipart/mixed":
        sub_payload: typing.List[email.message.Message] = msg.get_payload()

        # Return early, we don't care about non-lists
        if not isinstance(sub_payload, list):
            return msg, False

        # Try to find excel file in sub payload
        for potential_file in sub_payload:
            msg, name_or_not = _filter(potential_file)
            if name_or_not:
                return msg, name_or_not

        return msg, False

    # Non-excel files
    if (
        msg.get_content_type() != EXCEL_MIME
        and msg.get_content_type() != "application/octet-stream"
    ):
        # Outlook for iOS (e.g.) will attach the file as octet-stream
        return msg, False

    return msg, _encode_name(msg.get_filename()).lower().endswith(".xlsx")


def filter_files(msg: email.message.Message) -> typing.List[email.message.Message]:
    _files = []
    for _ in typing.cast(typing.Iterable[email.message.Message], msg.get_payload()):
        msg, name_or_not = _filter(_)
        if name_or_not:
            _files.append(msg)

    return _files


class Mail:
    def __init__(self):
        self.log = logging.getLogger("PyCor").getChild("Mail")

        self.username = config.MAIL_USER
        self.password = config.MAIL_PASS

        self.imap: typing.Optional[imaplib.IMAP4_SSL] = None
        self.smtp: typing.Optional[smtplib.SMTP] = None
        self.smtp_connected: typing.Optional[float] = None

        # Login
        self.imap_login()

        self.log.info("%s - Successfully logged in", self.username)

    def imap_login(self):
        try:
            self.imap = imaplib.IMAP4_SSL(config.MAIL_IMAP)
            self.imap.login(self.username, self.password)
            self.imap.select("INBOX")
        except (imaplib.IMAP4.error, ConnectionError, TimeoutError):
            self.log.exception("Failed to login to IMAP server.")
            raise LoginException

    def smtp_login(self):
        # Log out if logged in
        if self.smtp:
            self.logout(smtp_only=True)

        # Retry login a few times
        for _ in range(5):
            try:
                self.smtp = smtplib.SMTP(config.MAIL_SMTP, 587)
                self.smtp.ehlo()
                self.smtp.starttls()
                self.smtp.login(config.MAIL_USER, config.MAIL_PASS)
                self.smtp_connected = time.time()
                return
            except smtplib.SMTPException:
                self.log.exception("Failed to login to SMTP server.")
                time.sleep(30)

        raise LoginException

    def logout(self, smtp_only=False):
        if self.smtp:
            try:
                self.smtp.quit()
            except smtplib.SMTPServerDisconnected:
                # Ignore disconnect, that's kinda what we want to do anyway
                pass
        self.smtp_connected = None

        if self.imap and not smtp_only:
            self.imap.close()
            self.imap.logout()

    def check_inbox(
        self, valid_filenames: typing.Dict[str, excel.Corrector]
    ) -> typing.Optional[typing.List[typing.Dict]]:
        ret, message_str = self.imap.search(None, "(UNSEEN)")
        corr_files = []

        if ret == "OK":
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
                elif any(
                    student_email.endswith(f"@{domain}")
                    for domain in config.ACCEPTED_DOMAINS
                ):
                    # Forward mails to admin if subject contains "problem"
                    if (
                        msg["Subject"]
                        and "problem" in msg["Subject"].lower()
                        and config.ADMIN_CONTACT
                    ):
                        msg.replace_header(
                            "Subject", f"PyCor: {msg['Subject']} from {msg['From']}"
                        )
                        msg.replace_header(
                            "From", "PyCor <{}>".format(config.MAIL_FROM)
                        )
                        msg.replace_header("To", config.ADMIN_CONTACT)
                        msg.replace_header("Date", formatdate(localtime=True))
                        self.send(config.ADMIN_CONTACT, "", msg)
                        self.send(student_email, *Generator.problem_forwarded())
                        continue

                    possible_files = filter_files(msg)

                    if len(possible_files) != 1:
                        # No file, multiple files or invalid file. Notify student
                        self.log.warning(
                            "Student submitted %s files.", len(possible_files)
                        )
                        self.send(student_email, *Generator.invalid_attachment())
                        continue

                    file_name = _encode_name(possible_files[0].get_filename())
                    stripped_filename = file_name.lower().replace(".xlsx", "").strip()
                    subject_corrector = None
                    for valid_filename, corr in valid_filenames.items():
                        if valid_filename == stripped_filename:
                            subject_corrector = corr

                    if not subject_corrector:
                        # Unknown subject. Notify student
                        self.log.warning("Student submitted unknown subject.")

                        self.send(
                            student_email, *Generator.unknown_attachment(file_name)
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
    ) -> typing.Optional[Path]:
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
            return None

        # Save file in proper folder
        file_path = user_dir / "{}_{}.xlsx".format(
            datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H.%M.%S"),
            utils.random_string(),
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

        # Attach html content if it's not a forwarded mail
        if isinstance(content, str):
            msg.attach(email.mime.text.MIMEText(content, "html", "utf-8"))

        try:
            # Avoid reconnecting multiple times
            if (
                self.smtp_connected is None
                or time.time() - self.smtp_connected > 5 * 60  # Over 5 minutes
            ):
                self.smtp_login()

            # Retry sending the mail a few times
            # TODO: Hello potential successor, please move mails to a different worker thread or just use Celery.
            #       Message handling should not be in the main thread if every small thing
            #       might go down at random.
            for _ in range(6):
                try:
                    self.smtp.sendmail(config.MAIL_FROM, recipient, msg.as_bytes())
                    self.log.info("Sent mail to %s", recipient)
                    break
                except smtplib.SMTPServerDisconnected:
                    if _ < 5:
                        self.log.error("Failed to send email, will retry")
                        self.smtp_login()
                    else:
                        self.log.critical(
                            "SMTP server still did not respond correctly. Raising exception."
                        )
                        raise LoginException
        except LoginException:
            self.log.error("Failed to send mail to %s", recipient)
            raise
        else:
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

    def forward_mails(self):
        for code_name, details in getattr(config, "MAIL_FORWARDS", {}).items():
            self.log.info("Downloading %s mails", code_name)

            try:
                imap = imaplib.IMAP4_SSL("imap.gmail.com")
                imap.login(details["username"], details["password"])
                imap.select("INBOX")

                ret, message_str = imap.search(None, "(UNSEEN)")
                if ret == "OK":
                    message_ids = message_str[0].split()
                    for message_id in message_ids:
                        _, data = imap.fetch(message_id, "(RFC822)")

                        msg: email.message.Message = email.message_from_bytes(
                            data[0][1]
                        )

                        self.log.info(
                            "%s - Forwarding message from %s (%s)",
                            self.username,
                            msg["From"],
                            msg["Subject"],
                        )

                        student_email = email.utils.parseaddr(msg["From"])[1]

                        if (
                            any(
                                _ in student_email
                                for _ in ["noreply", "no-reply", "mailer-daemon"]
                            )
                            or student_email == self.username
                            or student_email.startswith(details["username"])
                        ):
                            # Ignore mailer-daemon, no-reply, or own account
                            continue
                        elif any(
                            student_email.endswith(f"@{domain}")
                            for domain in config.ACCEPTED_DOMAINS
                        ):
                            possible_files = filter_files(msg)
                            if len(possible_files) != 1:
                                # No file, multiple files or invalid file. Notify student
                                self.log.warning(
                                    "Student submitted %s files.", len(possible_files)
                                )
                                self.send(
                                    student_email, *Generator.invalid_attachment()
                                )
                                continue

                            # Set correct name
                            possible_files[0].set_param(
                                "filename",
                                f"{code_name}.xlsx",
                                header="content-disposition",
                            )
                            possible_files[0].set_param(
                                "name",
                                f"{code_name}.xlsx",
                                header="content-type",
                            )

                            new_msg = email.mime.multipart.MIMEMultipart("alternative")
                            new_msg["From"] = msg["From"]
                            new_msg["To"] = config.MAIL_FROM
                            new_msg["Subject"] = "Forwarded mail from {}".format(
                                details["username"]
                            )
                            new_msg["Date"] = msg["Date"]
                            new_msg.attach(possible_files[0])

                            # Save file to inbox
                            try:
                                self.imap.append(
                                    "INBOX",
                                    "",
                                    imaplib.Time2Internaldate(time.time()),
                                    str(new_msg).encode("utf-8"),
                                )
                            except imaplib.IMAP4.abort:
                                # Retry saving the mail
                                self.imap_login()
                                self.imap.append(
                                    "INBOX",
                                    "",
                                    imaplib.Time2Internaldate(time.time()),
                                    str(new_msg).encode("utf-8"),
                                )
                        else:
                            # Notify sender about wrong email address
                            self.log.debug("Wrong address")
                            self.send(student_email, *Generator.wrong_address())

            except (imaplib.IMAP4.error, ConnectionError):
                self.log.exception("Failed to login to forwarding IMAP server.")


class Generator:
    @staticmethod
    def wrong_address() -> typing.Tuple[str, str]:
        return (
            "Falscher Account!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    die Haus??bung ist <b>zwingend mit Ihrem FH-Aachen-Account ('*****@alumni.fh-aachen.de') einzusenden.
                    </b> Bitte senden Sie Ihre L??sungen erneut ein. Die erfolgte Abgabe wird nicht weiter verarbeitet 
                    und nicht gewertet.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def malformed_attachment() -> typing.Tuple[str, str]:
        return (
            "Ung??ltige Excel-Datei im Anhang!",
            """
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben eine scheinbar <b>ung??ltige</b> Excel-Datei eingereicht. Bitte stellen Sie sicher, dass 
                    Sie eine aktuelle Datei als Vorlage genutzt haben.<br>
                    Sollte dies bereits die aktuellste Version sein, wenden Sie sich bitte an Ihre(n) Dozenten/-in oder 
                    Professor/-in, oder senden Sie uns eine Mail mit dem Betreff 'Problem'. 
                    Wir nehmen schnellstm??glich Kontakt mit Ihnen auf.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def invalid_attachment() -> typing.Tuple[str, str]:
        return (
            "Keine/mehrere Excel-Datei im Anhang!",
            """
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>keine oder mehrere Excel-Datei eingereicht</b>. Das g??ltige Format ist .xlsx und 
                    lediglich einzelne Dateien k??nnen akzeptiert werden.<br>
                    Sollten Sie Schwierigkeiten mit dem Einreichen Ihrer L??sungen haben, senden Sie bitte eine Email 
                    mit dem Betreff 'Problem'. Wir nehmen schnellstm??glich Kontakt mit Ihnen auf. Allgemeine Fragen 
                    zur Bearbeitung der Haus??bung werden nicht beantwortet!
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def unknown_attachment(submitted_name: str) -> typing.Tuple[str, str]:
        return (
            "Unbekanntes Modul.",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben eine Excel-Datei eingereicht, die keinem Modul zuzuordnen war. M??glicherweise wurde die 
                    Datei beim Herunterladen umbenannt, oder die Abgabefrist ist bereits verstrichen. Bitte vergleichen 
                    Sie den Namen der eingeschickten Datei ({submitted_name}) mit Ihren Aufgabenblatt/der Datei auf 
                    Ilias. <br/>
                    Sollte der Fehler nicht ersichtlich sein, melden Sie sich bitte bei Ihrem/-r Professor/-in oder der 
                    PyCor-Administration.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def error_processing(corrector_title: str) -> typing.Tuple[str, str]:
        return (
            "Fehler bei der Verarbeitung!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Beim Verarbeiten Ihrer eingesendeten Datei sind Fehler aufgetreten. M??glicherweise ist sie defekt 
                    oder enth??lt nicht alle notwendingen Informationen wie beispielsweise die Matrikelnummer.
                    Sollte dies mehrmals auftreten, nehmen Sie bitte Kontakt mit der PyCor-Administration 
                    (bspw. per Mail mit "Problem" im Betreff an diese Adresse) oder mit Ihrem/-r Professor/-in auf.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_passed(
        corrector_title: str, exercises: typing.List[int], mat_num: int
    ) -> typing.Tuple[str, str]:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Teil(e) erfolgreich gel??st!  Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>Aufgabe(n) {exercise_no} erfolgreich gel??st</b>! Herzlichen Gl??ckwunsch! 
                    Bitte drucken Sie diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit Ihrer 
                    handschriftlichen L??sung vor.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
        """,
        )

    @staticmethod
    def exercise_blocked(
        corrector_title: str, exercises: typing.List[int], max_tries: int
    ) -> typing.Tuple[str, str]:
        exercise_no = str([x + 1 for x in exercises])
        return (
            f"Sie wurden gesperrt!",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>{max_tries} fehlerhafte L??sungen zu Aufgabe(n) {exercise_no} eingereicht</b> und 
                    wurden daher <b>vor??bergehend gesperrt</b>. Weitere L??sungsversuche werden nur <b>nach 
                    pers??nlicher(!) R??cksprache</b> mit Ihrem/-r Professor/-in oder den zust??ndigen HiWis zugelassen.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_congrats(corrector_title: str, mat_num: int) -> typing.Tuple[str, str]:
        return (
            f"Haus??bung vollst??ndig gel??st! Mat. Num.: {mat_num}",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Sie haben <b>die gesamte Haus??bung erfolgreich gel??st</b>! Herzlichen Gl??ckwunsch! Bitte drucken 
                    Sie diese Email aus und legen Sie sie zur Testatunterzeichnung zusammen mit den handschriftlichen 
                    L??sungen aller Aufgabenteile vor.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
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

    @staticmethod
    def exercise_erroneous(
        corrector_title: str, ignored: typing.List[int]
    ) -> typing.Tuple[str, str]:
        return (
            "Problem: Fehler bei der Auswertung",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    bei der Verarbeitung von {'Aufgaben' if len(ignored) > 1 else 'Aufgabe'} 
                    {', '.join([str(_) for _ in ignored])} ist ein Fehler aufgetreten. Bitte stellen Sie sicher, dass
                    die Nummerierung der Aufgaben nicht ver??ndert wurde.<br>
                    Sollte der Fehler nicht zu finden sein, antworten Sie bitte auf diese Mail oder wenden Sie sich an 
                    Ihre/-n Professor/-in.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def exercise_ignored(corrector_title: str) -> typing.Tuple[str, str]:
        return (
            f"Problem: Fehler bei der Auswertung von {corrector_title}",
            f"""
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    die eingesendete Datei wurde nicht korrigiert, da keine nicht bereits bestandene (oder blockierte) 
                    Aufgabe vollst??ndig angegeben wurde. Wenn Sie Teil-L??sungen einreichen m??chten, setzen Sie 
                    einfach alle anderen Werte der Aufgabe auf 0 o.??.<br>
                    Sollten weiterhin Probleme bestehen, antworten Sie bitte einfach auf diese Mail.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    <b>{corrector_title}</b> und PyCor
                </p>
            </html>
            """,
        )

    @staticmethod
    def problem_forwarded() -> typing.Tuple[str, str]:
        return (
            "Problemmeldung weitergeleitet",
            """
            <html>
                <p>
                    Liebe(r) Studierende(r),<br><br>
                    Ihre Problemmeldung wurde an die PyCor-Administration weitergeleitet, die sich schnellstm??glich 
                    zur??ckmelden wird.<br>
                    Die eingesandte Datei wird nicht korrigiert.
                </p>
                <p>
                    Mit freundlichen Gr????en<br>
                    PyCor
                </p>
            </html>
            """,
        )
