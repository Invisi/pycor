# -*- coding: utf-8 -*-
import datetime
import logging
import logging.handlers
import random
import string
import sys
import traceback
from pathlib import Path

import sentry_sdk  # type: ignore


def setup_logger(level=logging.DEBUG):
    # Create logs folder
    if not Path("logs").exists():
        Path("logs").mkdir()

    log = logging.getLogger("PyCor")
    fmt = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
    )

    def crash_handler(exc_type, value, tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, value, tb)
            return
        log.critical("Uncaught exception", exc_info=(exc_type, value, tb))
        input("Press return to exit.")
        sys.exit(1)

    sys.excepthook = crash_handler

    hldr = logging.handlers.TimedRotatingFileHandler(
        "logs/PyCor.log", when="W0", encoding="utf-8", backupCount=16
    )
    hldr.setFormatter(fmt)
    hldr.setLevel(logging.DEBUG)
    log.addHandler(hldr)

    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(fmt)
    stream.setLevel(level)
    log.setLevel(logging.DEBUG)
    log.addHandler(stream)
    return log


def write_ignore(subject_folder: Path, message: str):
    with (subject_folder / "PYCOR_IGNORE.txt").open("a") as e:
        e.write(message)


def write_error(subject_folder: Path, message: str):
    """
    Writes specified error message to PYCOR_ERROR.txt in {subject_folder}.
    Also creates PYCOR_IGNORE.txt to stop PyCor from running against a brick wall every few minutes.

    :param subject_folder: Path to subject folder
    :param message: Error message to write
    :return:
    """
    dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    error_file = subject_folder / "PYCOR_ERROR.txt"

    # If there's an exception append it
    if not all([_ is None for _ in sys.exc_info()]):
        message += "\n" + traceback.format_exc()

    with error_file.open("a") as e:
        e.write("{} - {}\n".format(dt, message))

    # Ignore folder after this
    write_ignore(
        subject_folder,
        "{} - Löschen Sie diese Datei, sobald der Fehler in PYCOR_ERROR.txt behoben wurde.\n".format(
            dt
        ),
    )


def import_config():
    # Import config in exe created by PyInstaller
    if getattr(sys, "frozen", False):
        import importlib.util

        config_file = Path(sys.executable).parent / "config.py"
        if not config_file.exists():
            if not (Path(sys.executable).parent / "config.example.py").exists():
                from shutil import copyfile

                copyfile(
                    Path(sys._MEIPASS) / "config.example.py",
                    Path(sys.executable).parent / "config.example.py",
                )

            print("config.py is missing! An example config was created.")
            input("Press return to exit.")
            sys.exit(1)

        spec = importlib.util.spec_from_file_location("config", config_file)
        config = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(config)
    else:
        import config  # type: ignore
    return config


def setup_sentry(release):
    config = import_config()
    if config.DISABLE_OUTGOING_MAIL:
        environment = "dev"
    else:
        environment = "prod"

    def before_send(event, hint):
        if "exc_info" in hint:
            _, exc_value, _ = hint["exc_info"]
            if isinstance(exc_value, KeyboardInterrupt):
                return None
        return event

    sentry_sdk.init(
        dsn=config.SENTRY_DSN,
        release=release,
        environment=environment,
        send_default_pii=True,
        before_send=before_send,
        ca_certs=str(Path(__file__).parent / "cacert.pem"),
    )


def random_string():
    return "".join(random.choices(string.digits + string.ascii_letters, k=6))
