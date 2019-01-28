# -*- coding: utf-8 -*-
import datetime
import logging
import logging.handlers
import random
import string
import sys
import traceback
from pathlib import Path

import sentry_sdk

import config


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
    with (subject_folder / "PYCOR_IGNORE.txt").open("a") as e:
        e.write(
            "{} - Delete this file once the error specified in PYCOR_ERROR.txt is fixed.\n".format(
                dt
            )
        )


def import_config():
    # Import config in exe created by PyInstaller
    if getattr(sys, "frozen", False):
        import importlib.util

        config_file = Path(sys.executable).parent / "config.py"
        if not config_file.exists():
            print("config.py is missing!")
            input("Press return to exit.")
            sys.exit(1)

        spec = importlib.util.spec_from_file_location("config", config_file)
        config = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(config)
    else:
        import config
    return config


def setup_sentry(release):
    if config.DISABLE_OUTGOING_MAIL:
        environment = "dev"
    else:
        environment = "prod"
    sentry_sdk.init(
        dsn=config.SENTRY_DSN,
        release=release,
        environment=environment,
        ca_certs=str(Path(__file__).parent / "cacert.pem"),
    )


def random_string():
    return "".join(random.choices(string.digits + string.ascii_letters, k=6))
