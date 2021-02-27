import argparse
import datetime
import getpass
import os
import time
from pathlib import Path

from cryptography.fernet import Fernet

from . import config
from . import utils

__version__ = "2021-02-27"

from pycor import log, main

if __name__ == "__main__":
    # Ignore warnings if not in debug mode
    if not config.DEBUG:
        os.environ["PYTHONWARNINGS"] = "ignore"

    parser = argparse.ArgumentParser(
        description="PyCor - Python Correction",
        epilog="Runs normally when no argument is supplied",
    )
    parser.add_argument("-p", "--psw", action="store_true", help="Create password file")
    parser.add_argument(
        "-s",
        "--secret",
        action="store_true",
        help="Generate password passphrase/secret",
    )

    args = parser.parse_args()
    if args.psw:  # Create password file
        psw = Path("psw")
        pw = getpass.getpass("Enter password (will not be echoed): ")
        pw2 = getpass.getpass("Enter password again: ")

        if pw != pw2 or not pw:
            print("Passwords do not match or empty password supplied!")
            exit(1)
        else:
            psw.write_bytes(Fernet(config.PSW_PASSPHRASE).encrypt(bytes(pw, "utf-8")))
            exit()
    elif args.secret:
        key = Fernet.generate_key()
        print("Please set this passphrase in config.py: {}".format(key.decode("utf-8")))
        exit()

    # Initialize Sentry
    if hasattr(config, "SENTRY_DSN") and config.SENTRY_DSN:
        utils.setup_sentry(__version__)

    log.info("Welcome to PyCor v.%s", __version__)
    log.info("PyCor is now running!")

    while True:
        main()

        # Wait until a multiple of DELAY_SLEEP is on the clock
        current = datetime.datetime.now()
        current_time = current.time()
        sleep_time = (
            abs(current_time.minute % config.DELAY_SLEEP - config.DELAY_SLEEP) * 60
            - current_time.second
        )
        next_execution = current + datetime.timedelta(seconds=sleep_time)
        log.info("Pausing until %s", next_execution.strftime("%H:%M:%S"))
        time.sleep(sleep_time)
