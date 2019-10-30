import argparse
import getpass
import logging
import os
import time
from pathlib import Path
from sys import exit

from cryptography.fernet import Fernet

import excel
import mail
import post
import utils

__version__ = "2019-10-30"


def switch_tolerance(lower_tolerance, higher_tolerance):
    if lower_tolerance > higher_tolerance:
        t = higher_tolerance
        higher_tolerance = lower_tolerance
        lower_tolerance = t

    return lower_tolerance, higher_tolerance


def compare(
    attempt: float or str,
    solution: float or str,
    tolerance_rel: int or float or None,
    tolerance_abs: int or float or None,
):
    """
    Compares student's attempt to solution and returns True if the attempt is within either tolerance margin

    :param attempt:
    :param solution:
    :param tolerance_rel:
    :param tolerance_abs:
    :return:
    """
    try:
        # In case the student made a space after the comma...
        if isinstance(attempt, str) and isinstance(solution, float):
            attempt = float(attempt.replace(",", ".").replace(" ", ""))

        # Compare string values
        if isinstance(solution, str):
            return solution.lower().strip() == attempt.lower().strip()

        # Solution was empty
        if solution is None:
            return True

        log.debug(
            "Comparing %s (%s) to %s (%s) with rel:%s abs:%s",
            attempt,
            type(attempt),
            solution,
            type(solution),
            tolerance_rel,
            tolerance_abs,
        )

        # Compare numerical values
        if isinstance(solution, float):
            absolute, relative = False, False
            if tolerance_rel is not None:
                lower_tolerance, higher_tolerance = switch_tolerance(
                    (1 - tolerance_rel / 100.0) * solution,
                    (1 + tolerance_rel / 100.0) * solution,
                )

                relative = lower_tolerance <= attempt <= higher_tolerance
            if tolerance_abs is not None:
                lower_tolerance, higher_tolerance = switch_tolerance(
                    solution - tolerance_abs, solution + tolerance_abs
                )
                absolute = lower_tolerance <= attempt <= higher_tolerance
            return absolute or relative or attempt == solution
    except (TypeError, ValueError):  # Unexpected values or failed cast
        log.exception(
            "Unexpected values in comparison: Student: %s (%s), Corrector: %s (%s), Tolerance: %s (%s)",
            attempt,
            type(attempt),
            solution,
            type(solution),
            tolerance_rel,
            type(tolerance_rel),
        )
    return False


def find_valid_filenames():
    """
    Searches for corrector files in configured folders. Returns dictionary containing codename as key
    and :class:`excel.Corrector` as value.

    :return: Dict[str, excel.Corrector]
    """
    # Reset dict
    valid_filenames = {}

    # Find configured groups
    for group in config.FOLDERS:
        group = Path(os.path.abspath(group))

        # Iterate over subjects
        for subject in group.iterdir():
            # Ignore files and blacklisted folders
            if not subject.is_dir() or subject.name in config.FOLDER_IGNORE:
                continue

            # Try to find corrector.xlsx
            exc = excel.Corrector.from_path(subject)
            if exc and exc.valid:
                if exc.codename.lower() in valid_filenames:
                    log.error("Duplicate codenames: %s", exc.codename)
                    utils.write_error(
                        exc.parent_path,
                        (
                            f"Der Dateiname {exc.codename} ist bereits registriert "
                            f'für Corrector "{valid_filenames[exc.codename.lower()].get_relevant_path()}!"'
                        ),
                    )
                    continue

                # Append subject and Corrector
                log.info(
                    'Registered %s with file name "%s"',
                    exc.corrector_title,
                    exc.codename,
                )
                valid_filenames[exc.codename.lower()] = exc
    return valid_filenames


def main():
    # Dict containing file name as key and Corrector as value
    valid_filenames = find_valid_filenames()

    if len(valid_filenames) == 0:
        log.info("There's nothing to do.")
        return

    # Idling mail instance
    mail_instance = mail.Mail()

    # Check inbox for new mails/submitted files
    student_files = mail_instance.check_inbox(valid_filenames)

    # Sort by codename/module number
    student_files.sort(key=lambda x: x["corrector"].codename)

    # Correct each file
    for sf in student_files:
        try:
            corrector = sf["corrector"]
            e = excel.Student(sf["student"])

            # Couldn't find any solutions in submitted file
            if len(e.solutions) == 0:
                log.warning("Found no solutions in submitted file")
                mail_instance.send(
                    e.student_email, *mail.Generator.malformed_attachment()
                )
                continue

            real_solutions = corrector.generate_solutions(e.mat_num, e.dummies)

            # Couldn't find any solutions in submitted file
            if len(e.solutions) != len(real_solutions):
                log.warning("Found more/fewer tasks in submitted file")
                mail_instance.send(
                    e.student_email, *mail.Generator.malformed_attachment()
                )
                continue

            compared_solutions = []

            # List of passed/blocked exercises
            exercises_blocked = []
            exercises_passed = []
            for idx, student_solution in enumerate(e.solutions):
                # Ignore exercise if one of the fields is empty
                if None in student_solution:
                    continue

                # region First block/pass check for exercise
                # Check if user is blocked or passed the exercise previously
                blocked, passed = e.get_stats(idx, corrector.max_attempts)
                if blocked:
                    log.info(
                        "Ignoring exercise %s since %s is already blocked",
                        idx + 1,
                        e.student_email,
                    )
                    exercises_blocked.append(idx)
                    continue
                elif passed:
                    log.debug(
                        "Ignoring exercise %s since %s has already passed this exercise",
                        idx + 1,
                        e.student_email,
                    )
                    exercises_passed.append(idx)
                    continue

                log.debug("Processing exercise %s for %s", idx + 1, e.student_email)

                # endregion

                # region Comparison of submitted solutions with corrector
                corrector_solution = real_solutions[idx]
                # Make sure the student didn't somehow delete any exercise part
                if len(student_solution) != len(corrector_solution):
                    log.warning(
                        "%s may have tampered with the excel file, got different amount of sub "
                        "exercises for exercise %s",
                        e.student_email,
                        idx + 1,
                    )
                    continue

                # Vector for single exercise
                exercise_solved = {
                    "exercise": idx,
                    "correct": [False] * len(student_solution),
                    "var_names": [],
                }
                for partial_idx, partial in enumerate(student_solution):
                    exercise_solved["correct"][partial_idx] = compare(
                        partial,
                        corrector_solution[partial_idx]["value"],
                        corrector_solution[partial_idx]["tolerance_rel"],
                        corrector_solution[partial_idx]["tolerance_abs"],
                    )
                    exercise_solved["var_names"].append(
                        corrector_solution[partial_idx]["name"]
                    )

                # Update student block/pass stats
                perc = int(
                    sum(exercise_solved["correct"])
                    / len(exercise_solved["correct"])
                    * 100
                )
                blocked, passed = e.update_stats(idx, perc, corrector.max_attempts)
                if passed:
                    exercises_passed.append(idx)
                if blocked:
                    exercises_blocked.append(idx)

                compared_solutions.append(exercise_solved)
            # endregion

            # region Sending passed/blocked/congrats mails
            # Send results
            results = ""
            for solution in compared_solutions:
                results += mail.Generator.exercise_details(solution)

            # May be empty if nothing was submitted
            if len(results) > 0:
                mail_instance.send(
                    e.student_email, f"Ergebnisse: {corrector.corrector_title}", results
                )
                log.debug("Sending results")

            # Send mail informing about passed exercises
            if len(exercises_passed) > 0:
                mail_instance.send(
                    e.student_email,
                    *mail.Generator.exercise_passed(
                        corrector.corrector_title, exercises_passed, e.mat_num
                    ),
                )
                log.debug("Sending passed")

            # Send mail informing about blocked exercises
            if len(exercises_blocked) > 0:
                mail_instance.send(
                    e.student_email,
                    *mail.Generator.exercise_blocked(
                        corrector.corrector_title,
                        exercises_blocked,
                        corrector.max_attempts,
                    ),
                )
                log.debug("Sending blocked")

            # Send final congrats
            if len(exercises_passed) == len(real_solutions):
                mail_instance.send(
                    e.student_email,
                    *mail.Generator.exercise_congrats(
                        corrector.corrector_title, e.mat_num
                    ),
                )
                log.debug("Sending final congrats")
            # endregion
        except excel.ExcelFileException:
            log.exception("Error during processing of student file.")
        except IOError:
            log.exception("Critical error during processing. Quitting.")
            raise

    if len(student_files) > 0:
        # Run post processing on all matched correctors
        for corrector in set(_["corrector"] for _ in student_files):
            post.PostProcessing(
                corrector.parent_path, len(corrector.exercise_ranges)
            ).run()


if __name__ == "__main__":
    # Import config
    config = utils.import_config()
    log = utils.setup_logger(logging.DEBUG if config.DEBUG else logging.INFO)

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
            print("Passwords to not match or empty password supplied!")
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

        # Wait x minutes between each run
        log.info("Pausing for %s minutes", config.DELAY_SLEEP)
        time.sleep(config.DELAY_SLEEP * 60)
