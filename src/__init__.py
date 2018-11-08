import datetime
import logging
import os
import time
import traceback
from pathlib import Path

import numpy as np

import excel
import mail
import post
import utils

__version__ = '1.5.0'


def compare(attempt: float or str, solution: float or str, tolerance: int or float):
    try:
        # In case the student made a space after the comma...
        if isinstance(attempt, str) and isinstance(solution, float):
            attempt = float(attempt.replace(',', '.').replace(' ', ''))

        # Make sure to convert ints to floats
        if isinstance(solution, int):
            solution = float(solution)
        if isinstance(attempt, int):
            attempt = float(attempt)
        if isinstance(tolerance, int):
            tolerance = float(tolerance)

        # Compare string values
        if isinstance(solution, str):
            return solution == attempt

        # Compare numerical values
        if isinstance(solution, float):
            return (1 - tolerance / 100) * solution <= attempt <= (1 + tolerance / 100) * solution
    except (TypeError, ValueError):
        log.error(traceback.format_exc())

    # Unexpected values
    log.error('Unexpected values in comparison: Student: %s (%s), Corrector: %s (%s), Tolerance: %s (%s)', attempt,
              type(attempt), solution, type(solution), tolerance, type(tolerance))
    return False


def get_stats(student_folder: Path, exercise: int, max_tries: int) -> tuple:
    try:
        exercise_file = student_folder / 'Exercise{}_block.txt'.format(exercise + 1)

        if os.path.exists(exercise_file):
            # noinspection PyTypeChecker
            block_status = np.loadtxt(exercise_file)

            # Check if user's try list doesn't match the specified max_tries
            if len(block_status) > max_tries:
                # Shorten list
                block_status = block_status[0:max_tries]
                np.savetxt(exercise_file, block_status, fmt='%3.2f')
            elif len(block_status) < max_tries:
                # Extend list
                block_status = block_status + [0] * (max_tries - len(block_status))
                np.savetxt(exercise_file, block_status, fmt='%3.2f')

            if 0 < block_status[-1] < 100:
                return True, False
            elif block_status[-1] == 0:
                return False, False
            else:
                return False, True
        else:
            return False, False
    except IOError:
        log.error(traceback.format_exc())
        raise


def update_stats(student_folder: Path, exercise: int, correct_perc: int, max_tries: int, mat_num: int) -> tuple:
    """
    Updates and saves the student's statistics

    :param student_folder: Full path to student's folder
    :param exercise: Exercise number [beginning at 0]
    :param correct_perc: Percentage of correctly answered sub tasks
    :param max_tries: Maximum amount of tries before being blocked
    :param mat_num: Student's matriculation number
    :return:
    """
    try:
        blocked = False
        passed = correct_perc == 100

        student_email = student_folder.name
        exercise_file = student_folder / 'Exercise{}_block.txt'.format(exercise + 1)
        student_data = student_folder / 'data'

        # Does the file exist already?
        if exercise_file.exists():
            # noinspection PyTypeChecker
            block_status = np.loadtxt(exercise_file)
        else:
            block_status = np.zeros(max_tries)

        # Check if user still has tries
        try_row = [r[0] for r in enumerate(block_status) if r[1] == 0]

        # There's at least one attempt left, let's log the results!
        if len(try_row) > 0:
            block_status[try_row[0]] = correct_perc
            np.savetxt(exercise_file, block_status, fmt='%3.2f')
        else:
            blocked = True

        # Check if student failed his last try
        if 0 < block_status[-1] < 100:
            blocked = True
            log.info('Blocked %s for exercise %s', student_email, exercise + 1)

        # Create user-specific data folder
        if not student_data.exists():
            student_data.mkdir(exist_ok=True)

        current_datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Save percentage and date/time to file
        exercise_all = student_data / 'Exercise{}.txt'.format(exercise + 1)
        with exercise_all.open('a') as ea:
            ea.write('{} - {}\n'.format(current_datetime, correct_perc))

        # Save used mat_num to file
        mat_num_file = student_data / 'mat_num.txt'
        with mat_num_file.open('a') as mnf:
            mnf.write('{} - {}\n'.format(current_datetime, mat_num))

        return blocked, passed
    except IOError:
        log.error(traceback.format_exc())
        raise


def main():
    # Iterate over configured folders
    for prof_folder in config.FOLDERS:
        prof_folder = Path(prof_folder).resolve()
        log.debug("Accessing folder: %s", prof_folder)

        # Iterate over the subjects folders.
        for subject_folder in prof_folder.iterdir():
            if not subject_folder.is_dir() or subject_folder.name in config.FOLDER_IGNORE:
                log.debug('Skipping %s', subject_folder)
                continue
            log.debug("Accessing subfolder: %s", subject_folder)

            # Try to find and open corrector
            try:
                exc = excel.ExcelCorrector.from_subject_folder(subject_folder)
                if exc:
                    student_files = exc.email.check_inbox()
                    for s_file in student_files:
                        student_email = s_file.parent.name
                        try:
                            e = excel.ExcelStudent(s_file)

                            real_solutions = exc.generate_solutions(e.mat_num, e.dummies)
                            e.destroy()

                            compared_solutions = []

                            # List of passed/blocked exercises
                            exercises_blocked = []
                            exercises_passed = []
                            for exercise_idx, student_solution in enumerate(e.solutions):
                                # Ignore exercise if one of the fields is empty
                                if None in student_solution:
                                    continue

                                # Check if user is blocked or passed the exercise previously
                                blocked, passed = get_stats(e.student_folder, exercise_idx, exc.max_tries)
                                if blocked:
                                    log.info('Ignoring exercise %s since %s is already blocked', exercise_idx + 1,
                                             e.student_email)
                                    exercises_blocked.append(exercise_idx)
                                    continue
                                if passed:
                                    log.debug('Ignoring exercise %s since %s has already passed this exercise',
                                              exercise_idx + 1,
                                              e.student_email)
                                    exercises_passed.append(exercise_idx)
                                    continue

                                log.debug('Processing exercise %s for %s', exercise_idx + 1, e.student_email)

                                corrector_solution = real_solutions[exercise_idx]
                                # type: dict
                                # Make sure the student didn't somehow delete any exercise part
                                if len(student_solution) != len(corrector_solution):
                                    log.warning(
                                        '%s may have tampered with the excel file, got different amount of sub '
                                        'exercises for exercise %s', e.student_email, exercise_idx + 1)
                                    continue

                                # Vector for single exercise
                                exercise_solved = {'exercise': exercise_idx,
                                                   'correct': [False] * len(student_solution),
                                                   'var_names': []}
                                for partial_idx, partial in enumerate(student_solution):
                                    exercise_solved['correct'][partial_idx] = compare(partial,
                                                                                      corrector_solution[
                                                                                          partial_idx]['value'],
                                                                                      corrector_solution[
                                                                                          partial_idx]['tolerance'])
                                    exercise_solved['var_names'].append(corrector_solution[partial_idx]['name'])

                                # Update student block/pass stats
                                perc = int(sum(exercise_solved['correct']) / len(exercise_solved['correct']) * 100)
                                blocked, passed = update_stats(e.student_folder, exercise_idx, perc, exc.max_tries,
                                                               e.mat_num)
                                if passed:
                                    exercises_passed.append(exercise_idx)
                                if blocked:
                                    exercises_blocked.append(exercise_idx)

                                compared_solutions.append(exercise_solved)

                            # Send results
                            results = ''
                            for solution in compared_solutions:
                                results += mail.Generator.exercise_details(solution)

                            # May be empty if nothing was submitted
                            if len(results) > 0:
                                exc.email.send(e.student_email, 'Ergebnisse: {}'.format(exc.subject_name), results)

                            # Send mail informing about passed exercises
                            if len(exercises_passed):
                                exc.email.send(e.student_email, *mail.Generator.exercise_passed(exc.subject_name,
                                                                                                exercises_passed,
                                                                                                e.mat_num))
                            # Send mail informing about blocked exercises
                            if len(exercises_blocked) > 0:
                                exc.email.send(e.student_email, *mail.Generator.exercise_blocked(exc.subject_name,
                                                                                                 exercises_blocked,
                                                                                                 exc.max_tries))

                            # Send final congrats
                            if len(exercises_passed) == len(real_solutions):
                                exc.email.send(e.student_email,
                                               *mail.Generator.exercise_congrats(exc.subject_name, e.mat_num))

                        except excel.ExcelFileException:
                            log.warning(traceback.format_exc())
                            log.error('Error during processing of %s', s_file.parts[-3:])
                            exc.email.send(student_email, *mail.Generator.error_processing(exc.subject_name))
                        except IOError:
                            log.error(traceback.format_exc())
                            log.error('Error during processing of %s', s_file.parts[-3:])
                            exc.email.send(student_email, *mail.Generator.error_processing(exc.subject_name))
                            raise  # Fail securely

                    exc.destroy()

                    # Run post processing if we got new submissions
                    if len(student_files) > 0:
                        post.PostProcessing(subject_folder, len(exc.exercise_ranges)).run()
                else:
                    log.debug('Ignoring folder %s due to missing corrector file', subject_folder)
            except excel.ExcelFileException:
                log.debug(traceback.format_exc())

            log.info("____________")
            time.sleep(config.DELAY_MAILBOXES)


if __name__ == '__main__':
    # Import config
    config = utils.import_config()

    log = utils.setup_logger(logging.DEBUG if config.DEBUG else logging.INFO)

    log.info("Welcome to PyCor v.%s", __version__)
    log.info('____________')
    log.info("PyCor is now running!")

    while True:
        main()

        # Wait 5 minutes between each run
        log.info('Pausing for {} minutes'.format(config.DELAY_SLEEP))
        time.sleep(config.DELAY_SLEEP)
