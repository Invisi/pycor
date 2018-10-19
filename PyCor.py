# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

PyCor automatic corrector.

This module is intended to provide an easy to use and intuitive tool for any
person correcting exercises in FH Aachen.

The flow chart of PyCor would be as follows:

1. Starts iterating over a list of folders (i.e.: Bung\PyCor, Hoettges\PyCor)
2. At each one of these folders, subfolders are identified as differents
    subjects. Then we iterate over these subfolders, using correction().
3. Before starting correction(), a "corrector.xls" is searched. If not found,
    PyCor skips this subfolder.
4. "corrector.xls" is read, looking for key parameters (i.e.: deadline). These
    variables are stored and are the input for "corrector(args)". All the args
    should be resetted before checking the "corrector.xls" file. Thus a
    resetter is needed.

"""

import os
import time

import auxiliar_functions as af
import config
import email_functions as emf
import excel_functions as exf
import postprocessing_functions as pf

__version__ = "1.4"
delay = 15


def main():
    """Invokes the correction function iteratively."""

    # Iterate over configured folders
    for prof_folder in config.FOLDERS:
        prof_folder = os.path.abspath(prof_folder)
        log.info("Accessing folder: %s", prof_folder)

        # Iterate over the subjects folders.
        for subject_folder in os.listdir(prof_folder):
            current_dir = os.path.join(prof_folder, subject_folder)
            if not os.path.isdir(current_dir) or subject_folder in config.FOLDER_IGNORE:
                log.debug('Skipping %s', subject_folder)
                continue
            log.info("Accessing subfolder: %s", current_dir)

            # Is there a valid 'corrector.xlsx' file?
            log.info("Checking if there is a valid corrector...")
            corr_path, psw_corr = exf.corrector_ready(current_dir)

            # Let's start with the correction process
            if corr_path and False:
                # Get the important data from the corrector...

                [Subject, usn, psw, TMax, TInit, TEnd, TotalTeils, Cerror] = \
                    exf.corrector_get_data(corr_path, psw_corr)

                log.info("Correcting: %s", Subject)
                emf.check_INBOX(usn, psw, current_dir, corr_path, Subject,
                                psw_corr, TotalTeils)

                # TODO: Run in another thread/as another process
                # Post processing only when a corrector is ready.
                log.info("Producing postprocessing..")

                pf.full_analyze(current_dir, TotalTeils)

            log.info("____________")


if __name__ == '__main__':
    log = af.setup_logger()

    af.welcome(__version__)

    # Insert 'while' conditions here..
    while af.system_checkout():
        main()
        log.info("Wait %ss", delay)
        time.sleep(delay)
    else:
        print("PyCor could not find the correct system setup and stopped.")
