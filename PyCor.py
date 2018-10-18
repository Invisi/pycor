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
import email_functions as emf
import excel_functions as exf
import folder_functions as ff
import postprocessing_functions as pf

__version__ = "1.4"
delay = 15


def main():
    """Invokes the correction function iteratively."""

    af.Welcome(__version__)  # Prints message..

    # We get Profs. directories.

    folders = ff.get_profs()

    # Iterating over the Profs. folders.

    for profs in folders:

        profs = profs[:-1]  # This deletes the line jump '\n'

        print("Accessing folder:", profs)

        # Iterate over the subjects folders.

        subf = os.listdir(profs)
        print("Filtering the folders list.")
        subf_filt = ff.filter_folders(profs + '\\', subf)

        for dir in subf_filt:

            current_dir = profs + "\\" + dir + "\\"
            print("Accessing subfolder: ", current_dir)

            # Is there a valid 'corrector.xlsx' file?
            print("Checking if there is a valid corrector...")
            [flag_corr, psw_corr, corr_path] = exf.corrector_ready(current_dir)
            # corr_path = current_dir + 'corrector.xlsx'    # or other: 'xls'

            # Let's start with the correction process

            if flag_corr == 1:
                # Get the important data from the corrector...

                [Subject, usn, psw, TMax, TInit, TEnd, TotalTeils, Cerror] = \
                    exf.corrector_get_data(corr_path, psw_corr)

                Subject = Subject.encode('utf-8', 'ignore')

                print("Correcting: ", Subject)

                emf.check_INBOX(usn, psw, current_dir, corr_path, Subject,
                                psw_corr, TotalTeils)

                # Post processing only when a corrector is ready.

                print("Producing postprocessing..")

                pf.full_analyze(current_dir, TotalTeils)

            af.NextFolder()  # Prints message..


if __name__ == '__main__':

    # Insert 'while' conditions here..

    while af.system_checkout() == 1:
        main()
        print("Wait (s)", delay)
        time.sleep(delay)
    else:
        print("PyCor could not find the correct system setup and stopped.")
