# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'folder_functions' is a complementary sub-library containing different
functions created for the general PyCor super-library.

This module is intended to provide access to folders and help data handling.

Functions:

    +get_folders(file):
    Given a filename 'file', this function reads this 'file' which contains the
    name of the folders which have to be iterated over.

    +get_filename(mainfolder):
    Looks for '*.folders' in the 'mainfolder' path.

    +get_profs():
    Gets directly the Profs directories.

    +create_folder(Sub_dir):
    This function creates a specified folder 'Sub_dir', only in case it was
    not already existing.

"""

import os


def create_folder(Sub_dir):
    # TODO: Move to utils
    """This function creates a specified folder 'Sub_dir', only in case it was
    not already existing."""

    if not os.path.exists(Sub_dir):  # Before creating, check it.
        os.makedirs(Sub_dir)
        print("Creating new directory at: ", Sub_dir)

    else:  # If the directory was already existing...
        print("Directory already existing: ", Sub_dir)

    return

