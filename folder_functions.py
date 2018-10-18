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


def create_token(path):

    f = open(path + 'token', 'w')
    f.close()

    return


def get_folders(file):
    """Given a filename 'file', this function reads this 'file' which contains
    the name of the folders which have to be iterated over."""

    data = open(file, 'r')
    folders = []
    for line in data:
        folders.append(line)

    data.close()

    return folders


def get_filename(mainfolder):
    """Looks for '*.folders' in the 'mainfolder' path."""

    # Initialize foldername
    foldername = ""
    names = os.listdir(mainfolder)

    for item in names:
        if item[-7:] == "folders":
            foldername = item

    return foldername


def get_profs():
    """Gets directly the Profs directories."""

    mainfolder = "."
    filename = get_filename(mainfolder)
    folders = get_folders(filename)

    return folders


def create_folder(Sub_dir):
    """This function creates a specified folder 'Sub_dir', only in case it was
    not already existing."""

    if not os.path.exists(Sub_dir):    # Before creating, check it.

        os.makedirs(Sub_dir)

        print "Creating new directory at: ", Sub_dir

    else:    # If the directory was already existing...

        print "Directory already existing: ", Sub_dir

    return


def filter_folders(current_dir, list_folders):

    nk = len(list_folders)
    path = []
    list_filtered = []
    for k in range(0, nk):
        path.append(current_dir + list_folders[k])

        if os.path.isdir(path[k]):
            list_filtered.append(list_folders[k])

    if '_postprocessing' in list_filtered:

        list_filtered.remove('_postprocessing')

    if 'PyCor_documentation' in list_filtered:

        list_filtered.remove('PyCor_documentation')

    return list_filtered
