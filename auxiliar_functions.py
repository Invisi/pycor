# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'auxiliar_functions' is a complementary sub-library containing different
functions created for the general PyCor super-library.

This module is intended to provide some auxiliar functions which doesn't fit
in the remaining Python libraries prepared for PyCor.

Functions:

    +Welcome(__version__):
    This functions prints a Welcome message once PyCor is started. This
    message includes information as: PyCor version and running time. The
    printed message is only visible in the server where PyCor is running.

    +NextFolder():
    This function prints a message warning about the end of a folder by
    PyCor. It does not require any input and the message is only visible in the
    server where PyCor is running.

"""
import logging
import os
import sys
import time

import simplecrypt


def welcome(__version__):
    """This functions prints a Welcome message once PyCor is started. This
    message includes information as: PyCor version and running time. The
    printed message is only visible in the server where PyCor is running."""

    print("____________\n____________\nWelcome to PyCor v.", __version__)
    print("Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.")
    print("____________\n")
    print("PyCor is now running!")
    t_i = time.localtime()
    print("Current date (Y, M, D): ", t_i[0], t_i[1], t_i[2])
    print("Current time (h, m, s): ", t_i[3], t_i[4], t_i[5])

    return


def setup_logger():
    log = logging.getLogger('PyCor')
    log.setLevel(logging.DEBUG)

    # hldr = logging.FileHandler('PyCor.log')
    fmt = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S")
    # hldr.setFormatter(fmt)
    # log.addHandler(hldr)
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(fmt)
    log.addHandler(stream)
    return log


def system_checkout():
    path = 'C:\\Windows\\dava'

    if os.environ.get('ENV', '') == 'DEVELOPMENT':
        return True

    try:
        system_file = open(path, 'r')
        system_file_enc = system_file.read()
        system_file.close()

        system_file_dec = simplecrypt.decrypt('license', system_file_enc)

        print("Decrypted: ", system_file_dec)

        if system_file_dec == 'OK':
            print("System is ready for PyCor.")
            return True
        print("System is not ready for PyCor.")
        return False
    except (IOError, simplecrypt.DecryptionException):
        return False
