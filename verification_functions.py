# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'verification_functions' is a complementary sub-library containing different
functions created for the general PyCor super-library.

This module is intended to provide the necessary functions for the correction.

Functions:

    +compare(exer_x, corr_x, error_x):
    This function compares 'exer_x' to 'corr_x', accounting for 'error_x' in
    case of numerical values or direct match for text.

    +correction_Teil(Teil, varNames, solutions, Cerror, exer_sol, j):
    This function is the kernel of PyCor. Compares submitted values of
    'exer_sol' to 'solutions' obtained from the updated corrector. 'j' states
    which Teil is corrected (j = i + 1, 'i' from submitted). 'Teil' list
    provides indication on which variable belongs to which Teil. 'Cerror'
    provides the accuracy margin for each variable.

    +search_zero(Teil_j_block):
    This function looks for the position of the first zero in an array
    'Teil_j_block'. It is previously checked that there is at least one
    zero. The position of this zero is returned as 'k'.

    +is_blocked(Sub_dir, j):
    This function checks that the exercise 'j' located in a given path
    'Sub_dir' is not blocked. Returns a flag with value 0 for not blocked and
    1 for blocked.

    +is_passed(Sub_dir, j):
    This function checks that the exercise 'j' located in a given path
    'Sub_dir' is not passed. Returns a flag with value 0 for not passed and
    1 for passed.

    +update_stats(Sub_dir, j, r, usn, psw, eaddress, TrialsMax):
    This function is one of the main functions of PyCor. It takes the last
    result 'r' of and exercise 'j' and stores it in the path 'Sub_dir'. The
    e-mail of the correction 'usn' and its password 'psw' are necessary to warn
    the students with any change in their status: blocked/passed. The number of
    maximum attempts 'TrialsMax' is also necessary since it directly affects
    the blocking process.

"""


import email_functions as emf
import folder_functions as ff


import numpy as np
import time


def compare(exer_x, corr_x, error_x):
    """This function compares 'exer_x' to 'corr_x', accounting for 'error_x' in
    case of numerical values or direct match for text."""

    # Check if it's numeric or string. When string, compare literally.
    # When numeric, apply the 'error_x' margin.

    resol_x = 0

    try:

        # In case the stundent made a blank space after the comma...
        if isinstance(corr_x, float) and isinstance(exer_x, unicode):

            print "Converting the submitted variable from unicode to float..."

            exer_x = exer_x.replace(',', '.')
            exer_x = float(exer_x.replace(' ', ''))

        # Make sure we have floats!
        if isinstance(corr_x, int):
            corr_x = float(corr_x)
        if isinstance(exer_x, int):
            exer_x = float(exer_x)
        if isinstance(error_x, int):
            error_x = float(error_x)

        # This is the main comparison:

        if (type(corr_x) == unicode and exer_x == corr_x):

            resol_x = 1

        elif (type(corr_x) == unicode and exer_x != corr_x):

            resol_x = 0

        # this accounts for > 0
        elif exer_x <= (1.0 + 0.01*error_x)*corr_x and \
                (1.0 - 0.01*error_x)*corr_x <= exer_x:

            resol_x = 1

        # this accounts for < 0
        elif exer_x >= (1.0 + 0.01*error_x)*corr_x and \
                (1.0 - 0.01*error_x)*corr_x >= exer_x:

            resol_x = 1

        else:

            resol_x = 0
            print "submitted: ", exer_x
            print "real solution: ", corr_x

    except Exception:

            print "Unexpected variable comparison!"
            resol_x = 0
            print "submitted: ", exer_x
            print "real solution: ", corr_x

    return resol_x


def correct_Teil(Teil, solutions, Cerror, exer_sol, j, varNames, usn, psw,
                 eaddress, Subject):
    """This function is the kernel of PyCor. Compares submitted values of
    'exer_sol' to 'solutions' obtained from the updated corrector. 'j' states
    which Teil is corrected (j = i + 1, 'i' from submitted). 'Teil' list
    provides indication on which variable belongs to which Teil. 'Cerror'
    provides the accuracy margin for each variable."""

    nk = len(Teil)
    resol = np.zeros(nk)
    active = np.zeros(nk)
    names = []

    for k in range(0, nk):

        # print "Checking variable no. ", k

        if Teil[k] == j:    # Then, we can 'compare' the values...

            # print "Which corresponds to a submitted exercise."

            active[k] = 1

            # print "Compare to corrector.xlsx solution."

            resol[k] = compare(exer_sol[k], solutions[k], Cerror[k])   # 0 or 1
            names.append(varNames[k])

    # Check all the acitve[k] == 1 results. Give back result (percentage)

    result = 100.0 * np.sum(resol)/np.sum(active)
    # print "Obtained results: ", resol
    # print "Over active vars: ", active
    # print "Whose var. names: ", names

    # We may delete elements from arrays 'resol' and 'active' to fit dimensions
    # of 'names'. Let's call it: 'reshaping solution vectors'.

    print "Reshaping solution vectors."
    resol_filtered = resol[np.where(active > 0.5)]

    if result == 0:

        result = 0.01    # We will use 0 as an empty slot in the results saving

    # We should send an e-mail with the results!!

    # print "Sending e-mail with correction results for exercise: ", j
    #
    # # 'html' format:
    # [R_subj, R_msg] = emf.generatehtmlmsg_Results(j, names, resol_filtered,
    #                                               Subject)
    # emf.send_email(usn, psw, eaddress, R_subj, R_msg)

    return result, names, resol_filtered


def search_zero(Teil_j_block):
    """This function looks for the position of the first zero in an array
    'Teil_j_block'. It is previously checked that there is at least one
    zero. The position of this zero is returned as 'k'."""

    # By using the magic of enumerators:

    k = next(x[0] for x in enumerate(Teil_j_block) if x[1] == 0.0)

    return k


def is_blocked(Sub_dir, j, TrialsMax):
    """This function checks that the exercise 'j' located in a given path
    'Sub_dir' is not blocked. Returns a flag with value 0 for not blocked and
    1 for blocked."""

    flag_blocked = 0
    Teil_j_str = "Exercise" + str(j)

    try:
        block_dir = Sub_dir + "\\" + Teil_j_str + "_block.txt"
        Teil_j_block = np.loadtxt(block_dir)

        # if correctly loaded, check if < TrialsMax

        nt = len(Teil_j_block)

        if nt < TrialsMax:  # Extend the dimensions of Teil_j_block

            Teil_j_block = np.append(Teil_j_block,
                                     np.zeros(int(TrialsMax - nt)))
            np.savetxt(block_dir, Teil_j_block, fmt='%3.2f')

        elif nt > TrialsMax:  # Extend the dimensions of Teil_j_block

            Teil_j_block = Teil_j_block[0:int(TrialsMax)]
            np.savetxt(block_dir, Teil_j_block, fmt='%3.2f')

        if Teil_j_block[-1] != 0:

            print "This student is already blocked for exercise: ", j

            flag_blocked = 1

    except Exception:

        flag_blocked = 0

    return flag_blocked


def is_passed(Sub_dir, j):
    """This function checks that the exercise 'j' located in a given path
    'Sub_dir' is not passed. Returns a flag with value 0 for not passed and
    1 for passed."""

    flag_passed = 0
    Teil_j_str = "Exercise" + str(j)

    try:

        Teil_j_block = np.loadtxt(Sub_dir + "\\data\\" + Teil_j_str +
                                  "_all.txt")

        nk = len(Teil_j_block)

        for k in range(0, nk):

            if Teil_j_block[k] == 100:

                print "This student has already passed exercise: ", j
                flag_passed = 1

    except Exception:

        flag_passed = 0

    return flag_passed


def update_stats(Sub_dir, j, r, usn, psw, eaddress, TrialsMax, Subject,
                 MatNum):
    """This function is one of the main functions of PyCor. It takes the last
    result 'r' of and exercise 'j' and stores it in the path 'Sub_dir'. The
    e-mail of the correction 'usn' and its password 'psw' are necessary to warn
    the students with any change in their status: blocked/passed. The number of
    maximum attempts 'TrialsMax' is also necessary since it directly affects
    the blocking process."""

    Teil_j_str = "Exercise" + str(j)
    flag_blocked = 0    # this flag will mark if the student is or was blocked.
    flag_passed = 0     # this flag will mark if the student has passed.
    # flag_final = 0      # this flag will mark if ALL Teile are passed.

    # Let's start with 'Teil_j_block' -----------------------------------------
    # This is the most important stored variable.

    try:

        Teil_j_block = np.loadtxt(Sub_dir + "\\" + Teil_j_str + "_block.txt")

        # if correctly loaded, check if < TrialsMax

        if Teil_j_block[-1] != 0:

            print "This student is already blocked for exercise: ", j

            flag_blocked = 1

        else:

            k = search_zero(Teil_j_block)   # 'k' is the first zero position
            print "First zero at position: ", k
            Teil_j_block[k] = r

            # Is him now blocked? Check that last item is not 0 or 100.

            if Teil_j_block[-1] != 0 and Teil_j_block[-1] != 100:

                print "This student has been now blocked for exercise: ", j

                flag_blocked = 1

    except Exception:

        Teil_j_block = np.zeros(int(TrialsMax))
        Teil_j_block[0] = r

    # Whatever we did with the array, let's save it now!

    block_dir = Sub_dir + "\\" + Teil_j_str + "_block.txt"
    print "Saving blocking stats at: ", block_dir
    np.savetxt(block_dir, Teil_j_block, fmt='%3.2f')

    # Check which is the status of the student: blocked? passed? nothing?

    # if flag_blocked == 1:
    #
    #     # Send e-mail to warn the student about his Exercise 'j' blocking.
    #
    #     [B_subj, B_msg] = emf.generatemsg_blockedTeil_i(j, Subject, TrialsMax)
    #     emf.send_email(usn, psw, eaddress, B_subj, B_msg)

    if r == 100:

        flag_passed = 1

        # # Send e-mail to warn the student about his Exercise 'j' blocking.
        #
        # [P_subj, P_msg] = emf.generatemsg_passedTeil_i(j, Subject, MatNum)
        # emf.send_email(usn, psw, eaddress, P_subj, P_msg)

    # Some data should be stored appart: ../data ------------------------------

    data_dir = Sub_dir + "\\" + "data"
    print "Creating folder: ", data_dir
    ff.create_folder(data_dir)

    # -------------------------------------------------------------------------
    # Let's start with 'Teil_j_all' -------------------------------------------
    # We have checked before that the exercise was not previously passed.
    # We can therefore store this result.

    all_dir = data_dir + "\\" + Teil_j_str + "_all.txt"

    try:    # Is there already a Teil_j_all file?

        Teil_j_all = np.loadtxt(all_dir)
        Teil_j_all = np.append(Teil_j_all, [r])

    except Exception:   # or shall we create it?

        Teil_j_all = np.array([r])

    # anycase, we should save it.

    np.savetxt(all_dir, Teil_j_all, fmt='%3.2f')

    # -------------------------------------------------------------------------
    # Let's start with time variables -----------------------------------------

    ti = time.localtime()

    Y_dir = data_dir + "\\" + Teil_j_str + "_Y.txt"
    M_dir = data_dir + "\\" + Teil_j_str + "_M.txt"
    D_dir = data_dir + "\\" + Teil_j_str + "_D.txt"
    h_dir = data_dir + "\\" + Teil_j_str + "_h.txt"

    try:

        Teil_j_Y = np.loadtxt(Y_dir)
        Teil_j_Y = np.append(Teil_j_Y, [ti[0]])

        Teil_j_M = np.loadtxt(M_dir)
        Teil_j_M = np.append(Teil_j_M, [ti[1]])

        Teil_j_D = np.loadtxt(D_dir)
        Teil_j_D = np.append(Teil_j_D, [ti[2]])

        Teil_j_h = np.loadtxt(h_dir)
        Teil_j_h = np.append(Teil_j_h, [ti[3]])

    except Exception:

        Teil_j_Y = np.array([ti[0]])
        Teil_j_M = np.array([ti[1]])
        Teil_j_D = np.array([ti[2]])
        Teil_j_h = np.array([ti[3]])

    # anycase, we should save it.

    np.savetxt(Y_dir, Teil_j_Y, fmt='%d')
    np.savetxt(M_dir, Teil_j_M, fmt='%d')
    np.savetxt(D_dir, Teil_j_D, fmt='%d')
    np.savetxt(h_dir, Teil_j_h, fmt='%d')

    # -------------------------------------------------------------------------
    # Let's start MatNum ------------------------------------------------------

    MatNum_dir = data_dir + "\\" + 'MatNum.txt'

    try:

        MN = np.loadtxt(MatNum_dir)
        MN = np.append(MN, MatNum)

    except Exception:

        MN = np.array([MatNum])

    np.savetxt(MatNum_dir, MN, fmt='%d')

    # -------------------------------------------------------------------------

    print "Stats updated!"

    return flag_blocked, flag_passed


def check_final(Sub_dir, TotalTeils):

    print "Total number of exercises: ", TotalTeils

    flag_final = 0
    data_dir = Sub_dir + "\\" + "data"

    jmax = int(TotalTeils)
    passed = np.zeros([jmax])

    for j in range(0, jmax):

        Teil_j_str = "Exercise" + str(j + 1)
        all_j_dir = data_dir + "\\" + Teil_j_str + "_all.txt"

        try:

            Teil_j_all = np.loadtxt(all_j_dir)
            MaxScore = np.max(Teil_j_all)

            if MaxScore == 100:

                passed[j] = 1

            else:

                passed[j] = 0

        except Exception:

            passed[j] = 0

    if np.sum(passed) == jmax:

        flag_final = 1

    return flag_final

def check_MNvalid(MN):

    # 0 for single value, 1 otherwise

    nk = len(MN)
    flag = 0

    for k in range(0, nk):

        if MN[k] != MN[0]:

            flag = 1

    return flag


def count_different(MN):

    MN_red = np.unique(MN)
    k = len(MN_red)

    return k
