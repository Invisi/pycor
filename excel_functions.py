# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'excel_functions' is a complementary sub-library containing different functions
created for the general PyCor super-library.

This module is intended to provide access to excel files.

Functions:

    +exercise_get_data(path):
    This function takes the 'path' to the submitted file by the student.
    Once opened, read the Matr. Num. and 'dummies', automatically detects the
    number of Teils and gets the initial and ending indexes. Also, reads the
    solutions provided by the student which can be later compared to the
    corrector.

    +check_allNone(exercise_k):
    This function checks if any of the elements of a list (exercise_k) has
    been filled by the student. This means, the student is submitting this
    Teil(k+1) also.This function returns 0 if 'All None' is not true (i.e.: one
    element is different than None).

    +exercise_check_submitted(TeilInit, TeilEnd, solutions)
    This function checks which exercises have been submitted. The result
    of this function is a list with the indexes of the submitted exercises. The
    real submitted exercises would be 'i + 1', since Python starts lists with
    index 0.

    +corrector_is_there(path):
    This functions checks that a given 'path' there is a file named
    corrector.xlsx. It is not inteded to check if the file fulfills the needs
    for a complete correction.

    +corrector_checkout(path):
    This function takes the "path" of the corrector file uploaded by a given
    Prof. and checks that everything is ready for correction! Four main checks
    are carried out: 1. Name of the subject is specified, 2. Submission
    deadline has not expired still, 3. Successful log in to the specified
    e-mail address, and 4. Maximum number of attempts is over 1. This function
    returns a flag called "flag_ALL" which checks that all previous flags have
    been considered positively.

    +corrector_ready(path):
    This function combines 'corrector_is_there()' and 'corrector_checkout()'
    to make sure there is a valid 'corrector.xlsx' file in a given folder
    'path'. Returns a flag, taking value 1 if everything is ready.

    +corrector_get_data(corrector_path):
    This function takes the 'path' of the corrector file and gets the most
    relevant information out of it: the e-mail username for the correction
    'usn', the corresponding password 'psw', the maximum number of attempts
    allowed 'TrialsMax', where the exercises start 'TeilInit' and finish
    'TeilEnd', and the tolerable error for each variable 'Cerror'.

    +exerinfo_copy(MatNum, dummies, corr_path):
    This function's primary goal is to copy the matriculation number
    'MatNum' and the 'dummies' to the corrector file, given by the path
    'corr_path'.

    +correction(Sub_dir, exer_path, corr_path, usn, psw, eaddress):
    This function is the kernel of PyCor. It takes the task submitted
    (given by the path 'exer_path') and the corrector file (given by the path
    'corr_path') and compares the submitted exercises. It depends on other
    functions located in other libraries. It else requests the base folder
    'Sub_dir' in order to store the students stats, the username 'usn' and
    password 'psw' of the e-mail which warns the students and the students
    e-mail address 'eaddress'.

"""
import datetime
import logging
import os
import traceback

import numpy as np
import pywintypes
import simplecrypt
import win32com.client

import email_functions as emf
import verification_functions as vf

excel = win32com.client.Dispatch("Excel.Application")
log = logging.getLogger('PyCor')


# Functions related to the submitted 'exercise':

# def restart():
#
#     try:
#         os.system("taskkill /im EXCEL.EXE")
#         excel = client.Dispatch("Excel.Application")
#         print("Restarted EXCEL.EXE.")
#     except Exception:
#         print("EXCEL.EXE was already closed.")
#
#     return


def check_exercise(path):
    try:
        wb = excel.Workbooks.Open(path)
        ws1 = wb.Worksheets(1)
        MatNum = ws1.Range("B3").Value
        wb.Close(SaveChanges=False)
        print("MatNum: ", MatNum)
        print("The submitted file is successfully opened and closed.")
        flag = 1
    except Exception:
        print("The submitted file cannot be opened and/or closed.")
        flag = 0

    return flag


def exercise_get_data(path):
    """This function takes the 'path' to the submitted file by the student.
    Once opened, read the Matr. Num. and 'dummies', automatically detects the
    number of Teils and gets the initial and ending indexes. Also, reads the
    solutions provided by the student which can be later compared to the
    corrector."""

    # Open the Workbook and worksheet

    wb = excel.Workbooks.Open(path)
    ws1 = wb.Worksheets(1)  # Index = 1

    # Get the data we need to copy to the corrector

    MatNum = ws1.Range("B3").Value

    if MatNum == None:  # Someone forgot his Mat Num
        MatNum = 0

    dummies = ws1.Range("A6:CV6").Value  # Here, no need for: .Value

    # Exercises structure information

    # TotalTeils = 0
    index_num = 16
    index = 'A' + str(index_num)
    TeilInit = []
    TeilEnd = []
    solutions = []
    cell_previous = 0

    flag = 0

    while flag == 0:
        cell_active = ws1.Range(index).Value

        if cell_active > cell_previous:
            TeilInit.append(index)
            TeilEnd.append('A' + str(index_num - 1))
        elif cell_active == cell_previous:
            pass
        else:
            flag = 1
            TeilEnd.append('A' + str(index_num - 1))
            TeilEnd = TeilEnd[1::]
            break

        solutions.append(ws1.Range('C' + str(index_num)).Value)
        cell_previous = cell_active
        index_num = index_num + 1
        index = index[0:1] + str(index_num)

    # Close the Workbook before exit

    wb.Close(SaveChanges=False)

    return MatNum, dummies, TeilInit, TeilEnd, solutions


def check_allNone(exercise_k):
    """This function checks if any of the elements of a list (exercise_k) has
    been filled by the student. This means, the student is submitting this
    Teil(k+1) also. This function returns 0 if 'All None' is not true (i.e.:
    one element is different than None)."""

    flag = 1  # 1, for all None.

    n = len(exercise_k)

    for i in range(0, n):
        if exercise_k[i] == '-':
            pass
        elif exercise_k[i] is None:
            flag = 0  # allNone is false! One item is at least filled.

    return flag


def exercise_check_submitted(TeilInit, TeilEnd, solutions):
    """This function checks which exercises have been submitted. The result
    of this function is a list with the indexes of the submitted exercises. The
    real submitted exercises would be 'i + 1', since Python starts lists with
    index 0."""

    nI = len(TeilInit)

    for i in range(0, nI):
        TeilInit[i] = int(TeilInit[i][1::])
        TeilEnd[i] = int(TeilEnd[i][1::])

    offset = TeilInit[0]
    TeilInit = np.asarray(TeilInit) - offset
    TeilEnd = np.asarray(TeilEnd) - offset

    submitted = []

    for i in range(0, nI):
        exercise_i = solutions[TeilInit[i]:(TeilEnd[i] + 1)]

        # Now, we must check if there is a single value != None

        flag = check_allNone(exercise_i)

        if flag == 0:
            submitted.append(i)

    return submitted


# Functions related to the 'corrector' file:
def get_corrector_name(path):
    """This functions checks that a given 'path' there is a file named
    corrector.xlsx. It is not intended to check if the file fulfills the needs
    for a complete correction, but may check that the important cells are
    accessible at least."""

    items = os.listdir(path)
    extensions = ['xlsx', 'xlsm', 'xls']

    for item in items:
        corr_path = os.path.join(path, item)
        if os.path.isfile(corr_path) and any([item.endswith(x) for x in extensions]):
            return corr_path

    log.warning("corrector.xls* has not been found in the folder.")
    return


def check_corrector(path, password):
    # type: (str, str) -> bool
    """
    This function takes the "path" of the corrector file uploaded by a given
    Prof. and checks that everything is ready for correction! Four main checks
    are carried out:
    1. Name of the subject is specified,
    2. Submission deadline has not expired still,
    3. Successful log in to the specified e-mail address, and
    4. Maximum number of attempts is over 1.
    This function returns a flag called "flag_ALL" which checks that all previous flags have
    been considered positively.
    """

    # Open the workbook
    wb = excel.Workbooks.Open(path, True, False, None, password)
    ws1 = wb.Worksheets(1)

    try:
        subject_name = ws1.Range('B11').Value

        day, month, year = (int(x) for x in ws1.Range('D9:B9').Value[0])
        deadline = datetime.date(year, month, day)

        email = ''.join(ws1.Range('B11:C11').Value[0]).replace(' ', '')
        email_password = ws1.Range('B12').Value

        max_tries = ws1.Range('E12').Value

        wb.Close(SaveChanges=False)

        # Check for valid subject name
        if not subject_name:
            log.error('Empty subject field in %s. Please specify a valid name.', path)
            return False

        # Check deadline, allow for same-day submissions
        if (deadline - datetime.date.today()).days < 0:
            log.info('Ignoring %s due to deadline (%s)', path, deadline)
            return False

        if emf.check_login(email, email_password) != 1:
            log.error('Invalid login data for %s', path)
            return False

        if max_tries < 1:
            log.error('Invalid number of max tries')
            return False

        return True
    except pywintypes.com_error:
        traceback.print_exc()
        return False


def corrector_ready(path):
    """This function combines 'corrector_is_there()' and 'corrector_checkout()'
    to make sure there is a valid 'corrector.xlsx' file in a given folder
    'path'. Returns a flag, taking value 1 if everything is ready."""
    # Grab password
    psw = ''
    try:
        psw_file = open(os.path.join(path, 'psw'), 'r')
        psw_enc = psw_file.read()
        psw_file.close()
        psw = simplecrypt.decrypt('password', psw_enc)
    except IOError:
        pass
    except simplecrypt.DecryptionException:
        log.error('Failed to decrypt psw for %s', path)
        log.error(traceback.format_exc())
        return False, '', ''

    # Check that corrector.xlsx is there...
    corr_path = get_corrector_name(path)

    # ...and check that its content is adequate!
    if corr_path and check_corrector(corr_path, psw):
        return corr_path, psw
    return None, psw


def corrector_get_data(corrector_path, psw):
    """This function takes the 'path' of the corrector file and gets the most
    relevant information out of it: the e-mail username for the correction
    'usn', the corresponding password 'psw', the maximum number of attempts
    allowed 'TrialsMax', where the exercises start 'TeilInit' and finish
    'TeilEnd', and the tolerable error for each variable 'Cerror'."""

    # Hey, open the wb! and then the ws..

    wb = excel.Workbooks.Open(corrector_path, True, False, None, psw)
    ws1 = wb.Worksheets(1)  # Index = 1

    # Name of the corrected subject:

    Subject = ws1.Range("B1").Value

    # Associated e-mail address:

    email_p1 = ws1.Range("B11").Value
    email_p2 = ws1.Range("C11").Value
    usn = email_p1 + email_p2  # p1: username, p2: @domain.
    usn = usn.replace(' ', '')  # Delete the spaces apearing!
    psw = ws1.Range("B12").Value

    # Maximum number of trials:

    TrialsMax = ws1.Range("E12").Value

    # Exercises structure: TeilInit, TeilEnd

    index_num = 16
    index = 'A' + str(index_num)
    TeilInit = []
    TeilEnd = []
    solutions = []
    Cerror = []
    cell_previous = 0

    flag = 0

    while flag == 0:

        cell_active = ws1.Range(index).Value

        if cell_active > cell_previous:
            TeilInit.append(index)
            TeilEnd.append('A' + str(index_num - 1))
        elif cell_active == cell_previous:
            pass
        else:
            flag = 1
            TeilEnd.append('A' + str(index_num - 1))
            TeilEnd = TeilEnd[1::]
            break

        solutions.append(ws1.Range('C' + str(index_num)).Value)
        Cerror.append(ws1.Range('D' + str(index_num)).Value)

        TotalTeils = cell_previous
        cell_previous = cell_active
        index_num = index_num + 1
        index = index[0:1] + str(index_num)

    # Close the Workbook before exit

    wb.Close(SaveChanges=False)

    return Subject, usn, psw, TrialsMax, TeilInit, TeilEnd, TotalTeils, Cerror


def exerinfo_copy(MatNum, dummies, corr_path, psw_corr):
    """This function's primary goal is to copy the matriculation number
    'MatNum' and the 'dummies' to the corrector file, given by the path
    'corr_path'."""

    # Hey, open the wb! and the ws..

    wb = excel.Workbooks.Open(corr_path, True, False, None, psw_corr)
    ws1 = wb.Worksheets(1)  # Index = 1

    ws1.Range("B3").Value = MatNum
    ws1.Range("A6:CV6").Value = dummies

    # Maximum number of trials:

    TrialsMax = ws1.Range("E12").Value

    # Get the data...

    index_num = 16
    index = 'A' + str(index_num)
    Teil = []
    TeilInit = []
    TeilEnd = []
    solutions = []
    varNames = []
    Cerror = []
    cell_previous = 0

    flag = 0

    while flag == 0:
        cell_active = ws1.Range(index).Value

        if cell_active > cell_previous:
            TeilInit.append(index)
            TeilEnd.append('A' + str(index_num - 1))
        elif cell_active == cell_previous:
            pass
        else:
            flag = 1
            TeilEnd.append('A' + str(index_num - 1))
            TeilEnd = TeilEnd[1::]
            break

        Teil.append(ws1.Range('A' + str(index_num)).Value)
        varNames.append(ws1.Range('B' + str(index_num).Value))

        # print((varNames))

        solutions.append(ws1.Range('C' + str(index_num)).Value)
        Cerror.append(ws1.Range('D' + str(index_num)).Value)

        cell_previous = cell_active
        index_num = index_num + 1
        index = index[0:1] + str(index_num)

    # Close, but do not save.

    wb.Close(SaveChanges=False)

    return Teil, varNames, solutions, Cerror, TrialsMax


def correction(Sub_dir, exer_path, corr_path, usn, psw, eaddress, Subject,
               psw_corr, TotalTeils):
    """This function is the kernel of PyCor. It takes the task submitted
    (given by the path 'exer_path') and the corrector file (given by the path
    'corr_path') and compares the submitted exercises. It depends on other
    functions located in other libraries. It else requests the base folder
    'Sub_dir' in order to store the students stats, the username 'usn' and
    password 'psw' of the e-mail which warns the students and the students
    e-mail address 'eaddress'."""

    print("Getting most relevant data from student's submitted task.")
    [MN, d, exer_TI, exer_TE, exer_sol] = exercise_get_data(exer_path)
    print("Checking which exercises have been submitted from within the task.")
    subm = exercise_check_submitted(exer_TI, exer_TE, exer_sol)

    # Let's copy the MatNum and dummies to the corrector.
    print("Copying MatNum and dummies to corrector.xlsx")
    [Teil, varNames, solutions, Cerror, TrialsMax] = \
        exerinfo_copy(MN, d, corr_path, psw_corr)

    print("Checking the submitted exercises.")

    Results_subj = ''
    Results_msg = ''

    send_passed = 0
    j_passed = []

    send_block = 0
    j_blocked = []

    for i in subm:  # Indeed, number of Teil is 'j = i + 1'.

        j = i + 1

        print("Correcting exercise: ", j)

        # We may check if exercise is blocked or passed, to avoid unnecessary
        # correction.

        flag_blocked = vf.is_blocked(Sub_dir, j, TrialsMax)
        flag_passed = vf.is_passed(Sub_dir, j)

        if flag_blocked == 0 and flag_passed == 0:
            """ If not blocked already and not passed."""

            # 'Teil' includes the number of exercise. Only when j == Teil,
            # some correction is finally performed.
            print("Not blocked nor passed previously. ")

            r, names, resol_filtered = vf.correct_Teil(Teil, solutions, Cerror,
                                                       exer_sol, j, varNames,
                                                       usn, psw, eaddress,
                                                       Subject)

            # Let's send the e-mail with results!

            print("Sending e-mail with correction results for exercise: ", j)

            # 'html' format:
            [R_subj, R_msg] = emf.generatehtmlmsg_Results(j, names,
                                                          resol_filtered,
                                                          Subject)
            # Results_subj =
            Results_msg = Results_msg + R_subj + '<br>' + R_msg + '<br>'

            # emf.send_email(usn, psw, eaddress, R_subj, R_msg)

            # Once corrected, let's add this value to the student's history.
            print("Updating student's stats with new results.")

            [flag_blocked, flag_passed] = \
                vf.update_stats(Sub_dir, j, r, usn, psw, eaddress, TrialsMax,
                                Subject, MN)

            # Did we get anything to notify after 'update_stats'? Sure we did.

            if flag_blocked == 1:
                print("Student has been blocked exercise: ", j)

                send_block = 1
                j_blocked.append(j)

                # Send e-mail to warn the student about his Exercise 'j' block

                # [B_subj, B_msg] = emf.generatemsg_blockedTeil_i(j, Subject,
                #                                                 TrialsMax)
                # emf.send_email(usn, psw, eaddress, B_subj, B_msg)

            if flag_passed == 1:
                send_passed = 1

                # -------------------------------------------------------------

                print("Student has passed exercise: ", j)
                # Save this Teil, for future notifications..
                j_passed.append(j)

        elif flag_blocked == 1:

            send_block = 1
            j_blocked.append(j)

            """This case only corresponds to previously blocked. If the
            student is blocked this time, he will be notified after
            'vf.update_stats()'. So, only if the Exercise is not corrected, we
            will end in this place of the code"""

            # [B_subj, B_msg] = emf.generatemsg_blockedTeil_i(j, Subject,
            #                                                 TrialsMax)
            # emf.send_email(usn, psw, eaddress, B_subj, B_msg)

    """
    Here, we have already finished with all the 'subm' Teils. We can
    generate a single message with all blockings and all results. We also need
    something to know when there is a message for:
        -html results (always)
        -blocked Teil (when flag_blocked == 1 at least once) send_block == 1
        -passed Teil (when flag_passed == 1 at least once) send_passed == 1

        +extra: ALL passed. send_final == 1
    """

    # Let's generate Results e-mail!

    Results_subj = 'Ergebnisse: ' + Subject
    emf.send_email(usn, psw, eaddress, Results_subj, Results_msg)

    # Let's generate "passed exercises" message!

    if send_passed == 1:
        # Then, we have at least 1 passed exercise..
        print("Sending passed congratulations!")

        [P_subj, P_msg] = emf.generatemsg_passedTeil_i(j_passed, Subject, MN)
        emf.send_email(usn, psw, eaddress, P_subj, P_msg)

    if send_block == 1:
        # Then, we have at least 1 blocked exercise...
        print("Sending blockage e-mail.")

        [B_subj, B_msg] = emf.generatemsg_blockedTeil_i(j_blocked, Subject,
                                                        TrialsMax)
        emf.send_email(usn, psw, eaddress, B_subj, B_msg)

    flag_final = vf.check_final(Sub_dir, TotalTeils)  # Teil max.

    if flag_final == 1:
        [F_subj, F_msg] = emf.generatemsg_final(Subject, MN)
        emf.send_email(usn, psw, eaddress, F_subj, F_msg)

    return

# -----------------------------
# Functions testing
# -----------------------------

# corr_globalpath = 'C:\Users\Usuario\Dropbox\\2016 PyCor\
#     PyCor\Example corrector.xlsx'

# flag_start = corrector_checkout(corr_globalpath)

# exercise_globalpath = 'C:\Users\Usuario\Dropbox\\2016 PyCor\
#    PyCor\Example exercise.xlsx'

# [MatNum, dum, TeilInit, TeilEnd, sol] = exercise_get_data(
#   exercise_globalpath)
# submitted = exercise_check_submitted(TeilInit, TeilEnd, sol)
# print(submitted)

# excel.Quit()
