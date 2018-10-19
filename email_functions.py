# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'email_functions' is a complementary sub-library containing different functions
created for the general PyCor super-library.

This module is intended to provide access to e-mail data (google and FH Aachen
accounts) and allow e-mail sending capabilities.

Functions:

    +choose_server(usn):
    Given an e-mail 'usn', detects the e-mail termination and provides with the
    correct server information (both IMAP and SMTP). Available options are:
    gmail and FH Aachen e-mail accounts. It is missing the information for FH
    Aachen SMTP server.

    +check_login(usn, psw):
    This function checks that connection is possible given a username 'usn'
    and a password 'psw'. It returns 'OK' or 'OFF' correspondingly as a
    variable named 'login_status'.

    +send_email(usn, psw, eaddress, msg, subj):
    This function sends an e-mail to 'eaddress' given a defined 'msg'.
    'msg' corresponds to the body text of the e-mail. 'subj' corresponds to the
    e-mail subject. 'msg' and 'subj' could be produced by an additional
    function in this sub-library; which might contemplate all the cases
    occurring during PyCor running. 'usn' and 'psw' correspond to the username
    and password of the sender account.

    +verify(eaddress, usn, psw):
    This function checks if 'eaddress' is a valid FH Aachen e-mail account
    and sends the corresponding warning e-mail to the student in a negative
    case. 'usn' and 'psw' are the send e-mail username and password. It returns
    a flag taking 0 when the student e-mail address is invalid and 1 when it is
    a valid one.

    +check_single_att(msg, eaddress, Sub_dir):
    This function is much more than a general check out. First, it checks
    that there is a valid attachment, with a readable Excel format. It also
    decodes the name in case it has special characters, avoiding some problems.
    Finally, it stores the file using a name format which allows to recognize
    the date of submission. It returns a flag with 0 for no valid attachment or
    several attachments; and 1 for a single valid attachment. It also returns
    the path of the saved file as 'file_path'. This allows future location for
    correction of this file.

    +check_INBOX(usn, psw, current_dir, corr_path, TrialsMax):
    This function is directly called from PyCor.py and triggers a chain of
    functions leading to the full correction of all the student tasks allocated
    in the INBOX of the e-mail address 'usn'. 'psw' corresponds to its password
    and this function requires also requires the base folder path defined as
    'current_dir' and the 'corr_path', defined as current_dir +
    'corrector.xlsx'.

    +generatemsg_wrongeaddress():
    This function provides with the right subject and message text for the
    case in which the student is sending an e-mail with a wrong e-mail
    address.

    +generatemsg_invalidattach():
    This function provides with the right subject and message text for the
    case in which the student is not sending a valid attachment.

    +generatemsg_passedTeil_i(Number):
    This function provides with the right subject and message text for the
    case in which the student passes an exercise (Teil). The number of the
    passed exercise is provided as 'Number'.

    +generatemsg_blockedTeil_i(Number):
    This function provides with the right subject and message text for the
    case in which the student gets blocked an exercise (Teil). The number of
    the blocked exercise is provided as 'Number'.

    +generatemsg_Results(Number, names, resol):
    This function takes the number of exercise 'Number', the name of the
    variables which have been corrected 'names' and the veredict of the
    corrector 'resol' and produces the subject 'subj' and the message 'msg' of
    the e-mail to be forwarded to the corresponding student.

"""

import email
import imaplib  # To receive emails.
import smtplib  # To send emails.
import sys
import time
import email.header
import email.mime.multipart
import email.mime.text

import excel_functions as exf
import folder_functions as ff

delay = 3  # s.


def choose_server(usn):
    """Given an e-mail 'usn', detects the e-mail termination and provides with the
    correct server information (both IMAP and SMTP). Available options are:
    gmail and FH Aachen e-mail accounts. It is missing the information for FH
    Aachen SMTP server."""

    if usn[-9:] == "gmail.com":

        mail_server = 'imap.gmail.com'
        smtp_server = 'smtp.gmail.com:465'

    elif usn[-12:] == 'fh-aachen.de':

        mail_server = 'mail.fh-aachen.de'
        smtp_server = 'mail.fh-aachen.de:587'  # This server should be
    # selected from FH info

    else:

        print("Not detected known mail server")
        mail_server = ''
        smtp_server = ''

    return mail_server, smtp_server


def check_login(usn, psw):
    """This function checks that connection is possible given a username 'usn'
    and a password 'psw'. It returns 'OK' or 'OFF' correspondingly as a
    variable named 'login_status'."""

    mail_server, smtp_server = choose_server(usn)

    try:

        M = imaplib.IMAP4_SSL(mail_server)
        M.login(usn, psw)
        login_status = 1  # "ON"
        print("Successfully logged in: ", usn)

        time.sleep(delay)

        M.logout()

        print("Successfully logged out: ", usn)

    except Exception:

        login_status = 0  # "OFF"
        print("Please, check the username and password. Something went wrong.")
        print("If username and password are correct, server is failing.")
        print("If error persists, contact admin.")
        # sys.exit()    # This is the appropiate way to handle an exception

    return login_status


def send_email(usn, psw, eaddress, subj, msg):
    """This function sends an e-mail to 'eaddress' given a defined 'msg'.
    'msg' corresponds to the body text of the e-mail. 'subj' corresponds to the
    e-mail subject. 'msg' and 'subj' could be produced by an additional
    function in this sub-library; which might contemplate all the cases
    occurring during PyCor running. 'usn' and 'psw' correspond to the username
    and password of the sender account."""

    # Preparing the message

    Message = email.mime.multipart.MIMEMultipart('alternative')
    Message['From'] = usn
    Message['To'] = eaddress
    Message['Subject'] = subj

    Message.attach(MIMEText(msg, 'html'))

    mail_server, smtp_server = choose_server(usn)

    try:

        server = smtplib.SMTP_SSL(smtp_server)
        server.login(usn, psw)
        server.sendmail(usn, eaddress, Message.as_string())

        time.sleep(delay)

        print("Sending e-mail to: ", eaddress)

        server.quit()

    except Exception:

        print("Unexpected error: unable to send email to ", eaddress)

    return


def verify(eaddress, usn, psw, Subject):
    """This function checks if 'eaddress' is a valid FH Aachen e-mail account
    and sends the corresponding warning e-mail to the student in a negative
    case. 'usn' and 'psw' are the send e-mail username and password. It returns
    a flag taking 0 when the student e-mail address is invalid and 1 when it is
    a valid one."""

    if eaddress[-12:] == 'fh-aachen.de':

        # Sub_dir = current_dir + eaddress  # current_dir already has '\'

        print("Correct e-mail address! Submission will be accounted.")
        # print("allocating files to directory: ", Sub_dir)

        flag = 1

    elif eaddress[:13] == 'mailer-daemon':

        print("OK, mailer-daemon sent this e-mail. Not to be considered.")

        flag = 0

    elif eaddress[:8] == 'no-reply':

        print("OK, no-reply sent this e-mail. Not to be considered.")

        flag = 0

    elif eaddress[:7] == 'noreply':

        print("OK, no-reply sent this e-mail. Not to be considered.")

        flag = 0

    else:

        W_subj, W_msg = generatemsg_wrongeaddress(Subject)
        send_email(usn, psw, eaddress, W_subj, W_msg)
        # Sub_dir = current_dir + '_Invalid' + '\\' + eaddress

        print("Wrong e-mail address. Student has been warned.")
        # print("allocating files to directory: ", Sub_dir)

        flag = 0

    return flag


def check_single_att(msg, eaddress, Sub_dir):
    """This function is much more than a general check out. First, it checks
    that there is a valid attachment, with a readable Excel format. It also
    decodes the name in case it has special characters, avoiding some problems.
    Finally, it stores the file using a name format which allows to recognize
    the date of submission. It returns a flag with 0 for no valid attachment or
    several attachments; and 1 for a single valid attachment. It also returns
    the path of the saved file as 'file_path'. This allows future location for
    correction of this file."""

    cont_valids = 0

    # i_part = -1
    for part in msg.walk():  # Walk through the content of the email
        # i_part = i_part + 1
        # multipart are just containers, so we skip them

        # print("e-mail content type: ", part.get_content_maintype())
        # print(part)
        if part.get_content_maintype() == 'multipart':
            continue

        if part.get_content_maintype() == 'text':
            continue

        if part.get_content_maintype() == 'application':
            print("There is one submitted application.")

            submitted_file = part.get_payload(decode=True)
        # is this part an attachment?

        if part.get('Content-Disposition') is None:
            continue

            # Dealing with the file name that we receive and the one which may
            # be used at saving..

        filename = part.get_filename()  # This is the name of the part

        try:
            if email.header.decode_header(filename)[0][1] is not None:
                filename = str(email.header.decode_header(filename)[0][0]).decode(
                    email.header.decode_header(filename)[0][1])
                print("New name:")
                print(filename)
        except Exception:
            print("Didnt decode anything in: ")
            print(filename)

        try:  # If the file doesn't have any extension

            extension = filename[-3:]  # knowing the extension may be useful

        except Exception:  # Let's save whatever has been sent as a txt file.

            extension = 'txt'

        # Count number of times there is this extension in the attachments.

        if extension == 'lsx' in filename[-6:]:
            extension = 'xlsx'
            cont_valids += 1
            # i_excel = i_part

        elif extension == 'lsm' in filename[-6:]:
            extension = 'xlsm'
            cont_valids += 1
            # i_excel = i_part

        elif 'xlsx' in filename[-6:]:
            extension = 'xlsx'
            cont_valids += 1
            # i_excel = i_part

        elif 'xlsm' in filename[-6:]:
            extension = 'xlsm'
            cont_valids += 1
            # i_excel = i_part

        elif 'xls' in filename[-6:]:
            extension = 'xls'
            cont_valids += 1
            # i_excel = i_part

    print("Number of excel files attached: ", cont_valids)

    file_path = ''

    if cont_valids == 1:

        flag = 1

        # Save the file in "Sub_dir"

        time_att = time.localtime()
        file_path = (Sub_dir + '/' + str(time_att[0]) + str(time_att[1]) +
                     str(time_att[2]) + '_' + str(time_att[3]) +
                     str(time_att[4]) + str(time_att[5]) + '.' + extension)

        print("Saving the attachment.")
        fp = open(file_path, 'wb')
        fp.write(submitted_file)  # part.get_payload(decode=True))
        fp.close()
        print("Valid attachment! File has been saved for correction.")
        print("Checking that the saved file is a valid Excel file.")

        flag = exf.check_exercise(file_path)

    # -------------------------------------------------------------------------
    else:

        flag = 0

    print("Leaving the attachment checking routines..")
    return flag, file_path


def notify_no_attach(usn, psw, eaddress, Subject):
    [I_subj, I_msg] = generatemsg_invalidattach(Subject)

    send_email(usn, psw, eaddress, I_subj, I_msg)

    print("No file has been saved.")
    print("Invalid attachment. Student has been warned.")

    return


def check_INBOX(usn, psw, current_dir, corr_path, Subject, psw_corr,
                TotalTeils):
    """This function is directly called from PyCor.py and triggers a chain of
    functions leading to the full correction of all the student tasks allocated
    in the INBOX of the e-mail address 'usn'. 'psw' corresponds to its password
    and this function requires also requires the base folder path defined as
    'current_dir' and the 'corr_path', defined as current_dir +
    'corrector.xlsx'."""

    mail_server, smtp_server = choose_server(usn)

    try:

        M = imaplib.IMAP4_SSL(mail_server)
        M.login(usn, psw)

        print("Successfully logged in: ", usn)

        INBOXstatus, UnseenInfo = M.status('INBOX', "(UNSEEN)")
        INBOXcounter = int(UnseenInfo[0].split()[2].strip(').,]'))

        print("INBOX Status: ", INBOXstatus)
        print("Number of unread messages: ", INBOXcounter)

        M.select('INBOX')
        UNSEENstatus, UNSEENdata = M.search(None, 'UNSEEN')

        # Let's check now the e-mails, one-by-one -----------------------------

        for i in UNSEENdata[0].split():

            status, data = M.fetch(i, '(RFC822)')
            msg = email.message_from_string(data[0][1])
            # Printing general info:

            print('\nMessage number %s: %s' % (i, msg['Subject']))
            print('Student: ', msg['From'])

            # Checking e-mail address! ----------------------------------flag_1
            eaddress = str(msg['From'].split()[-1])[1:-1]

            flag_1 = verify(eaddress, usn, psw, Subject)

            if flag_1 == 1:  # if the e-mail address exists, create a folder

                Sub_dir = current_dir + eaddress  # current_dir already has '\'
                ff.create_folder(Sub_dir)

                # -----------------------------------------------------------------
                # Checking only 1 attachment --------------------------------flag_2

                [flag_2, exer_path] = check_single_att(msg, eaddress, Sub_dir)

                if flag_2 == 0:
                    print("Notifying student due to not valid attachment...")
                    notify_no_attach(usn, psw, eaddress, Subject)

            else:

                flag_2 = 0

            # -----------------------------------------------------------------

            # We can now start correction!

            if flag_2 == 1:
                print("Correcting exercise: ", exer_path)

                exf.correction(Sub_dir, exer_path, corr_path, usn, psw,
                               eaddress, Subject, psw_corr, TotalTeils)

                time.sleep(delay)

        M.close()
        M.logout()

    except Exception:

        print("PyCor could not check INBOX of: ", usn)
        print("Please, check the username and password. Something went wrong.")
        print("If username and password are correct, server is failing.")
        print("If error persists, contact admin.")

        sys.exit()  # This is the appropiate way to handle an exception

    return


def generatemsg_wrongeaddress(SubjectCorr):
    """This function provides with the right subject and message text for the
    case in which the student is sending an e-mail with a wrong e-mail
    address."""

    Subject = open("_AutoReply\A1_s.txt", "r")
    Message = open("_AutoReply\A1.txt", "r")
    subj = Subject.read()
    msg = Message.read()
    Subject.close()
    Message.close()

    signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'
    # print(type(signature))
    msg = msg.replace('PyCor', signature)
    msg = str((msg.decode('utf-8')).encode('ascii', 'xmlcharrefreplace'))

    return subj, msg


def generatemsg_invalidattach(SubjectCorr):
    """This function provides with the right subject and message text for the
    case in which the student is not sending a valid attachment."""

    Subject = open("_AutoReply\A2_s.txt", "r")
    Message = open("_AutoReply\A2.txt", "r")
    subj = Subject.read()
    msg = Message.read()
    Subject.close()
    Message.close()

    signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'

    msg = msg.replace('PyCor', signature)
    # print("Let's encode now.")
    msg = str((msg.decode('utf-8')).encode('ascii', 'xmlcharrefreplace'))
    # print("no problem now.")

    return subj, msg


def generatemsg_passedTeil_i(Number, SubjectCorr, MatNum):
    """This function provides with the right subject and message text for the
    case in which the student passes an exercise (Teil). The number of the
    passed exercise is provided as 'Number'."""

    Subject = open("_AutoReply\A3_s.txt", "r")
    Message = open("_AutoReply\A3.txt", "r")
    subj = Subject.read()
    subj = subj + ' Mat. Num.: ' + str(int(MatNum))
    msg = Message.read()
    Subject.close()
    Message.close()

    # Let's change Teil.No for the given Number!

    # subj = subj.replace('TeilNo', str(Number))
    # for j in Number:
    #     msg = msg.replace('Teil(e) ', 'Teil(e) ' + str(j) + ' ')
    msg = msg.replace('Teil(e) ', 'Teil(e) ' + str(Number) + ' ')

    signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'
    msg = msg.replace('PyCor', signature)
    msg = str((msg.decode('utf-8')).encode('ascii', 'xmlcharrefreplace'))

    return subj, msg


def generatemsg_blockedTeil_i(Number, SubjectCorr, TrialsMax):
    """This function provides with the right subject and message text for the
    case in which the student gets blocked an exercise (Teil). The number of
    the blocked exercise is provided as 'Number'."""

    Subject = open("_AutoReply\A4_s.txt", "r")
    Message = open("_AutoReply\A4.txt", "r")
    subj = Subject.read()
    msg = Message.read()
    Subject.close()
    Message.close()

    # Let's change Teil.No for the given Number!

    # subj = subj.replace('TeilNo', str(Number))
    msg = msg.replace('TeilNo', str(Number))
    msg = msg.replace('TrialsMax', str(int(TrialsMax)))

    signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'
    msg = msg.replace('PyCor', signature)
    msg = str((msg.decode('utf-8')).encode('ascii', 'xmlcharrefreplace'))

    return subj, msg


def generatemsg_final(SubjectCorr, MatNum):
    """This function provides with the final congratulating e-mail message and
    subject."""

    Subject = open("_AutoReply\A6_s.txt", "r")
    Message = open("_AutoReply\A6.txt", "r")
    subj = Subject.read()
    subj = subj + ' Mat. Num.: ' + str(int(MatNum))
    msg = Message.read()
    Subject.close()
    Message.close()

    # Let's change Teil.No for the given Number!

    signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'
    msg = msg.replace('PyCor', signature)
    msg = str((msg.decode('utf-8')).encode('ascii', 'xmlcharrefreplace'))

    return subj, msg


def generatehtmlmsg_Results(Number, names, resol, SubjectCorr):
    # subj = "Results: exercise " + str(Number)
    subj = "Detaillierte Ergebnisse: <b>Teilaufgabe " + str(Number) + "</b>"

    Message = open("_AutoReply\A5.txt", "r")
    msg = Message.read()
    Message.close()

    nt = len(resol)
    resultsPyCor = ''

    for i in range(0, nt):

        if resol[i] == 0:

            resultsPyCor = (resultsPyCor + '<tr><td>' +
                            names[i].encode('ascii', 'xmlcharrefreplace') + '</td>' +
                            '<td>' + '<font color="red">' + 'Falsch' +
                            '</font>' + '</td></tr>')
            # print(resultsPyCor)
        else:
            resultsPyCor = (resultsPyCor + '<tr><td>' +
                            names[i].encode('ascii', 'xmlcharrefreplace') + '</td>' +
                            '<td>' + '<font color="green">' + 'Richtig'
                            + '</font>' + '</td></tr>')
            # print(resultsPyCor)
        # resultsPyCor = resultsPyCor + '<tr><td>' + names[i] + '</td>' + \
        #                               '<td>' + str(resol[i]) + '</td></tr>'

    # Inserting results:
    msg = msg.replace('PrintResults', str(resultsPyCor))
    # Inserting Subject name:
    # signature = '<b>' + SubjectCorr + '</b>' + ' und PyCor'
    signature = ''
    msg = msg.replace('PyCor', signature)

    return subj, msg


def generatemsg_Results(Number, names, resol, usn, psw, eaddress):
    """This function takes the number of exercise 'Number', the name of the
    variables which have been corrected 'names' and the veredict of the
    corrector 'resol' and produces the subject 'subj' and the message 'msg' of
    the e-mail to be forwarded to the corresponding student."""

    subj = "Results: exercise " + str(Number)
    msg = []

    # resol = np.where(resol == 1.0, 'correct', resol)
    # resol = np.where(resol == 0.0, 'INcorrect', resol)

    # count = 0
    nt = len(resol)
    for i in range(0, nt):
        msg.append(names[i] + ': ' + str(resol[i]) + '\n')

    msg = str(msg)

    # Preparing the e-mail.

    Message_R = email.mime.text.MIMEText(msg)
    Message_R['From'] = usn
    Message_R['To'] = eaddress
    Message_R['Subject'] = subj

    mail_server, smtp_server = choose_server(usn)

    try:

        server = smtplib.SMTP_SSL(smtp_server)
        server.login(usn, psw)
        server.sendmail(usn, eaddress, Message_R.as_string())

        time.sleep(delay)

        print("Sending e-mail to: ", eaddress)

        server.quit()

    except Exception:

        print("Unexpected error: unable to send email to ", eaddress)

    return

# -----------------------------
# Functions testing
# -----------------------------

# check_login(usn, psw)

# eaddress = 'davahue@gmail.com'
# subj = 'Hello Daniel'
# msg = 'This is one automatic e-mail!'

# send_email(usn, psw, eaddress, subj, msg)

# Number = 3
# generatemsg_blockedTeil_i(Number)
