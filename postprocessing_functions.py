# -*- coding: utf-8 -*-

"""
Copyright 2016, Daniel Valero and Daniel Bung, All rights reserved.

'postprocessing_functions' is a complementary sub-library containing different
functions created for the general PyCor super-library.

This module is intended to provide some functions which allow PyCor user access
the data and visualize it easily.

Functions:

    +check_MatNum(current_dir):
    This function checks which of the studetns have used different
    Matriculation Numbers by checking the number they used in all their
    submissions. Students are not notified, instead, FH Aachen personnel have
    access to a generated file '../_postprocessing/cheaters.txt' which the
    complete list of students.

"""

import csv
import os

import matplotlib.pyplot as plt
import numpy as np

import folder_functions as ff
import verification_functions as vf


# --------------------------------------------------------------- CSV functions


def generate_generalCSV(current_dir, jmax):
    print("Generating the GeneralInfo.csv")

    # Created the headers: ----------------------------------------------------
    printable = []
    printable_new = ['Student', 'Matr. Num.', 'No. Matr. Num. used']

    for j in range(0, jmax):
        printable_new.append('Exercise ' + str(j + 1))

    printable.append(printable_new)

    # -------------------------------------------------------------------------

    student_folders = os.listdir(current_dir)
    filtered_folders = ff.filter_folders(current_dir, student_folders)

    # -------------------------------------------------------------------------

    for dir in filtered_folders:

        # Save printable_new[0]

        printable_new = [dir]

        # r = np.zeros([jmax])
        MatNum_dir = current_dir + dir + '\\data\\' + 'MatNum.txt'

        # Check how many MatNums have been used.

        try:

            MN = np.loadtxt(MatNum_dir)
            nMN = vf.count_different(MN)
            MatNum = MN[-1]

        except Exception:

            nMN = 0
            MatNum = 0

        # Save printable_new[1]
        printable_new.append(MatNum)
        printable_new.append(nMN)

        # Get the results from all exercises

        for j in range(0, jmax):

            Teil_j_str = "Exercise" + str(j + 1)
            all_j_dir = current_dir + dir + '\\data\\' + Teil_j_str + "_all.txt"

            try:

                Teil_j_all = np.loadtxt(all_j_dir)
                # r[j] = np.max(Teil_j_all)
                nj = np.max(Teil_j_all)

            except Exception:

                # r[j] = ''
                nj = ''

            # printable_new.append(r[j])
            printable_new.append(nj)

        printable.append(printable_new)

    # print(printable)

    # Save this data as a CSV file:

    post_dir = current_dir + '_postprocessing'
    ff.create_folder(post_dir)

    # Create the 'coma' delimiter CSV

    csv_dir_coma = post_dir + '\\GeneralInfo_coma.csv'

    # Let's try to save it

    try:

        with open(csv_dir_coma, 'wb') as csvfile:

            spamwriter = csv.writer(csvfile, delimiter=',', dialect='excel')

            nk = len(printable)
            for k in range(0, nk):
                spamwriter.writerow(printable[k][:])

    except Exception:

        print("Could not save: ", csv_dir_coma)

    # Create the 'semicolon' delimiter CSV

    csv_dir_semicolon = post_dir + '\\GeneralInfo_semicolon.csv'

    # Let's try to save it:

    try:

        with open(csv_dir_semicolon, 'wb') as csvfile:

            spamwriter = csv.writer(csvfile, delimiter=';', dialect='excel')

            nk = len(printable)

            for k in range(0, nk):

                for m in range(0, len(printable[k][:])):

                    try:
                        printable[k][m] = int(np.floor(printable[k][m]))
                    except Exception:
                        pass

                that_row = printable[k][:]
                spamwriter.writerow(that_row)

    except Exception:

        print("Could not save: ", csv_dir_semicolon)

    return


def generate_attemptsCSV(current_dir, jmax):
    print("Generating the TotalAttempts.csv")

    # Created the headers: ----------------------------------------------------
    printable = []
    printable_new = ['Student', 'Matr. Num.', 'No. Matr. Num. used']

    for j in range(0, jmax):
        printable_new.append('Exercise ' + str(j + 1))

    printable.append(printable_new)

    # -------------------------------------------------------------------------

    student_folders = os.listdir(current_dir)
    filtered_folders = ff.filter_folders(current_dir, student_folders)

    # -------------------------------------------------------------------------

    for dir in filtered_folders:

        # Save printable_new[0]

        printable_new = [dir]

        MatNum_dir = current_dir + dir + '\\data\\' + 'MatNum.txt'

        # Check how many MatNums have been used.

        try:

            MN = np.loadtxt(MatNum_dir)
            nMN = vf.count_different(MN)
            MatNum = MN[-1]

        except Exception:

            nMN = 0
            MatNum = 0
        # Save printable_new[1]
        printable_new.append(MatNum)
        printable_new.append(nMN)

        # Get the results from all exercises

        for j in range(0, jmax):

            Teil_j_str = "Exercise" + str(j + 1)
            all_j_dir = current_dir + dir + '\\data\\' + Teil_j_str + "_all.txt"

            try:

                Teil_j_all = np.loadtxt(all_j_dir)
                nj = Teil_j_all.size

            except Exception:

                nj = ''

            printable_new.append(nj)

        printable.append(printable_new)

    # print(printable)

    # Save this data as a CSV file:

    post_dir = current_dir + '_postprocessing'
    ff.create_folder(post_dir)

    # Create the 'coma' delimiter CSV

    csv_dir_coma = post_dir + '\\AttemptsInfo_coma.csv'

    try:

        with open(csv_dir_coma, 'wb') as csvfile:

            spamwriter = csv.writer(csvfile, delimiter=',', dialect='excel')

            nk = len(printable)
            for k in range(0, nk):
                spamwriter.writerow(printable[k][:])

    except Exception:

        print("Could not save: ", csv_dir_coma)

    # Create the 'semicolon' delimiter CSV

    csv_dir_semicolon = post_dir + '\\AttemptsInfo_semicolon.csv'

    try:

        with open(csv_dir_semicolon, 'wb') as csvfile:

            spamwriter = csv.writer(csvfile, delimiter=';', dialect='excel')

            nk = len(printable)
            for k in range(0, nk):

                for m in range(0, len(printable[k][:])):

                    try:
                        printable[k][m] = int(np.floor(printable[k][m]))
                    except Exception:
                        pass

                that_row = printable[k][:]
                spamwriter.writerow(that_row)

    except Exception:

        print("Could not save: ", csv_dir_semicolon)

    return


# --------------------------------------------------------------- txt functions


def check_MatNum(current_dir):
    """This function checks which of the studetns have used different
    Matriculation Numbers by checking the number they used in all their
    submissions. Students are not notified, instead, FH Aachen personnel have
    access to a generated file '../_postprocessing/cheaters.txt' which the
    complete list of students."""

    print("Generating cheaters.txt")

    # Access each students folder (but not all folders are from stundets,
    # i.e.: ../_postprocessing)

    cheaters = ['List of students using several Matriculation Numbers:']

    student_folders = os.listdir(current_dir)
    filtered_folders = ff.filter_folders(current_dir, student_folders)

    for dir in filtered_folders:

        MatNum_dir = current_dir + dir + '\\data\\' + 'MatNum.txt'

        try:

            MN = np.loadtxt(MatNum_dir)
            flag = vf.check_MNvalid(MN)  # 0 for single value, 1 otherwise

            if flag == 1:
                cheaters.append(dir)

        except Exception:

            print("No MatNum in folder: ", current_dir + dir)

    # Create the postprocessing directory
    # Remember: current_dir = profs + "\\" + dir + "\\"

    post_dir = current_dir + '_postprocessing'
    cheaters_dir = post_dir + '\\cheaters.txt'
    ff.create_folder(post_dir)

    # Save text data

    try:
        thefile = open(cheaters_dir, 'w')

        for item in cheaters:
            thefile.write("%s\n" % item)

        thefile.close()

    except Exception:

        print("Could not save: ", cheaters_dir)

    # np.savetxt(cheaters_dir, cheaters)

    return


# ---------------------------------------------------------- Plotting functions


def bars_all(current_dir, jmax):
    print("Generating bars plot.")

    bars_labels = []

    for j in range(0, jmax):
        bars_labels.append('Ex. ' + str(j + 1))

    # Created the headers: ----------------------------------------------------

    passed = np.zeros(jmax)
    total = np.zeros(jmax)

    # -------------------------------------------------------------------------

    student_folders = os.listdir(current_dir)
    filtered_folders = ff.filter_folders(current_dir, student_folders)

    # -------------------------------------------------------------------------

    for dir in filtered_folders:

        # Get the results from all exercises

        for j in range(0, jmax):

            Teil_j_str = "Exercise" + str(j + 1)
            all_j_dir = current_dir + dir + '\\data\\' + Teil_j_str + "_all.txt"

            try:

                Teil_j_all = np.loadtxt(all_j_dir)
                MaxScore = np.max(Teil_j_all)

                if MaxScore == 100:

                    passed[j] += 1
                    total[j] += 1

                else:

                    total[j] += 1

            except Exception:

                pass

    # print(printable)

    # Save this data as a CSV file:

    post_dir = current_dir + '_postprocessing'
    ff.create_folder(post_dir)

    # Let's plot!

    # we already have 'bars_labels'

    plt.clf()

    ind = np.arange(jmax)
    width = 1.00

    p1 = plt.bar(ind, passed, width, color='lightgreen', zorder=3)
    p2 = plt.bar(ind, total, width, color='lightcoral', zorder=2)

    plt.ylabel('Number of students')

    plt.xticks(ind + width / 2., bars_labels)
    ymax = np.max(np.asarray(total))

    # plt.yticks(np.arange(0, ymax + 2, 1))
    plt.ylim((0, ymax + 1))
    plt.legend((p2[0], p1[0]), ('Submitted', 'Passed'))
    plt.grid()

    # plt.show()

    bars_dir_png = post_dir + '\\passed-submitted.png'
    bars_dir_svg = post_dir + '\\passed-submitted.svg'

    try:

        plt.savefig(bars_dir_png)
        plt.savefig(bars_dir_svg)

    except Exception:

        print("Could not save: ", bars_dir_png)

    plt.clf()

    return


def hist_submissions(current_dir, jmax):
    print("Generating histograms plot for number of submissions/exercise.")

    plot_labels = []

    for j in range(0, jmax):
        plot_labels.append('Exercise ' + str(j + 1))

    # -------------------------------------------------------------------------

    student_folders = os.listdir(current_dir)
    filtered_folders = ff.filter_folders(current_dir, student_folders)

    # -------------------------------------------------------------------------
    # Number of data points: --------------------------------------------------

    count = 0
    max_sub = 0

    for dir in filtered_folders:

        count = count + 1

        for j in range(0, jmax):

            Teil_j_str = "Exercise" + str(j + 1)
            all_j_dir = current_dir + dir + '\\data\\' + Teil_j_str + "_all.txt"

            try:

                Teil_j_all = np.loadtxt(all_j_dir)
                sub = Teil_j_all.size

                if max_sub < sub:
                    max_sub = sub

            except Exception:

                pass

    # Let's initialize the array where we will fill in the plotting data.

    total_data = np.empty((count, jmax))
    total_data.fill(np.nan)

    passed_data = np.empty((count, jmax))
    passed_data.fill(np.nan)

    i = 0

    for dir in filtered_folders:

        # Get the results from all exercises

        for j in range(0, jmax):

            Teil_j_str = "Exercise" + str(j + 1)
            all_j_dir = current_dir + dir + '\\data\\' + Teil_j_str + "_all.txt"

            try:

                Teil_j_all = np.loadtxt(all_j_dir)
                MaxScore = np.max(Teil_j_all)
                nj = Teil_j_all.size

                total_data[i, j] = nj

                if MaxScore == 100:

                    passed_data[i, j] = nj

                else:

                    passed_data[i, j] = 0

            except Exception:

                total_data[i, j] = 0
                passed_data[i, j] = 0

        i = i + 1

    # Let's prepare the data to plot.

    total_data = total_data.astype(int)
    passed_data = passed_data.astype(int)

    # Let's plot for all 'jmax' exercises.
    post_dir = current_dir + '_postprocessing'
    ff.create_folder(post_dir)

    for j in range(0, jmax):

        if len(total_data[:, j]) >= 1:

            y_total = np.bincount(total_data[:, j])
            y_pass = np.bincount(passed_data[:, j])

            # zeros to be discarded: someone who tried '0' times, has not tried

            y_total[0] = 0
            y_pass[0] = 0

            for i in range(len(y_pass), len(y_total)):
                y_pass = np.append(y_pass, 0)

            plt.clf()

            y_total = np.append(y_total, 0)
            y_pass = np.append(y_pass, 0)
            x = np.arange(0, y_total.size, 1)

            # width = 1.0

            plt.plot(x, y_pass, 'k--')
            plt.fill_between(x, y_total, y_pass, where=y_total > y_pass,
                             facecolor='lightcoral', interpolate=True,
                             label='Submitted', zorder=2)
            plt.plot(x, y_pass, 'k-')
            plt.fill_between(x, y_pass, 0, where=y_pass > 0,
                             facecolor='lightgreen', interpolate=True,
                             label='Passed', zorder=3)
            # p1 = plt.bar(x - 0.5, y_pass, width, color='lightgreen',
            #              edgecolor='k', zorder=3)
            # p2 = plt.bar(x - 0.5, y_total, width, color='lightcoral',
            #              edgecolor='k', zorder=2)

            xlabel_1 = 'Number of attempts'

            # plt.xticks(x + width/2., str(x))
            plt.xlabel(xlabel_1)
            plt.ylabel('Number of students')

            plt.xlim((-0.6, y_total.size))
            plt.ylim((0, np.max(y_total) + 1))

            title = 'Distribution of the number of\n '
            title = title + 'attempts per student, ex. ' + str(j + 1)

            plt.title(title)
            # plt.legend((p1[0], p2[0]), ('Passed', 'Submitted'))
            plt.legend(loc='best')

            plt.grid()

            bars_dir_png = post_dir + '\\' + plot_labels[j] + '_distr.png'
            bars_dir_svg = post_dir + '\\' + plot_labels[j] + '_distr.svg'

            try:

                plt.savefig(bars_dir_png)
                plt.savefig(bars_dir_svg)

            except Exception:

                print("Could not save: ", bars_dir_png)

    plt.clf()

    return


# ---------------------------------------------------------------- full_analyze


def full_analyze(current_dir, TotalTeils):
    jmax = int(TotalTeils)  # number of exercises.

    # ----------------------------------------------------------- CSV functions
    # Generate the general CSV file (stundets, passed Teils)
    print("____________")
    generate_generalCSV(current_dir, jmax)

    # Generate Total Attempts CSV file (students, all attempts)
    print("____________")
    generate_attemptsCSV(current_dir, jmax)

    # ----------------------------------------------------------- txt functions
    # Check who copied (different MatNum, same student)
    print("____________")
    check_MatNum(current_dir)

    # ------------------------------------------------------ Plotting functions

    # bars with all the passed/submitted exercises
    print("____________")
    bars_all(current_dir, jmax)
    hist_submissions(current_dir, jmax)

    return

# ----------------------------------------------------------- Functions testing

# path = 'C:\Users\Usuario\Dropbox\\2016 PyCor\PyCor\\folder 1\subject b\\'
# check_MatNum(path)
# jmax = 3
# generate_generalCSV(path, jmax)
# generate_attemptsCSV(path, jmax)
# bars_all(path, jmax)
# hist_submissions(path, jmax)
