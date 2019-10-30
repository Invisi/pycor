import csv
import logging
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

import utils

config = utils.import_config()


class PostProcessing:
    def __init__(self, subject_folder: Path, exercise_count: int):
        self.subject_folder = subject_folder
        self.log = logging.getLogger("PyCor").getChild("PostProcessing")
        self.exercise_count = exercise_count
        self.post_dir = subject_folder / "_postprocessing"

        # Create postprocessing folder
        if not self.post_dir.exists():
            self.post_dir.mkdir(exist_ok=True)

    def filter_folders(self):
        for folder in self.subject_folder.iterdir():
            # Ignore folders that are blacklisted or don't contain @
            if (
                not folder.is_dir()
                or folder.name in config.FOLDER_IGNORE
                or "@" not in folder.name
            ):
                continue
            yield folder

    @staticmethod
    def get_mat_num(folder: Path):
        try:
            # dtype has to be float since old data contained values too long for C long
            # noinspection PyTypeChecker
            mn = np.loadtxt(
                folder / "data" / "mat_num.txt",
                delimiter=" - ",
                usecols=1,
                dtype=float,
                ndmin=1,
            )
            return int(mn[-1]), np.unique(mn).size
        except IOError:
            return "", ""

    @staticmethod
    def load_txt(folder, ex):
        return np.loadtxt(
            folder / "data" / "Exercise{}.txt".format(ex + 1),
            delimiter=" - ",
            usecols=1,
            dtype=int,
            ndmin=1,
        )

    def write_csv(self, rows, name):
        comma_file = self.post_dir / "{}.csv".format(name)
        try:
            with comma_file.open("w", newline='') as c:
                c.write("sep=,\r\n")
                writer = csv.writer(c)
                for row in rows:
                    writer.writerow(row)
            self.log.info("Wrote %s files", name)
        except IOError:
            self.log.exception("Failed to save csv files.")

    def generate_attempt_info(self):
        rows_general = [
            ["Student", "Matr. Num.", "No. Matr. Num. used"]
            + ["Exercise {}".format(x + 1) for x in range(self.exercise_count)]
        ]
        rows_attempts = list(rows_general)

        for folder in self.filter_folders():
            # Get amount of mat nums used and last num
            mat_num, mn_count = self.get_mat_num(folder)

            row_general = [folder.name, mat_num, mn_count]
            row_attempts = list(row_general)

            # Percentage solved/Amount of tries
            for ex in range(self.exercise_count):
                try:
                    result = self.load_txt(folder, ex)
                    perc = np.max(result)
                    amount = len(result)
                except IOError:
                    perc = ""
                    amount = ""
                row_general.append(perc)
                row_attempts.append(amount)

            rows_attempts.append(row_attempts)
            rows_general.append(row_general)

        self.write_csv(rows_attempts, "AttemptsInfo")
        self.write_csv(rows_general, "GeneralInfo")

    def check_mat_num(self):
        cheaters = []
        for folder in self.filter_folders():
            # Get amount of mat nums used and last num
            _, mn_count = self.get_mat_num(folder)
            if isinstance(mn_count, int) and mn_count > 1:
                cheaters.append(folder.name + "\n")

        cheater_file = self.post_dir / "cheaters.txt"
        with cheater_file.open("w") as c:
            c.write("List of students using several matriculation numbers:\n")
            c.writelines(cheaters)
        self.log.info("Wrote cheater file.")

    def generate_bars(self):
        if self.exercise_count < 10:
            bar_labels = [
                "Ex. {}".format(x + 1) for x in range(self.exercise_count + 1)
            ]
        else:
            bar_labels = [str(x + 1) for x in range(self.exercise_count + 1)]
        passed = np.zeros(self.exercise_count)
        submitted = np.zeros(self.exercise_count)

        for folder in self.filter_folders():
            for ex in range(self.exercise_count):
                try:
                    result = self.load_txt(folder, ex)
                    highest_score = np.max(result)

                    if highest_score == 100:
                        passed[ex] += 1
                        submitted[ex] += 1
                except IOError:
                    pass

        # Generate bar plots
        ind = np.arange(self.exercise_count)
        width = 0.4

        plt.clf()
        p1 = plt.bar(ind - width / 2, passed, width, color="lightgreen")
        p2 = plt.bar(ind + width / 2, submitted, width, color="lightcoral")
        plt.ylabel("Number of students")
        plt.yticks(ind)
        plt.xticks(ind, bar_labels)

        plt.ylim(top=np.max(submitted + 1))
        plt.legend((p1[0], p2[0]), ("Passed", "Submitted"))
        plt.grid(True)

        # Save plots
        bars_png = self.post_dir / "passed-submitted.png"

        try:
            plt.savefig(bars_png)
            self.log.info("Wrote bar plots")
        except IOError:
            self.log.exception("Failed to save bar plots.")

    def generate_histograms(self):
        # Collect data on amount of passed exercises and amount of tries
        total = []
        passed = []
        for idx, folder in enumerate(self.filter_folders()):
            total.append([])
            passed.append([])

            for ex in range(self.exercise_count):
                total[idx].append([])
                passed[idx].append([])

                try:
                    result = self.load_txt(folder, ex)
                    highest_score = np.max(result)
                    tries = result.size

                    total[idx][ex] = tries
                    if highest_score == 100:
                        passed[idx][ex] = tries
                    else:
                        passed[idx][ex] = 0
                except IOError:
                    total[idx][ex] = 0
                    passed[idx][ex] = 0

        # Prepare data for plot
        total = np.array(total)
        passed = np.array(passed)

        # Plot exercise
        for ex in range(self.exercise_count):
            # Ignore missing exercise data
            if len(total[:, ex]) >= 1:
                y_total = np.bincount(total[:, ex])
                y_passed = np.bincount(passed[:, ex])

                y_total[0] = 0
                y_passed[0] = 0

                for i in range(len(y_passed), len(y_total)):
                    y_passed = np.append(y_passed, 0)

                y_total = np.append(y_total, 0)
                y_passed = np.append(y_passed, 0)
                x = np.arange(0, y_total.size, 1)

                plt.clf()
                plt.plot(x, y_total, "k--")
                plt.fill_between(
                    x,
                    y_total,
                    y_passed,
                    where=y_total > y_passed,
                    facecolor="lightcoral",
                    interpolate=True,
                    label="Submitted",
                    zorder=2,
                )
                plt.plot(x, y_passed, "k-")
                plt.fill_between(
                    x,
                    y_passed,
                    0,
                    where=y_passed > 0,
                    facecolor="lightgreen",
                    interpolate=True,
                    label="Passed",
                    zorder=3,
                )

                plt.xlabel("Number of attempts")
                plt.xticks(np.arange(np.max(y_total)) + 1)
                plt.ylabel("Number of students")
                plt.yticks(np.arange(y_total.size))

                plt.xlim((-0.6, y_total.size))
                plt.ylim((0, np.max(y_total) + 1))

                title = "Distribution of the number of\nattempts per student, ex. {}".format(
                    ex + 1
                )
                plt.title(title)

                plt.legend(loc="best")
                plt.grid()

                label = "Exercise {}".format(ex + 1)
                bars_png = self.post_dir / (label + "_distr.png")
                try:
                    plt.savefig(bars_png)
                except IOError:
                    self.log.exception("Failed to save hist plots.")
        self.log.info("Generated exercise histograms.")

    def run(self):
        self.generate_attempt_info()
        self.check_mat_num()
        self.generate_bars()
        self.generate_histograms()
