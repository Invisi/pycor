import datetime
import logging
import os
import traceback
from pathlib import Path
from typing import Optional

import pywintypes
import simplecrypt
import win32com.client

import mail
import utils

config = utils.import_config()


class ExcelFileException(Exception):
    pass


class Excel:
    def __init__(self, excel_file: Path):
        self.log = logging.getLogger('PyCor').getChild('Excel')

        self.excel = win32com.client.Dispatch('Excel.Application')

        if config.SHOW_EXCEL:
            self.excel.Visible = True

        self.excel.DisplayAlerts = False  # "Do you want to save your work?"
        self.excel.AskToUpdateLinks = False  # Links = Copied values from another sheet, we might want to detect those
        self.excel_file = excel_file
        self.log.info('Opening %s', self.get_relevant_path())

        self.exercise_ranges = None
        self.solutions = None

    def set_exercise_rows(self, ws, solutions=False):
        offset = 16
        index_num = offset
        previous_exercise = 0
        exercise_row_begin = []
        exercise_row_end = []
        self.exercise_ranges = []

        if solutions:
            self.solutions = []

        while True:
            current_exercise = ws.Cells(index_num, 1).Value  # A{index_num}

            if not isinstance(current_exercise, float):
                exercise_row_end.append(index_num - 1)
                break

            current_exercise = int(current_exercise)

            if current_exercise > previous_exercise:
                if previous_exercise > 0:
                    exercise_row_end.append(index_num - 1)
                exercise_row_begin.append(index_num)

            if solutions:
                if len(self.solutions) <= current_exercise - 1:
                    self.solutions.append([])
                self.solutions[current_exercise - 1].append(ws.Cells(index_num, 3).Value)  # C{index_sum}

            previous_exercise = current_exercise
            index_num += + 1

        for i in range(len(exercise_row_begin)):
            self.exercise_ranges.append([exercise_row_begin[i], exercise_row_end[i]])

    def destroy(self):
        self.excel.Application.Quit()
        del self.excel

    def get_relevant_path(self):
        return os.sep.join(self.excel_file.parts[-3:])


class ExcelCorrector(Excel):
    def __init__(self, excel_file: Path):
        super().__init__(excel_file)

        self.subject_folder = excel_file.parent
        self.password = self.find_password()

        self.email = None
        # type: mail.Mail

        # Excel info
        self.subject_name = None
        self.deadline = None
        self.max_tries = None

        self.read_data()

    def read_data(self):
        wb = None
        try:
            wb = self.excel.Workbooks.Open(self.excel_file, True, False, None, self.password)
            ws = wb.Worksheets(1)
            # Check for valid subject name
            self.subject_name = ws.Range('B1').Value
            if not self.subject_name:
                self.log.error('Empty subject field in %s. Please specify a valid name.', self.get_relevant_path())
                raise ExcelFileException('Invalid subject field')

            # Check deadline, allow for same-day submissions
            deadline = [int(x or -1) for x in ws.Range('B9:D9').Value[0]]
            if -1 in deadline:
                self.log.error('Invalid date in deadline')
                raise ExcelFileException('Invalid deadline')
            self.deadline = datetime.date(deadline[2], deadline[1], deadline[0])
            if (self.deadline - datetime.date.today()).days < 0:
                self.log.info('Ignoring %s due to deadline (%s)', self.get_relevant_path(), self.deadline)
                raise ExcelFileException('Over deadline')

            self.max_tries = int(ws.Range('E12').Value or 0)
            if self.max_tries < 1:
                self.log.error('Invalid number of max tries')
                raise ExcelFileException('Invalid number of max tries')

            # Get amount of exercises
            self.set_exercise_rows(ws)

            email = ws.Range('B11:C11').Value[0]
            if None in email:
                self.log.error('Invalid email data')
                raise ExcelFileException('Invalid email login')
            email = ''.join(email).replace(' ', '')
            email_password = ws.Range('B12').Value
            # Login to mail account
            try:
                self.email = mail.Mail(email, email_password, self.subject_folder, self.subject_name)
            except mail.LoginException:
                self.log.error('Invalid login data for %s', self.get_relevant_path())
                self.log.error(traceback.format_exc())
                raise ExcelFileException('Invalid email login')
        except (pywintypes.com_error, TypeError, ValueError):
            self.log.error(traceback.format_exc())
            raise ExcelFileException('Failed to read relevant info from corrector')
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def generate_solutions(self, mat_num: int, dummies: []) -> Optional[list]:
        """
        Generates mat_num- and dummy-specific solutions and returns them in a two-dimensional list
        with the first dimension being the exercise and the second one containing
        the solution's name, value, and tolerance.

        :param mat_num: Student's matriculation number
        :param dummies: List of dummy values (a1-a100)
        :return:
        """
        wb = None
        try:
            wb = self.excel.Workbooks.Open(self.excel_file, True, False, None, self.password)
            ws = wb.Worksheets(1)

            ws.Range('B3').Value = mat_num
            ws.Range('A6:CV6').Value = dummies

            solutions = []

            for idx, exercise in enumerate(self.exercise_ranges):
                if len(solutions) <= idx:
                    solutions.append([])
                for cell_number in range(exercise[0], exercise[1] + 1):
                    solutions[idx].append({
                        'name': ws.Cells(cell_number, 2).Value,  # B{index}
                        'value': ws.Cells(cell_number, 3).Value,  # C{index}
                        'tolerance': ws.Cells(cell_number, 4).Value  # D{index}
                    })
            return solutions
        except (pywintypes.com_error, TypeError, ValueError):
            self.log.error(traceback.format_exc())
            raise ExcelFileException('Failed to generate solutions in corrector')
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def find_password(self) -> Optional[str]:
        # Look for psw file
        try:
            psw_file = open(self.subject_folder / 'psw', 'r')
            psw_enc = psw_file.read()
            psw_file.close()
            return simplecrypt.decrypt(config.PSW_PASSPHRASE, psw_enc)
        except IOError:  # psw file doesn't exist
            return ''
        except simplecrypt.DecryptionException:  # Decryption failed, ignore corrector file
            self.log.error('Failed to decrypt psw for %s', self.get_relevant_path())
            self.log.error(traceback.format_exc())
            return None

    @staticmethod
    def from_subject_folder(subject_folder: Path) -> Optional['ExcelCorrector']:
        extensions = ['.xlsx', '.xlsm', '.xls']

        # Ignore folders containing the ignore file
        ignore_file = subject_folder / 'PYCOR_IGNORE.txt'
        if ignore_file.exists():
            return

        for item in subject_folder.iterdir():
            if item.is_file() and item.suffix in extensions:
                return ExcelCorrector(item)
        return


class ExcelStudent(Excel):
    def __init__(self, excel_file: Path):
        super().__init__(excel_file)
        self.mat_num = None
        self.dummies = None

        self.student_folder = excel_file.parent
        self.student_email = self.student_folder.name

        self.read_data()

    def read_data(self):
        wb = None
        try:
            wb = self.excel.Workbooks.Open(self.excel_file)
            ws = wb.Worksheets(1)
            self.mat_num = int(ws.Range('B3').Value or -1)
            self.dummies = ws.Range('A6:CV6').Value[0]  # Get a1-a100 values

            if self.mat_num < 0:
                self.log.error('Invalid matriculation number specified')
                raise ExcelFileException('Invalid mat_num')

            self.set_exercise_rows(ws, solutions=True)
        except (pywintypes.com_error, TypeError, ValueError):
            self.log.error(traceback.format_exc())
            raise ExcelFileException('Failed to read relevant info')
        except AttributeError:
            self.log.error(traceback.format_exc())
            self.log.error('Looks like excel crashed. Let\'s quit.')
            raise
        finally:
            if wb:
                wb.Close(SaveChanges=False)
