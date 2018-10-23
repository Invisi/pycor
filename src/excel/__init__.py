import datetime
import logging
import os
import traceback
from typing import Optional

import pywintypes
import simplecrypt
import win32com.client

import config
import mail


class CorrectorException(Exception):
    pass


class Excel:
    def __init__(self, excel_file):
        self.log = logging.getLogger('PyCor').getChild('Excel')

        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel_file = excel_file
        self.subject_folder = os.path.dirname(excel_file)
        self.log.info('Opening %s', self.get_relevant_path())

    def destroy(self):
        self.excel.DisplayAlerts = False
        self.excel.Application.Quit()
        del self.excel

    def get_relevant_path(self):
        return self.excel_file.replace(config.BASE_FOLDER + os.sep, '')


class ExcelCorrector(Excel):
    def __init__(self, excel_file):
        super().__init__(excel_file)

        self.email = None
        # type: mail.Mail

        # Excel info
        self.subject_name = None
        self.deadline = None
        self.max_tries = None
        self.exercise_count = None

        self.read_corrector()

    def read_corrector(self):
        wb = self.excel.Workbooks.Open(self.excel_file, True, False, None, self.find_password())
        ws = wb.Worksheets(1)

        try:
            # Check for valid subject name
            self.subject_name = ws.Range('B1').Value
            if not self.subject_name:
                self.log.error('Empty subject field in %s. Please specify a valid name.', self.excel_file)
                raise CorrectorException('Invalid subject field')

            # Check deadline, allow for same-day submissions
            deadline = [int(x or -1) for x in ws.Range('B9:D9').Value[0]]
            if -1 in deadline:
                self.log.error('Invalid date in deadline')
                raise CorrectorException('Invalid deadline')
            self.deadline = datetime.date(deadline[2], deadline[1], deadline[0])
            if (self.deadline - datetime.date.today()).days < 0:
                self.log.info('Ignoring %s due to deadline (%s)', self.excel_file, self.deadline)
                raise CorrectorException('Over deadline')

            self.max_tries = ws.Range('E12').Value
            if self.max_tries < 1:
                self.log.error('Invalid number of max tries')
                raise CorrectorException('Invalid number of max tries')

            # Get amount of exercises
            index_num = 16
            cell_previous = 0
            self.exercise_count = 0
            while True:
                cell_active = ws.Cells(index_num, 1).Value  # Cell A{index_num}
                if not cell_active or cell_active < cell_previous:
                    break

                self.exercise_count = cell_previous
                cell_previous = cell_active
                index_num += + 1

            email = ws.Range('B11:C11').Value[0]
            if None in email:
                self.log.error('Invalid email data')
                raise CorrectorException('Invalid email login')
            email = ''.join(email).replace(' ', '')
            email_password = ws.Range('B12').Value
            # Login to mail account
            try:
                self.email = mail.Mail(email, email_password)
            except mail.LoginException:
                self.log.error('Invalid login data for %s', self.excel_file)
                self.log.error(traceback.format_exc())
                raise CorrectorException('Invalid email login')
        except (pywintypes.com_error, TypeError, ValueError):
            self.log.error(traceback.format_exc())
            raise CorrectorException('Failed to read relevant info from corrector')
        finally:
            wb.Close(SaveChanges=False)

    def find_password(self) -> Optional[str]:
        # Look for psw file
        try:
            psw_file = open(os.path.join(self.subject_folder, 'psw'), 'r')
            psw_enc = psw_file.read()
            psw_file.close()
            return simplecrypt.decrypt('password', psw_enc)
        except IOError:  # psw file doesn't exist
            return ''
        except simplecrypt.DecryptionException:  # Decryption failed, ignore corrector file
            self.log.error('Failed to decrypt psw for %s', self.subject_folder)
            self.log.error(traceback.format_exc())
            return None

    @staticmethod
    def from_subject_folder(subject_folder) -> Optional[Excel]:
        items = os.listdir(subject_folder)
        extensions = ['xlsx', 'xlsm', 'xls']

        # TODO: Mark folder as ignored via PYCOR_ERROR.txt or PYCOR_IGNORE.txt

        for item in items:
            corr_path = os.path.join(subject_folder, item)
            if os.path.isfile(corr_path) and any([item.endswith(x) for x in extensions]):
                return ExcelCorrector(corr_path)
        return None
