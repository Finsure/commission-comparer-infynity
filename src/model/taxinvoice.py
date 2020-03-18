import calendar
import time
import hashlib
import os.path

import xlsxwriter
from src import utils as u

ENCODING = 'utf-8'
PID = str(calendar.timegm(time.gmtime()))

# OUTPUT_DIR = '/var/www/mystro.com/data/rcti_comparison/'
OUTPUT_DIR = './Output/'
OUTPUT_DIR_PID = OUTPUT_DIR + PID + '/'
OUTPUT_DIR_REFERRER = OUTPUT_DIR_PID + 'referrer_rctis/'
OUTPUT_DIR_BROKER = OUTPUT_DIR_PID + 'broker_rctis/'
OUTPUT_DIR_BRANCH = OUTPUT_DIR_PID + 'branch_rctis/'
OUTPUT_DIR_SUMMARY = OUTPUT_DIR_PID + 'executive_summary/'


class TaxInvoice:

    def __init__(self, directory, filename):
        self.directory = directory
        self.filename = filename
        self._key = self.__generate_key()

    @property
    def full_path(self):
        self.__fix_path()
        return self.directory + self.filename

    @property
    def key(self):
        return self._key

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(self.filename.encode(ENCODING))
        return sha.hexdigest()

    def __fix_path(self):
        if self.directory[-1] != '/':
            self.directory += '/'

    def create_workbook(self, dir_):
        filename = self.filename
        if filename.endswith('.xls'):
            filename = filename[:-4]
        return xlsxwriter.Workbook(f"{dir_}DETAILED_{filename}.xlsx")

    def compare_numbers(self, n1, n2, margin):
        return u.compare_numbers(n1, n2, margin)


class InvoiceRow:

    def __init__(self):
        pass

    def compare_numbers(self, n1, n2, margin):
        return u.compare_numbers(n1, n2, margin)

    def serialize(self):
        return self.__dict__


def create_dirs():
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)

    if not os.path.exists(OUTPUT_DIR_PID):
        os.mkdir(OUTPUT_DIR_PID)

    if not os.path.exists(OUTPUT_DIR_REFERRER):
        os.mkdir(OUTPUT_DIR_REFERRER)

    if not os.path.exists(OUTPUT_DIR_BROKER):
        os.mkdir(OUTPUT_DIR_BROKER)

    if not os.path.exists(OUTPUT_DIR_BRANCH):
        os.mkdir(OUTPUT_DIR_BRANCH)

    if not os.path.exists(OUTPUT_DIR_SUMMARY):
        os.mkdir(OUTPUT_DIR_SUMMARY)


def new_error(file_a, file_b, msg, line='', value_a='', value_b='', tab=''):
    return {
        'file_a': file_a,
        'file_b': file_b,
        'tab': tab,
        'msg': msg,
        'line': line,
        'value_a': value_a,
        'value_b': value_b,
    }


def write_errors(errors: list, worksheet, row, col, header_fmt, filepath_a, filepath_b):
    # Write summary header
    worksheet.write(row, col, f'File Path A: {filepath_a}', header_fmt)
    worksheet.write(row, col + 1, f'File Path B: {filepath_b}', header_fmt)
    worksheet.write(row, col + 2, 'Message', header_fmt)
    worksheet.write(row, col + 3, 'Tab', header_fmt)
    worksheet.write(row, col + 4, 'Line', header_fmt)
    worksheet.write(row, col + 5, 'A Value', header_fmt)
    worksheet.write(row, col + 6, 'B Value', header_fmt)
    row += 1

    # Write errors
    for error in errors:
        worksheet.write(row, col, error['file_a'])
        worksheet.write(row, col + 1, error['file_b'])
        worksheet.write(row, col + 2, error['msg'])
        worksheet.write(row, col + 3, error['tab'])
        worksheet.write(row, col + 4, error['line'])
        worksheet.write(row, col + 5, error['value_a'])
        worksheet.write(row, col + 6, error['value_b'])
        row += 1

    return worksheet


def worksheet_write(worksheet, row, col, label, fmt_label, value, fmt_value):
    worksheet.write(row, col, label, fmt_label)
    worksheet.write(row, col + 1, value, fmt_value)


def get_header_format(workbook):
    return workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})


def get_title_format(workbook):
    return workbook.add_format({'font_size': 20, 'bold': True})


def get_error_format(workbook):
    return workbook.add_format({'font_color': 'red'})
