import calendar
import time
import hashlib
import os.path

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


class InvoiceRow:

    def __init__(self):
        pass

    def compare_numbers(self, n1, n2, margin):
        n1val = n1
        n2val = n2

        if str(n1).startswith('$'):
            n1val = float(n1[-1:])  # remove $
        if str(n2).startswith('$'):
            n2val = float(n2[-1:])  # remove $

        try:
            n1val = float(n1val)
            n2val = float(n2val)
        except ValueError:
            if n1val == '' or n2val == '':
                return n1val == n2val
            return False

        return abs(n1val - n2val) <= margin + 0.000001

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


def new_error(file_a, file_b, msg, line='', first_a='', first_b='', second_a='', second_b='', third_a='',
              third_b='', fourth_a='', fourth_b='', fifth_a='', fifth_b='', tab=''):
    return {
        'file_a': file_a,
        'file_b': file_b,
        'tab': tab,
        'msg': msg,
        'line': line,
        'first_a': first_a,
        'first_b': first_b,
        'second_a': second_a,
        'second_b': second_b,
        'third_a': third_a,
        'third_b': third_b,
        'fourth_a': fourth_a,
        'fourth_b': fourth_b,
        'fifth_a': fifth_a,
        'fifth_b': fifth_b
    }


def write_errors(errors: list, worksheet, row, col, header_fmt, filepath_a, filepath_b):
    # Write summary header
    worksheet.write(row, col, 'File Path A: ' + filepath_a, header_fmt)
    worksheet.write(row, col + 1, 'File Path B: ' + filepath_b, header_fmt)
    worksheet.write(row, col + 2, 'Message', header_fmt)
    worksheet.write(row, col + 3, 'Tab', header_fmt)
    worksheet.write(row, col + 4, 'Line', header_fmt)
    worksheet.write(row, col + 5, 'DEV A', header_fmt)
    worksheet.write(row, col + 6, 'Finsure A', header_fmt)
    worksheet.write(row, col + 7, 'DEV B', header_fmt)
    worksheet.write(row, col + 8, 'Finsure B', header_fmt)
    worksheet.write(row, col + 9, 'DEV C', header_fmt)
    worksheet.write(row, col + 10, 'Finsure C', header_fmt)
    worksheet.write(row, col + 11, 'DEV D', header_fmt)
    worksheet.write(row, col + 12, 'Finsure D', header_fmt)
    worksheet.write(row, col + 13, 'DEV E', header_fmt)
    worksheet.write(row, col + 14, 'Finsure E', header_fmt)
    row += 1

    # Write errors
    for error in errors:
        worksheet.write(row, col, error['file_a'])
        worksheet.write(row, col + 1, error['file_b'])
        worksheet.write(row, col + 2, error['msg'])
        worksheet.write(row, col + 3, error['tab'])
        worksheet.write(row, col + 4, error['line'])
        worksheet.write(row, col + 5, error['first_a'])
        worksheet.write(row, col + 6, error['first_b'])
        worksheet.write(row, col + 7, error['second_a'])
        worksheet.write(row, col + 8, error['second_b'])
        worksheet.write(row, col + 9, error['third_a'])
        worksheet.write(row, col + 10, error['third_b'])
        worksheet.write(row, col + 11, error['fourth_a'])
        worksheet.write(row, col + 12, error['fourth_b'])
        worksheet.write(row, col + 13, error['fifth_a'])
        worksheet.write(row, col + 14, error['fifth_b'])
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
