import calendar
import time
import hashlib
import os.path

ENCODING = 'utf-8'
PID = str(calendar.timegm(time.gmtime()))

# OUTPUT_DIR = '/var/www/mystro.com/data/rcti_comparison/'
OUTPUT_DIR = './Output/'

OUTPUT_DIR_REFERRER = OUTPUT_DIR + 'referrer_rctis/'
OUTPUT_DIR_REFERRER_PID = OUTPUT_DIR_REFERRER + PID + '/'

OUTPUT_DIR_BROKER = OUTPUT_DIR + 'broker_rctis/'
OUTPUT_DIR_BROKER_PID = OUTPUT_DIR_BROKER + PID + '/'

OUTPUT_DIR_SUMMARY = OUTPUT_DIR + 'executive_summary/'
OUTPUT_DIR_SUMMARY_PID = OUTPUT_DIR_SUMMARY + PID + '/'


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

        if n1 or n2 == '':
            return False

        if type(n1) == str:
            n1val = float(n1[-1:])  # remove $
        if type(n2) == str:
            n2val = float(n2[-1:])  # remove $

        return abs(n1val - n2val) <= margin

    def serialize(self):
        return self.__dict__


def create_summary_dir():
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)

    if not os.path.exists(OUTPUT_DIR_SUMMARY):
        os.mkdir(OUTPUT_DIR_SUMMARY)

    if not os.path.exists(OUTPUT_DIR_SUMMARY_PID):
        os.mkdir(OUTPUT_DIR_SUMMARY_PID)


def create_detailed_dir():
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)

    if not os.path.exists(OUTPUT_DIR_REFERRER):
        os.mkdir(OUTPUT_DIR_REFERRER)

    if not os.path.exists(OUTPUT_DIR_REFERRER_PID):
        os.mkdir(OUTPUT_DIR_REFERRER_PID)

    if not os.path.exists(OUTPUT_DIR_BROKER):
        os.mkdir(OUTPUT_DIR_BROKER)

    if not os.path.exists(OUTPUT_DIR_BROKER_PID):
        os.mkdir(OUTPUT_DIR_BROKER_PID)


def new_error(file, msg, line='', first_a='', first_b='', second_a='', second_b='', third_a='',
              third_b='', fourth_a='', fourth_b='', fifth_a='', fifth_b=''):
    return {
        'file': file,
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


def write_errors(errors: list, worksheet, row, col, header_fmt):
    # Write summary header
    worksheet.write(row, col, 'File', header_fmt)
    worksheet.write(row, col + 1, 'Message', header_fmt)
    worksheet.write(row, col + 2, 'Line', header_fmt)
    worksheet.write(row, col + 3, 'First Value', header_fmt)
    worksheet.write(row, col + 4, 'First To Compare', header_fmt)
    worksheet.write(row, col + 5, 'Second Value', header_fmt)
    worksheet.write(row, col + 6, 'Second To Compare', header_fmt)
    worksheet.write(row, col + 7, 'Third Value', header_fmt)
    worksheet.write(row, col + 8, 'Third To Compare', header_fmt)
    worksheet.write(row, col + 9, 'Fourth Value', header_fmt)
    worksheet.write(row, col + 10, 'Fourth To Compare', header_fmt)
    worksheet.write(row, col + 11, 'Fifth Value', header_fmt)
    worksheet.write(row, col + 12, 'Fifth To Compare', header_fmt)
    row += 1

    # Write errors
    for error in errors:
        worksheet.write(row, col, error['file'])
        worksheet.write(row, col + 1, error['msg'])
        worksheet.write(row, col + 2, error['line'])
        worksheet.write(row, col + 3, error['first_a'])
        worksheet.write(row, col + 4, error['first_b'])
        worksheet.write(row, col + 5, error['second_a'])
        worksheet.write(row, col + 6, error['second_b'])
        worksheet.write(row, col + 7, error['third_a'])
        worksheet.write(row, col + 8, error['third_b'])
        worksheet.write(row, col + 9, error['fourth_a'])
        worksheet.write(row, col + 10, error['fourth_b'])
        worksheet.write(row, col + 11, error['fifth_a'])
        worksheet.write(row, col + 12, error['fifth_b'])
        row += 1

    return worksheet


def worksheet_write(worksheet, row, col, label, fmt_label, value, fmt_value):
    worksheet.write(row, col, label, fmt_label)
    worksheet.write(row, col + 1, value, fmt_value)