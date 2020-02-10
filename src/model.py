import copy
import hashlib
import calendar
import time
import os
import os.path
import numpy

import xlsxwriter
import pandas
from bs4 import BeautifulSoup

from src.utils import merge_lists

ENCODING = 'utf-8'
PID = str(calendar.timegm(time.gmtime()))
# OUTPUT_DIR = '/var/www/mystro.com/data/rcti_comparison/'
OUTPUT_DIR = './Output/'
OUTPUT_DIR_REFERRER = OUTPUT_DIR + 'referrer_rctis/'
OUTPUT_DIR_REFERRER_PID = OUTPUT_DIR_REFERRER + PID + '/'
OUTPUT_DIR_SUMMARY = OUTPUT_DIR + 'executive_summary/'
OUTPUT_DIR_SUMMARY_PID = OUTPUT_DIR_SUMMARY + PID + '/'


class File:

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


class BrokerTaxInvoice(File):

    def __init__(self, directory, filename):
        File.__init__(self, directory, filename)
        self.parse()

    def parse(self):
        dataframe = pandas.read_excel(self.full_path)
        dataframe = dataframe.replace(numpy.nan, '', regex=True)

        dataframe_broker_info = dataframe.iloc[2:5, 0:2]

        account_info = dataframe.iloc[len(dataframe.index) - 1][1]
        account_info_parts = account_info.split(':')

        bsb = account_info_parts[1].strip().split('/')[0][1:]

        account = account_info_parts[1].strip().split('/')[1]
        if account[-1] == ')':
            account = account[:-1]

        self.from_ = dataframe_broker_info.iloc[0][1]
        self.to = dataframe_broker_info.iloc[1][1]
        self.abn = dataframe_broker_info.iloc[2][1]
        self.bsb = bsb
        self.account = account

        self.rows = self.parse_rows(dataframe)

    def parse_rows(self, dataframe):
        dataframe_rows = dataframe.iloc[8:len(dataframe.index) - 1]
        dataframe_rows = dataframe_rows.rename(columns=dataframe_rows.iloc[0]).drop(dataframe_rows.index[0])
        dataframe_rows = dataframe_rows.dropna(how='all')  # remove rows that don't have any value

        rows = []
        for index, row in dataframe_rows.iterrows():
            invoice_row = BrokerInvoiceRow(
                row['Commission Type'], row['Client'], row['Commission Ref ID'], row['Bank'],
                row['Loan Balance'], row['Amount Paid'], row['GST Paid'],
                row['Total Amount Paid'], row['Comments'], index)
            rows.append(invoice_row)
        return rows

    def compare_to(self, invoice, margin=0.000001):
        result = result_invoice_broker()
        result['filename'] = self.filename
        result['file'] = self.full_path

        has_pair = invoice is not None
        result['has_pair'] = has_pair

        result['from_a'] = self.from_
        result['to_a'] = self.to
        result['abn_a'] = self.abn
        result['bsb_a'] = self.bsb
        result['account_a'] = self.account

        if has_pair:
            result['from_b'] = invoice.from_
            result['to_b'] = invoice.to
            result['abn_b'] = invoice.abn
            result['bsb_b'] = invoice.bsb
            result['account_b'] = invoice.account
            result['equal_from'] = self.from_ == invoice.from_
            result['equal_to'] = self.to == invoice.to
            result['equal_abn'] = self.abn == invoice.abn
            result['equal_bsb'] = self.bsb == invoice.bsb
            result['equal_account'] = self.account == invoice.account
            result['equal_amount_rows'] = len(self.rows) == len(invoice.rows)

        result['overall'] = (result['equal_from'] and result['equal_to'] and result['equal_abn']
                             and result['equal_bsb'] and result['equal_account']
                             and result['equal_amount_rows'])

        # TODO: Finish this method #2 after finishing #1


class BrokerInvoiceRow(InvoiceRow):

    def __init__(self, commission_type, client, reference_id, bank, loan_balance, amount_paid,
                 gst_paid, total_amount_paid, comments, row_number):
        InvoiceRow.__init__(self)
        self.commission_type = str(commission_type)
        self.client = str(client)
        self.reference_id = str(reference_id)
        self.bank = str(bank)
        self.loan_balance = str(loan_balance)
        self.amount_paid = str(amount_paid)
        self.gst_paid = str(gst_paid)
        self.total_amount_paid = str(total_amount_paid)
        self.comments = str(comments)
        self.row_number = str(row_number)

        self._key = self.__generate_key()
        self._key_full = self.__generate_key_full()

    @property
    def key(self):
        return self._key

    @property
    def key_full(self):
        return self._key_full

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.reference_id.encode(ENCODING))
        sha.update(self.bank.encode(ENCODING))
        return sha.hexdigest()

    def __generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.reference_id.encode(ENCODING))
        sha.update(self.bank.encode(ENCODING))
        sha.update(self.loan_balance.encode(ENCODING))
        sha.update(self.amount_paid.encode(ENCODING))
        sha.update(self.gst_paid.encode(ENCODING))
        sha.update(self.total_amount_paid.encode(ENCODING))
        sha.update(self.comments.encode(ENCODING))
        return sha.hexdigest()

    def compare_to(self, row, margin=0.0000001, reverse=True):
        result = result_row_broker()
        result['row_number'] = self.row_number

        has_pair = row is not None
        result['has_pair'] = has_pair

        equal_loan_balance = False
        equal_amount_paid = False
        equal_gst_paid = False
        equal_total_amount_paid = False
        equal_comments = False

        if has_pair:
            equal_loan_balance = self.compare_numbers(self.loan_balance, row.loan_balance, margin)
            equal_amount_paid = self.compare_numbers(self.amount_paid, row.amount_paid, margin)
            equal_gst_paid = self.compare_numbers(self.gst_paid, row.gst_paid, margin)
            equal_total_amount_paid = self.compare_numbers(self.total_amount_paid, row.total_amount_paid, margin)
            equal_comments = self.comments == row.comments

        overall = equal_loan_balance and equal_amount_paid and equal_gst_paid and equal_total_amount_paid and equal_comments

        a = 'a'
        b = 'b'
        if reverse:
            a = 'b'
            b = 'a'

        result['overall'] = overall
        result['loan_balance'] = equal_loan_balance
        result['amount_paid'] = equal_amount_paid
        result['gst_paid'] = equal_gst_paid
        result['total_amount_paid'] = equal_total_amount_paid
        result['comments'] = equal_comments

        result['commission_type_' + a] = self.commission_type
        result['client_' + a] = self.client
        result['reference_id_' + a] = self.reference_id
        result['bank_' + a] = self.bank
        result['loan_balance_' + a] = self.loan_balance
        result['amount_paid_' + a] = self.amount_paid
        result['gst_paid_' + a] = self.gst_paid
        result['total_amount_paid_' + a] = self.total_amount_paid
        result['comments_' + a] = self.comments

        if has_pair:
            result['commission_type_' + b] = row.commission_type
            result['client_' + b] = row.client
            result['reference_id_' + b] = row.reference_id
            result['bank_' + b] = row.bank
            result['loan_balance_' + b] = row.loan_balance
            result['amount_paid_' + b] = row.amount_paid
            result['gst_paid_' + b] = row.gst_paid
            result['total_amount_paid_' + b] = row.total_amount_paid
            result['comments_' + b] = row.comments

        return result


class ReferrerTaxInvoice:

    def __init__(self, directory, filename):
        self.directory = directory
        self.filename = filename
        self.filetext = self.get_file_text()
        self.parse()

        self._key = None

    def get_full_path(self):
        self.__fix_path()
        return self.directory + self.filename

    def get_file_text(self):
        file = open(self.get_full_path(), 'r')
        return file.read()

    def parse(self):
        soup = BeautifulSoup(self.filetext, 'html.parser')

        self._from = self.parse_from(soup)
        self.from_abn = self.parse_from_abn(soup)
        self.to = self.parse_to(soup)
        self.to_abn = self.parse_to_abn(soup)
        self.bsb = self.parse_bsb(soup)
        self.account = self.parse_account(soup)
        self.final_total = self.parse_final_total(soup)
        self.rows = self.parse_rows(soup)

    def parse_from(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        _from = parts_info[1][:-4]
        _from = _from.strip()
        return _from

    def parse_from_abn(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        abn = parts_info[2][:-3]
        abn = abn.strip()
        return abn

    def parse_to(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        to = parts_info[3][:-4]
        to = to.strip()
        return to

    def parse_to_abn(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        abn = parts_info[4][:-5]
        abn = abn.strip()
        return abn

    def parse_bsb(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        bsb = parts_account[1].split(' - ')[0].strip()
        return bsb

    def parse_account(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        account = parts_account[2].split('/')[0].strip()
        return account

    def parse_final_total(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        final_total = parts_account[3].strip()
        return final_total

    def parse_rows(self, soup: BeautifulSoup):
        header = soup.find('tr')  # Find header
        header.extract()  # Remove header
        table_rows = soup.find_all('tr')
        row_number = 0
        rows = {}
        for tr in table_rows:
            row_number += 1
            tds = tr.find_all('td')
            try:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                         tds[3].text, tds[4].text, tds[5].text, row_number)
                rows[row.key_full()] = row
            except IndexError:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, '',
                                         tds[2].text, tds[3].text, tds[4].text, row_number)
                rows[row.key_full()] = row
        return rows

    def key(self):
        if self._key is None:
            self._key = self.__generate_key()
        return self._key

    def serialize(self):
        # we do this do we dont serialize the filetext because it is too big.
        text = self.filetext
        self.filetext = None
        serialized_obj = copy.copy(self.__dict__)
        self.filetext = text
        return serialized_obj

    # Man I hope I never need to maintain this!!!
    def compare_to(self, invoice, margin=0.0000001):
        result = result_invoice_referrer()
        result['filename'] = self.filename
        result['file'] = self.get_full_path()

        has_pair = invoice is not None
        #  If we reached here it means the file has a pair
        result['has_pair'] = has_pair
        if has_pair:
            result['equal_from'] = self._from == invoice._from
            result['equal_from_abn'] = self.from_abn == invoice.from_abn
            result['equal_to'] = self.to == invoice.to
            result['equal_to_abn'] = self.to_abn == invoice.to_abn
            result['equal_bsb'] = self.bsb == invoice.bsb
            result['equal_account'] = self.account == invoice.account
            result['equal_final_total'] = self.final_total == invoice.final_total
            result['equal_amount_rows'] = len(self.rows) == len(invoice.rows)

        # Results values for display purposes
        result['from_value_1'] = self._from
        result['from_abn_value_1'] = self.from_abn
        result['to_value_1'] = self.to
        result['to_abn_value_1'] = self.to_abn
        result['bsb_value_1'] = self.bsb
        result['account_value_1'] = self.account
        result['final_total_value_1'] = self.final_total
        if has_pair:
            result['from_value_2'] = invoice._from
            result['from_abn_value_2'] = invoice.from_abn
            result['to_value_2'] = invoice.to
            result['to_abn_value_2'] = invoice.to_abn
            result['bsb_value_2'] = invoice.bsb
            result['account_value_2'] = invoice.account
            result['final_total_value_2'] = invoice.final_total

        result['overall'] = (result['equal_from'] and result['equal_from_abn']
                             and result['equal_to'] and result['equal_to_abn']
                             and result['equal_bsb'] and result['equal_account']
                             and result['equal_final_total'] and result['equal_amount_rows'])

        if not has_pair:
            result_rows = {}
            for key in self.rows.keys():
                row_local = self.rows[key]
                result_rows[key] = row_local.compare_to(None)
            result['results_rows'] = result_rows
            return result

        # ensure both have been parsed
        if len(self.rows) == 0:
            self.parse()
        if len(invoice.rows) == 0:
            invoice.parse()

        keys_all = merge_lists(self.rows.keys(), invoice.rows.keys())

        result_rows = {}

        for key in keys_all:
            row_local = self.rows.get(key, None)
            row_invoice = invoice.rows.get(key, None)
            use_key = key

            # If we couldnt find the row by the ReferrerInvoiceRow.full_key() it means they are different
            # so we try to locate them by the ReferrerInvoiceRow.key()
            if row_local is None:
                for k in self.rows.keys():
                    if self.rows.get(k).key() == row_invoice.key():
                        row_local = self.rows[k]
                        keys_all.remove(row_local.key_full())
                        use_key = k
            elif row_invoice is None:
                for k in invoice.rows.keys():
                    if invoice.rows.get(k).key() == row_local.key():
                        row_invoice = invoice.rows[k]
                        keys_all.remove(row_invoice.key_full())
                        use_key = k

            if row_local is not None:
                result_rows[use_key] = row_local.compare_to(row_invoice, margin, False)
            else:
                result_rows[use_key] = row_invoice.compare_to(row_local, margin, True)

        result['results_rows'] = result_rows

        for key in result['results_rows'].keys():
            result['overall'] = result['overall'] and result['results_rows'][key]['overall']

        return result

    def _get_parts_info(self, soup: BeautifulSoup):
        body = soup.find('body')
        extracted_info = body.find('p').text
        info = ' '.join(extracted_info.split())
        parts_info = info.split(':')
        return parts_info

    def _get_parts_account(self, soup: BeautifulSoup):
        body = soup.find('body')
        extracted_account = body.find('p').find_next('p').text
        account = ' '.join(extracted_account.split())
        parts_account = account.split(':')
        return parts_account

    def __fix_path(self):
        if self.directory[-1] != '/':
            self.directory += '/'

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[:-5]  # Remove process ID and date stamp

        for index, part in enumerate(filename_parts):
            if part == "Referrer":
                del filename_parts[index - 1]  # Remove year-month stamp

        filename_forkey = '_'.join(filename_parts)
        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()


class ReferrerInvoiceRow:

    def __init__(self, commission_type, client, referrer, amount_paid, gst_paid, total, row_number):
        self.commission_type = commission_type
        self.client = client
        self.referrer = referrer
        self.amount_paid = amount_paid
        self.gst_paid = gst_paid
        self.total = total
        self.row_number = row_number

        self._key = None
        self._key_full = None

    def key(self):
        if self._key is None:
            self._key = self.__generate_key()
        return self._key

    def key_full(self):
        if self._key_full is None:
            self._key_full = self.__generate_key_full()
        return self._key_full

    def serialize(self):
        return self.__dict__

    def compare_to(self, row, margin=0.0000001, reverse=True):
        result = result_row_referrer()
        result['row_number'] = self.row_number

        has_pair = row is not None
        result['has_pair'] = has_pair

        equal_amount_paid = False
        equal_gst_paid = False
        equal_total = False
        if has_pair:
            equal_amount_paid = self.amount_paid == row.amount_paid
            equal_gst_paid = self.gst_paid == row.gst_paid
            equal_total = self.total == row.total

            # Recompare monetary values using the
            if not equal_amount_paid:
                equal_amount_paid = self.compare_numbers(self.amount_paid, row.amount_paid, margin)
            if not equal_gst_paid:
                equal_gst_paid = self.compare_numbers(self.gst_paid, row.gst_paid, margin)
            if not equal_total:
                equal_total = self.compare_numbers(self.total, row.total, margin)

        overall = equal_amount_paid and equal_gst_paid and equal_total
        # and equal_commission_type and equal_client and equal_referrer)

        first = '1'
        second = '2'
        if reverse:
            first = '2'
            second = '1'

        result['overall'] = overall
        result['amount_paid'] = equal_amount_paid
        result['gst_paid'] = equal_gst_paid
        result['total'] = equal_total
        result['commission_type_value_' + first] = self.commission_type
        result['client_value_' + first] = self.client
        result['referrer_value_' + first] = self.referrer
        result['amount_paid_value_' + first] = self.amount_paid
        result['gst_paid_value_' + first] = self.gst_paid
        result['total_value_' + first] = self.total

        if has_pair:
            result['commission_type_value_' + second] = row.commission_type
            result['client_value_' + second] = row.client
            result['referrer_value_' + second] = row.referrer
            result['amount_paid_value_' + second] = row.amount_paid
            result['gst_paid_value_' + second] = row.gst_paid
            result['total_value_' + second] = row.total

        return result

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

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.referrer.encode(ENCODING))
        return sha.hexdigest()

    def __generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.referrer.encode(ENCODING))
        sha.update(self.amount_paid.encode(ENCODING))
        sha.update(self.gst_paid.encode(ENCODING))
        sha.update(self.total.encode(ENCODING))
        return sha.hexdigest()


def result_invoice_referrer():
    return {
        'filename': '',
        'file': '',
        'has_pair': False,
        'equal_from': False,
        'equal_from_abn': False,
        'equal_to': False,
        'equal_to_abn': False,
        'equal_bsb': False,
        'equal_account': False,
        'equal_final_total': False,
        'equal_amount_rows': False,
        'overall': False,
        'results_rows': {},
        'from_value_1': '',
        'from_value_2': '',
        'from_abn_value_1': '',
        'from_abn_value_2': '',
        'to_value_1': '',
        'to_value_2': '',
        'to_abn_value_1': '',
        'to_abn_value_2': '',
        'bsb_value_1': '',
        'bsb_value_2': '',
        'account_value_1': '',
        'account_value_2': '',
        'final_total_value_1': '',
        'final_total_value_2': ''
    }


def result_invoice_broker():
    return {
        'filename': '',
        'file': '',
        'has_pair': False,
        'equal_from': False,
        'equal_to': False,
        'equal_abn': False,
        'equal_bsb': False,
        'equal_account': False,
        'equal_amount_rows': False,
        'overall': False,
        'invoice_rows': {},
        'from_a': '',
        'from_b': '',
        'to_a': '',
        'to_b': '',
        'abn_a': '',
        'abn_b': '',
        'bsb_a': '',
        'bsb_b': '',
        'account_a': '',
        'account_b': '',
    }


def result_row_referrer():
    return {
        'overall': False,
        'has_pair': False,
        'commission_type': False,
        'client': False,
        'referrer': False,
        'amount_paid': False,
        'gst_paid': False,
        'total': False,
        'row_number': 0,
        'commission_type_value_1': '',
        'commission_type_value_2': '',
        'client_value_1': '',
        'client_value_2': '',
        'referrer_value_1': '',
        'referrer_value_2': '',
        'amount_paid_value_1': '',
        'amount_paid_value_2': '',
        'gst_paid_value_1': '',
        'gst_paid_value_2': '',
        'total_value_1': '',
        'total_value_2': ''
    }


def result_row_broker():
    return {
        'overall': False,
        'has_pair': False,
        'commission_type': False,
        'client': False,
        'reference_id': False,
        'amount_paid': False,
        'gst_paid': False,
        'total_amount_paid': False,
        'comments': False,
        'row_number': 0,
        'commission_type_a': '',
        'commission_type_b': '',
        'client_a': '',
        'client_b': '',
        'reference_id_a': '',
        'reference_id_b': '',
        'amount_paid_a': '',
        'amount_paid_b': '',
        'gst_paid_a': '',
        'gst_paid_b': '',
        'total_amount_paid_a': '',
        'total_amount_paid_b': '',
        'comments_a': '',
        'comments_b': ''
    }


def new_error(file, msg, line='', first_value_1='', first_value_2='', second_value_1='',
              second_value_2='', third_value_1='', third_value_2=''):
    return {
        'file': file,
        'msg': msg,
        'line': line,
        'first_value_1': first_value_1,
        'first_value_2': first_value_2,
        'second_value_1': second_value_1,
        'second_value_2': second_value_2,
        'third_value_1': third_value_1,
        'third_value_2': third_value_2
    }


# This function is ugly as shit. We must figure out a better design to simplify things.
def create_summary(results: list):

    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)

    if not os.path.exists(OUTPUT_DIR_SUMMARY):
        os.mkdir(OUTPUT_DIR_SUMMARY)

    if not os.path.exists(OUTPUT_DIR_SUMMARY_PID):
        os.mkdir(OUTPUT_DIR_SUMMARY_PID)

    workbook = xlsxwriter.Workbook(OUTPUT_DIR_SUMMARY_PID + 'referrer_rcti_summary.xlsx')
    worksheet = workbook.add_worksheet('Summary')

    row = 0
    col = 0

    fmt_title = workbook.add_format({'font_size': 20, 'bold': True})
    fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white',
                                            'bg_color': 'black'})

    worksheet.merge_range('A1:I1', 'Commission Referrer RCTI Summary', fmt_title)
    row += 2

    list_errors = []
    for result in results:
        if result['overall'] is False:  # it means there is an issue
            file = result['file']

            # Error when the file doesnt have a pair
            if not result['has_pair']:
                msg = 'No corresponding commission file found'
                error = new_error(file, msg)
                list_errors.append(error)
                continue

            # From does not match
            if not result['equal_from']:
                msg = 'From name does not match'
                error = new_error(file, msg, '', result['from_value_1'], result['from_value_2'])
                list_errors.append(error)

            # From ABN does not match
            if not result['equal_from_abn']:
                msg = 'From ABN does not match'
                error = new_error(file, msg, '', result['from_abn_value_1'], result['from_abn_value_2'])
                list_errors.append(error)

            # To does not match
            if not result['equal_to']:
                msg = 'To name does not match'
                error = new_error(file, msg, '', result['to_value_1'], result['to_value_2'])
                list_errors.append(error)

            # To ABN does not match
            if not result['equal_to_abn']:
                msg = 'To ABN does not match'
                error = new_error(file, msg, '', result['to_abn_value_1'], result['to_abn_value_2'])
                list_errors.append(error)

            # BSB does not match
            if not result['equal_bsb']:
                msg = 'BSB does not match'
                error = new_error(file, msg, '', result['bsb_value_1'], result['bsb_value_2'])
                list_errors.append(error)

            # Account does not match
            if not result['equal_account']:
                msg = 'Account number does not match'
                error = new_error(file, msg, '', result['account_value_1'], result['account_value_2'])
                list_errors.append(error)

            # Total does not match
            if not result['equal_final_total']:
                msg = 'Total does not match'
                error = new_error(file, msg, '', result['final_total_value_1'], result['final_total_value_2'])
                list_errors.append(error)

            for key in result['results_rows'].keys():
                result_row = result['results_rows'][key]
                if result_row['overall'] is False:  # it means there is an issue

                    if not result_row['has_pair']:
                        msg = 'No corresponding row in comission file'
                        error = new_error(file, msg, '')  # TODO: Include row number
                        list_errors.append(error)
                        continue

                    values_list = safelist([])
                    if not result_row['amount_paid']:
                        values_list.append(result_row['amount_paid_value_1'])
                        values_list.append(result_row['amount_paid_value_2'])
                    if not result_row['gst_paid']:
                        values_list.append(result_row['gst_paid_value_1'])
                        values_list.append(result_row['gst_paid_value_2'])
                    if not result_row['total']:
                        values_list.append(result_row['total_value_1'])
                        values_list.append(result_row['total_value_2'])

                    msg = 'Values not match'
                    error = new_error(file, msg, result_row['row_number'], values_list.get(0, ''),
                                      values_list.get(1, ''), values_list.get(2, ''),
                                      values_list.get(3, ''), values_list.get(4, ''),
                                      values_list.get(5, ''))
                    list_errors.append(error)

    # Write summary header
    worksheet.write(row, col, 'File', fmt_table_header)
    worksheet.write(row, col + 1, 'Message', fmt_table_header)
    worksheet.write(row, col + 2, 'Line', fmt_table_header)
    worksheet.write(row, col + 3, 'First Value', fmt_table_header)
    worksheet.write(row, col + 4, 'First To Compare', fmt_table_header)
    worksheet.write(row, col + 5, 'Second Value', fmt_table_header)
    worksheet.write(row, col + 6, 'Second To Compare', fmt_table_header)
    worksheet.write(row, col + 7, 'Third Value', fmt_table_header)
    worksheet.write(row, col + 8, 'Third To Compare', fmt_table_header)
    row += 1

    # Write errors
    for error in list_errors:
        worksheet.write(row, col, error['file'])
        worksheet.write(row, col + 1, error['msg'])
        worksheet.write(row, col + 2, error['line'])
        worksheet.write(row, col + 3, error['first_value_1'])
        worksheet.write(row, col + 4, error['first_value_2'])
        worksheet.write(row, col + 5, error['second_value_1'])
        worksheet.write(row, col + 6, error['second_value_2'])
        worksheet.write(row, col + 7, error['third_value_1'])
        worksheet.write(row, col + 8, error['third_value_2'])
        row += 1

    workbook.close()


def create_all_datailed_report(results: list):
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)

    if not os.path.exists(OUTPUT_DIR_REFERRER):
        os.mkdir(OUTPUT_DIR_REFERRER)

    if not os.path.exists(OUTPUT_DIR_REFERRER_PID):
        os.mkdir(OUTPUT_DIR_REFERRER_PID)

    for result in results:
        create_detailed_report(result)


# I promise there was no other way! :(
def create_detailed_report(result: dict):
    # If there is no error we dont need to generate this report.
    if result['overall']:
        return

    workbook = xlsxwriter.Workbook(OUTPUT_DIR_REFERRER_PID + 'DETAILED_' + result['filename'] + '.xlsx')
    worksheet = workbook.add_worksheet('Detailed')

    fmt_error = workbook.add_format({'font_color': 'red'})
    fmt_bold = workbook.add_format({'bold': True})
    fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white',
                                            'bg_color': 'black'})

    row = 0
    col = 0
    comparison_col = 8

    worksheet.merge_range('A1:N1', result['filename'])
    row += 2

    format_ = fmt_error if not result['equal_from'] else None
    worksheet.write(row, col, 'From', fmt_bold)
    worksheet.write(row, col + 1, result['from_value_1'], format_)
    worksheet.write(row, comparison_col, 'From', fmt_bold)
    worksheet.write(row, comparison_col + 1, result['from_value_2'], format_)

    row += 1

    format_ = fmt_error if not result['equal_from_abn'] else None
    worksheet.write(row, col, 'From ABN', fmt_bold)
    worksheet.write(row, col + 1, result['from_abn_value_1'], format_)
    worksheet.write(row, comparison_col, 'From ABN', fmt_bold)
    worksheet.write(row, comparison_col + 1, result['from_abn_value_2'], format_)

    row += 1

    format_ = fmt_error if not result['equal_to'] else None
    worksheet.write(row, col, 'To', fmt_bold)
    worksheet.write(row, col + 1, result['to_value_1'], format_)
    worksheet.write(row, comparison_col, 'To', fmt_bold)
    worksheet.write(row, comparison_col + 1, result['to_value_2'], format_)

    row += 1

    format_ = fmt_error if not result['equal_to_abn'] else None
    worksheet.write(row, col, 'To ABN', fmt_bold)
    worksheet.write(row, col + 1, result['to_abn_value_1'], format_)
    worksheet.write(row, comparison_col, 'To ABN', fmt_bold)
    worksheet.write(row, comparison_col + 1, result['to_abn_value_2'], format_)

    row += 2

    if result['has_pair']:

        worksheet.write(row, col, 'Commission Type', fmt_table_header)
        worksheet.write(row, col + 1, 'Client', fmt_table_header)
        worksheet.write(row, col + 2, 'Referrer Name', fmt_table_header)
        worksheet.write(row, col + 3, 'Amount Paid', fmt_table_header)
        worksheet.write(row, col + 4, 'GST Paid', fmt_table_header)
        worksheet.write(row, col + 5, 'Total Amount Paid', fmt_table_header)

        worksheet.write(row, comparison_col, 'Commission Type', fmt_table_header)
        worksheet.write(row, comparison_col + 1, 'Client', fmt_table_header)
        worksheet.write(row, comparison_col + 2, 'Referrer Name', fmt_table_header)
        worksheet.write(row, comparison_col + 3, 'Amount Paid', fmt_table_header)
        worksheet.write(row, comparison_col + 4, 'GST Paid', fmt_table_header)
        worksheet.write(row, comparison_col + 5, 'Total Amount Paid', fmt_table_header)

        for key in result['results_rows'].keys():
            row += 1
            result_row = result['results_rows'][key]

            format_ = fmt_error if not result_row['has_pair'] else None
            worksheet.write(row, col, result_row['commission_type_value_1'], format_)
            worksheet.write(row, col + 1, result_row['client_value_1'], format_)
            worksheet.write(row, col + 2, result_row['referrer_value_1'], format_)
            worksheet.write(row, comparison_col, result_row['commission_type_value_2'], format_)
            worksheet.write(row, comparison_col + 1, result_row['client_value_2'], format_)
            worksheet.write(row, comparison_col + 2, result_row['referrer_value_2'], format_)

            format_ = fmt_error if not result_row['amount_paid'] else None
            worksheet.write(row, col + 3, result_row['amount_paid_value_1'], format_)
            worksheet.write(row, comparison_col + 3, result_row['amount_paid_value_2'], format_)

            format_ = fmt_error if not result_row['gst_paid'] else None
            worksheet.write(row, col + 4, result_row['gst_paid_value_1'], format_)
            worksheet.write(row, comparison_col + 4, result_row['gst_paid_value_2'], format_)

            format_ = fmt_error if not result_row['total'] else None
            worksheet.write(row, col + 5, result_row['total_value_1'], format_)
            worksheet.write(row, comparison_col + 5, result_row['total_value_2'], format_)

    else:
        worksheet.write(row, col, 'No match to compare to', fmt_error)

    row += 2

    format_ = fmt_error if not result['equal_bsb'] else None
    worksheet.write(row, col, 'BSB', fmt_bold)
    worksheet.write(row, col + 1, result['bsb_value_1'], format_)
    worksheet.write(row, comparison_col, 'BSB', fmt_bold)
    worksheet.write(row, comparison_col + 1, result['bsb_value_2'], format_)

    format_ = fmt_error if not result['equal_account'] else None
    worksheet.write(row, col + 2, 'Account', fmt_bold)
    worksheet.write(row, col + 3, result['account_value_1'], format_)
    worksheet.write(row, comparison_col + 2, 'Account', fmt_bold)
    worksheet.write(row, comparison_col + 3, result['account_value_2'], format_)

    format_ = fmt_error if not result['equal_final_total'] else None
    worksheet.write(row, col + 4, 'Amount Banked', fmt_bold)
    worksheet.write(row, col + 5, result['final_total_value_1'], format_)
    worksheet.write(row, comparison_col + 4, 'Amount Banked', fmt_bold)
    worksheet.write(row, comparison_col + 5, result['final_total_value_2'], format_)

    workbook.close()


class safelist(list):
    def get(self, index, default=None):
        try:
            return self.__getitem__(index)
        except IndexError:
            return default
