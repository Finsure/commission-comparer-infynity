import copy
import hashlib

import xlsxwriter
from bs4 import BeautifulSoup

from src.utils import merge_lists

ENCODING = 'utf-8'


class TaxInvoice:

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
                row = InvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                 tds[3].text, tds[4].text, tds[5].text, row_number)
                rows[row.key_full()] = row
            except IndexError as e:
                # print(self.filename)
                # print("NO REFERRER ERROR: " + str(e))
                row = InvoiceRow(tds[0].text, tds[1].text, '',
                                 tds[2].text, tds[3].text, tds[4].text, row_number)
                rows[row.key_full()] = row
        return rows

    def key(self):
        if self._key is None:
            self._key = self.__generate_key()
        return self._key

    def serialize(self):
        text = self.filetext
        self.filetext = None
        serialized_obj = copy.copy(self.__dict__)
        self.filetext = text
        return serialized_obj

    def compare_to(self, invoice, margin=0.0000001):  # noqa F821
        result = result_invoice()
        result['filename'] = self.filename
        result['file'] = self.get_full_path()
        if invoice is None:
            return result

        #  If we reached here it means the file has a pair
        result['has_pair'] = True
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
        result['from_value_2'] = invoice._from
        result['from_abn_value_1'] = self.from_abn
        result['from_abn_value_2'] = invoice.from_abn
        result['to_value_1'] = self.to
        result['to_value_2'] = invoice.to
        result['to_abn_value_1'] = self.to_abn
        result['to_abn_value_2'] = invoice.to_abn
        result['bsb_value_1'] = self.bsb
        result['bsb_value_2'] = invoice.bsb
        result['account_value_1'] = self.account
        result['account_value_2'] = invoice.account
        result['final_total_value_1'] = self.final_total
        result['final_total_value_2'] = invoice.final_total

        result['overall'] = (result['equal_from'] and result['equal_from_abn']
                             and result['equal_to'] and result['equal_to_abn']
                             and result['equal_bsb'] and result['equal_account']
                             and result['equal_final_total'] and result['equal_amount_rows'])

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

            # If we couldnt find the row by the InvoiceRow.full_key() it means they are different
            # so we try to locate them by the InvoiceRow.key()
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
                result_rows[use_key] = row_local.compare_to(row_invoice)
            else:
                result_rows[use_key] = row_invoice.compare_to(row_local)

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
        sha.update(self.filename.encode(ENCODING))
        return sha.hexdigest()


class InvoiceRow:

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

    def compare_to(self, row, margin=0.0000001):  # noqa F821
        result = result_row()
        result['row_number'] = self.row_number
        if row is None:
            return result

        result['has_pair'] = True

        equal_commission_type = self.commission_type == row.commission_type
        equal_client = self.client == row.client
        equal_referrer = self.referrer == row.referrer
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

        result['commission_type_value_1'] = self.commission_type
        result['commission_type_value_2'] = row.commission_type
        result['client_value_1'] = self.client
        result['client_value_2'] = row.client
        result['referrer_value_1'] = self.referrer
        result['referrer_value_1'] = row.referrer
        result['amount_paid_value_1'] = self.amount_paid
        result['amount_paid_value_2'] = row.amount_paid
        result['gst_paid_value_1'] = self.gst_paid
        result['gst_paid_value_2'] = row.gst_paid
        result['total_value_1'] = self.total
        result['total_value_2'] = row.total
        result['overall'] = overall
        result['commission_type'] = equal_commission_type
        result['client'] = equal_client
        result['referrer'] = equal_referrer
        result['amount_paid'] = equal_amount_paid
        result['gst_paid'] = equal_gst_paid
        result['total'] = equal_total

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


def result_invoice():
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
        'results_rows': [],
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


def result_row():
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
    workbook = xlsxwriter.Workbook('referrer_rcti_summary.xlsx')
    worksheet = workbook.add_worksheet('Summary')

    row = 0
    col = 0

    worksheet.write(row, col, 'Commission Referrer RCTI Summary')
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
                error['msg'] = 'From name does not match'
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
                        msg = 'No corresponding row in comission file.'
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
    worksheet.write(row, col, 'File')
    worksheet.write(row, col + 1, 'Message')
    worksheet.write(row, col + 2, 'Line')
    worksheet.write(row, col + 3, 'First Value')
    worksheet.write(row, col + 4, 'First To Compare')
    worksheet.write(row, col + 5, 'Second Value')
    worksheet.write(row, col + 6, 'Second To Compare')
    worksheet.write(row, col + 7, 'Third Value')
    worksheet.write(row, col + 8, 'Third To Compare')
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


class safelist(list):
    def get(self, index, default=None):
        try:
            return self.__getitem__(index)
        except IndexError:
            return default
