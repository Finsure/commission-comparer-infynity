import os
import numpy
import hashlib

import pandas
import xlsxwriter

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_SUMMARY_PID,
                                  OUTPUT_DIR_BROKER_PID, new_error, write_errors, worksheet_write)
from src.utils import merge_lists, safelist


class BrokerTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self._key = self.__generate_key()
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

        rows = {}
        for index, row in dataframe_rows.iterrows():
            invoice_row = BrokerInvoiceRow(
                row['Commission Type'], row['Client'], row['Commission Ref ID'], row['Bank'],
                row['Loan Balance'], row['Amount Paid'], row['GST Paid'],
                row['Total Amount Paid'], row['Comments'], index + 2)
            rows[invoice_row.key_full] = invoice_row
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

        if not has_pair:
            result_rows = {}
            for key in self.rows.keys():
                row_local = self.rows[key]
                result_rows[key] = row_local.compare_to(None)
            result['results_rows'] = result_rows
            return result

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

            # If we couldnt find the row by the BrokerInvoiceRow.full_key() it means they are different
            # so we try to locate them by the BrokerInvoiceRow.key()
            if row_local is None:
                for k in self.rows.keys():
                    if self.rows.get(k).key == row_invoice.key:
                        row_local = self.rows[k]
                        keys_all.remove(row_local.key_full)
                        use_key = k
            elif row_invoice is None:
                for k in invoice.rows.keys():
                    if invoice.rows.get(k).key == row_local.key:
                        row_invoice = invoice.rows[k]
                        keys_all.remove(row_invoice.key_full)
                        use_key = k

            if row_local is not None:
                result_rows[use_key] = row_local.compare_to(row_invoice, margin, False)
            else:
                result_rows[use_key] = row_invoice.compare_to(row_local, margin, True)

        result['results_rows'] = result_rows

        for key in result['results_rows'].keys():
            result['overall'] = result['overall'] and result['results_rows'][key]['overall']
        return result

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[:-6]  # Remove process ID and date stamp
        filename_forkey = '_'.join(filename_parts)

        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()


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
        # sha.update(self.bank.encode(ENCODING))
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
        equal_bank = False

        if has_pair:
            equal_bank = self.bank == row.bank
            equal_loan_balance = self.compare_numbers(self.loan_balance, row.loan_balance, margin)
            equal_amount_paid = self.compare_numbers(self.amount_paid, row.amount_paid, margin)
            equal_gst_paid = self.compare_numbers(self.gst_paid, row.gst_paid, margin)
            equal_total_amount_paid = self.compare_numbers(self.total_amount_paid, row.total_amount_paid, margin)
            equal_comments = self.comments == row.comments

        overall = equal_bank and equal_loan_balance and equal_amount_paid and equal_gst_paid and equal_total_amount_paid and equal_comments

        a = 'a'
        b = 'b'
        if reverse:
            a = 'b'
            b = 'a'

        result['overall'] = overall
        result['bank'] = equal_bank
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
        'bank_a': '',
        'bank_b': '',
        'loan_balance_a': '',
        'loan_balance_b': '',
        'amount_paid_a': '',
        'amount_paid_b': '',
        'gst_paid_a': '',
        'gst_paid_b': '',
        'total_amount_paid_a': '',
        'total_amount_paid_b': '',
        'comments_a': '',
        'comments_b': ''
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


def create_summary_broker(results: list):
    workbook = xlsxwriter.Workbook(OUTPUT_DIR_SUMMARY_PID + 'broker_rcti_summary.xlsx')
    worksheet = workbook.add_worksheet('Summary')

    row = 0
    col = 0

    fmt_title = workbook.add_format({'font_size': 20, 'bold': True})
    fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white',
                                            'bg_color': 'black'})

    worksheet.merge_range('A1:I1', 'Commission Broker RCTI Summary', fmt_title)
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
                msg = 'From does not match'
                error = new_error(file, msg, '', result['from_a'], result['from_b'])
                list_errors.append(error)

            # From ABN does not match
            if not result['equal_abn']:
                msg = 'ABN does not match'
                error = new_error(file, msg, '', result['abn_a'], result['abn_b'])
                list_errors.append(error)

            # To does not match
            if not result['equal_to']:
                msg = 'To does not match'
                error = new_error(file, msg, '', result['to_a'], result['to_b'])
                list_errors.append(error)

            # BSB does not match
            if not result['equal_bsb']:
                msg = 'BSB does not match'
                error = new_error(file, msg, '', result['bsb_a'], result['bsb_b'])
                list_errors.append(error)

            # Account does not match
            if not result['equal_account']:
                msg = 'Account number does not match'
                error = new_error(file, msg, '', result['account_a'], result['account_b'])
                list_errors.append(error)

            for key in result['results_rows'].keys():
                result_row = result['results_rows'][key]
                if result_row['overall'] is False:  # it means there is an issue

                    if not result_row['has_pair']:
                        msg = 'No corresponding row in comission file'
                        error = new_error(file, msg, '')
                        list_errors.append(error)
                        continue

                    values_list = safelist([])
                    if not result_row['loan_balance']:
                        values_list.append(result_row['loan_balance_a'])
                        values_list.append(result_row['loan_balance_b'])
                    if not result_row['amount_paid']:
                        values_list.append(result_row['amount_paid_a'])
                        values_list.append(result_row['amount_paid_b'])
                    if not result_row['gst_paid']:
                        values_list.append(result_row['gst_paid_a'])
                        values_list.append(result_row['gst_paid_b'])
                    if not result_row['total_amount_paid']:
                        values_list.append(result_row['total_amount_paid_a'])
                        values_list.append(result_row['total_amount_paid_b'])
                    if not result_row['comments']:
                        values_list.append(result_row['comments_a'])
                        values_list.append(result_row['comments_b'])
                    if not result_row['bank']:
                        values_list.append(result_row['bank_a'])
                        values_list.append(result_row['bank_b'])

                    msg = 'Values do not match'
                    error = new_error(file, msg, result_row['row_number'], values_list.get(0, ''),
                                      values_list.get(1, ''), values_list.get(2, ''),
                                      values_list.get(3, ''), values_list.get(4, ''),
                                      values_list.get(5, ''), values_list.get(6, ''),
                                      values_list.get(7, ''), values_list.get(8, ''),
                                      values_list.get(9, ''))
                    list_errors.append(error)

    worksheet = write_errors(list_errors, worksheet, row, col, fmt_table_header)
    workbook.close()


def _write_table_header(worksheet, row, col, format):
    worksheet.write(row, col, 'Commission Type', format)
    worksheet.write(row, col + 1, 'Client', format)
    worksheet.write(row, col + 2, 'Commission Ref ID', format)
    worksheet.write(row, col + 3, 'Bank', format)
    worksheet.write(row, col + 4, 'Loan Balance', format)
    worksheet.write(row, col + 5, 'Amount Paid', format)
    worksheet.write(row, col + 6, 'GST Paid', format)
    worksheet.write(row, col + 7, 'Total Amount Paid', format)
    worksheet.write(row, col + 8, 'Comments', format)


def create_detailed_broker(result: dict):
    if result['overall']:
        return

    suffix = '' if result['filename'].endswith('.xlsx') else '.xlsx'

    workbook = xlsxwriter.Workbook(OUTPUT_DIR_BROKER_PID + 'DETAILED_' + result['filename'] + suffix)
    worksheet = workbook.add_worksheet('Detailed')

    fmt_error = workbook.add_format({'font_color': 'red'})
    fmt_bold = workbook.add_format({'bold': True})
    fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})

    row = 0
    col_a = 0
    col_b = 10

    worksheet.merge_range('A1:N1', result['filename'])
    row += 2

    txt_from = 'From'
    format_ = fmt_error if not result['equal_from'] else None
    worksheet_write(worksheet, row, col_a, txt_from, fmt_bold, result['from_a'], format_)
    worksheet_write(worksheet, row, col_b, txt_from, fmt_bold, result['from_b'], format_)
    row += 1

    txt_to = 'To'
    format_ = fmt_error if not result['equal_to'] else None
    worksheet_write(worksheet, row, col_a, txt_to, fmt_bold, result['to_a'], format_)
    worksheet_write(worksheet, row, col_b, txt_to, fmt_bold, result['to_b'], format_)
    row += 1

    txt_abn = 'ABN'
    format_ = fmt_error if not result['equal_abn'] else None
    worksheet_write(worksheet, row, col_a, txt_abn, fmt_bold, result['to_a'], format_)
    worksheet_write(worksheet, row, col_b, txt_abn, fmt_bold, result['to_b'], format_)
    row += 1

    txt_bsb = 'BSB'
    format_ = fmt_error if not result['equal_bsb'] else None
    worksheet_write(worksheet, row, col_a, txt_bsb, fmt_bold, result['bsb_a'], format_)
    worksheet_write(worksheet, row, col_b, txt_bsb, fmt_bold, result['bsb_b'], format_)
    row += 1

    txt_account = 'Account'
    format_ = fmt_error if not result['equal_account'] else None
    worksheet_write(worksheet, row, col_a, txt_account, fmt_bold, result['account_a'], format_)
    worksheet_write(worksheet, row, col_b, txt_account, fmt_bold, result['account_b'], format_)
    row += 2

    if result['has_pair']:

        _write_table_header(worksheet, row, col_a, fmt_table_header)
        _write_table_header(worksheet, row, col_b, fmt_table_header)

        for key in result['results_rows'].keys():
            # print(key)
            row += 1
            result_row = result['results_rows'][key]

            format_ = fmt_error if not result_row['has_pair'] else None
            worksheet.write(row, col_a, result_row['commission_type_a'], format_)
            worksheet.write(row, col_a + 1, result_row['client_a'], format_)
            worksheet.write(row, col_a + 2, result_row['reference_id_a'], format_)

            worksheet.write(row, col_b, result_row['commission_type_b'], format_)
            worksheet.write(row, col_b + 1, result_row['client_b'], format_)
            worksheet.write(row, col_b + 2, result_row['reference_id_b'], format_)

            format_ = fmt_error if not result_row['bank'] else None
            worksheet.write(row, col_a + 3, result_row['bank_a'], format_)
            worksheet.write(row, col_b + 3, result_row['bank_b'], format_)

            format_ = fmt_error if not result_row['loan_balance'] else None
            worksheet.write(row, col_a + 4, result_row['loan_balance_a'], format_)
            worksheet.write(row, col_b + 4, result_row['loan_balance_b'], format_)

            format_ = fmt_error if not result_row['amount_paid'] else None
            worksheet.write(row, col_a + 5, result_row['amount_paid_a'], format_)
            worksheet.write(row, col_b + 5, result_row['amount_paid_b'], format_)

            format_ = fmt_error if not result_row['gst_paid'] else None
            worksheet.write(row, col_a + 6, result_row['gst_paid_a'], format_)
            worksheet.write(row, col_b + 6, result_row['gst_paid_b'], format_)

            format_ = fmt_error if not result_row['total_amount_paid'] else None
            worksheet.write(row, col_a + 7, result_row['total_amount_paid_a'], format_)
            worksheet.write(row, col_b + 7, result_row['total_amount_paid_b'], format_)

            format_ = fmt_error if not result_row['comments'] else None
            worksheet.write(row, col_a + 8, result_row['comments_a'], format_)
            worksheet.write(row, col_b + 8, result_row['comments_b'], format_)

    else:
        worksheet.write(row, col_a, 'No match to compare to', fmt_error)

    workbook.close()


def read_files_broker(dir_: str, files: list) -> dict:
    keys = {}
    for file in files:
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = BrokerTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return keys
