import os
import numpy
import hashlib

import pandas
import xlsxwriter

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_BROKER, new_error,
                                  get_header_format, get_error_format)

HEADER_BROKER = ['Commission Type', 'Client', 'Commission Ref ID', 'Bank', 'Loan Balance',
                 'Amount Paid', 'GST Paid', 'Total Amount Paid', 'Comments']


class BrokerTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.pair = None
        self.datarows = {}
        self.datarows_count = {}
        self.summary_errors = []
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

        self.parse_rows(dataframe)

    def parse_rows(self, dataframe):
        dataframe_rows = dataframe.iloc[8:len(dataframe.index) - 1]
        dataframe_rows = dataframe_rows.rename(columns=dataframe_rows.iloc[0]).drop(dataframe_rows.index[0])
        dataframe_rows = dataframe_rows.dropna(how='all')  # remove rows that don't have any value

        for index, row in dataframe_rows.iterrows():
            invoice_row = BrokerInvoiceRow(
                row['Commission Type'], row['Client'], row['Commission Ref ID'], row['Bank'],
                row['Loan Balance'], row['Amount Paid'], row['GST Paid'],
                row['Total Amount Paid'], row['Comments'], index + 2)
            # rows[invoice_row.key_full] = invoice_row
            self.__add_datarow(invoice_row)

    def process_comparison(self, margin=0.000001):
        if self.pair is None:
            return None
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        workbook = self.create_workbook()
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        worksheet = workbook.add_worksheet()
        row = 0
        col_a = 0
        col_b = 10

        format_ = fmt_error if not self.equal_from else None
        worksheet.write(row, col_a, 'From')
        worksheet.write(row, col_a + 1, self.from_, format_)
        worksheet.write(row, col_b, 'From')
        worksheet.write(row, col_b + 1, self.pair.from_, format_)
        row += 1
        format_ = fmt_error if not self.equal_to else None
        worksheet.write(row, col_a, 'To')
        worksheet.write(row, col_a + 1, self.to, format_)
        worksheet.write(row, col_b, 'To')
        worksheet.write(row, col_b + 1, self.pair.to, format_)
        row += 1
        format_ = fmt_error if not self.equal_abn else None
        worksheet.write(row, col_a, 'ABN')
        worksheet.write(row, col_a + 1, self.abn, format_)
        worksheet.write(row, col_b, 'ABN')
        worksheet.write(row, col_b + 1, self.pair.abn, format_)
        row += 1
        format_ = fmt_error if not self.equal_bsb else None
        worksheet.write(row, col_a, 'BSB')
        worksheet.write(row, col_a + 1, self.bsb, format_)
        worksheet.write(row, col_b, 'BSB')
        worksheet.write(row, col_b + 1, self.pair.bsb, format_)
        row += 1
        format_ = fmt_error if not self.equal_account else None
        worksheet.write(row, col_a, 'Account')
        worksheet.write(row, col_a + 1, self.account, format_)
        worksheet.write(row, col_b, 'Account')
        worksheet.write(row, col_b + 1, self.pair.account, format_)
        row += 2

        for index, item in enumerate(HEADER_BROKER):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        # Code below is just to find the errors and write them into the spreadsheets
        for key in self.datarows.keys():
            self_row = self.datarows[key]
            pair_row = self.pair.datarows.get(key, None)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += BrokerInvoiceRow.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right')
            else:
                self.summary_errors += BrokerInvoiceRow.write_row(
                    worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        alone_keys_infynity = set(self.pair.datarows.keys() - set(self.datarows.keys()))
        for key in alone_keys_infynity:
            self.summary_errors += BrokerInvoiceRow.write_row(
                worksheet, self, self.pair.datarows[key], row, fmt_error, 'right')
            row += 1

        workbook.close()
        return self.summary_errors

    def create_workbook(self):
        suffix = '' if self.filename.endswith('.xlsx') else '.xlsx'
        return xlsxwriter.Workbook(OUTPUT_DIR_BROKER + 'DETAILED_' + self.filename + suffix)

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[:-6]  # Remove process ID and date stamp
        filename_forkey = '_'.join(filename_parts)

        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()

    def __add_datarow(self, row):
        if row.key in self.datarows.keys():  # If the row already exists
            self.datarows_count[row.key] += 1  # Increment row count for that key
            row.key = row._generate_key(self.datarows_count[row.key])  # Generate new key for the record
            self.datarows[row.key] = row  # Add row to the list
        else:
            self.datarows_count[row.key] = 1  # Increment row count for that key
            self.datarows[row.key] = row  # Add row to the list

    @property
    def equal_from(self):
        if self.pair is None:
            return False
        return self.from_ == self.pair.from_

    @property
    def equal_to(self):
        if self.pair is None:
            return False
        return self.to == self.pair.to

    @property
    def equal_abn(self):
        if self.pair is None:
            return False
        return self.abn == self.pair.abn

    @property
    def equal_bsb(self):
        if self.pair is None:
            return False
        return self.bsb == self.pair.bsb

    @property
    def equal_account(self):
        if self.pair is None:
            return False
        return self.account == self.pair.account


class BrokerInvoiceRow(InvoiceRow):

    def __init__(self, commission_type, client, reference_id, bank, loan_balance, amount_paid,
                 gst_paid, total_amount_paid, comments, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

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

        self._key = self._generate_key()
        self._key_full = self.__generate_key_full()

    # region Properties
    @property
    def key(self):
        return self._key

    @key.setter
    def key(self, k):
        self._key = k

    @property
    def key_full(self):
        return self._key_full

    @property
    def pair(self):
        return self._pair

    @pair.setter
    def pair(self, pair):
        self._pair = pair

    @property
    def margin(self):
        return self._margin

    @margin.setter
    def margin(self, margin):
        self._margin = margin

    @property
    def equal_bank(self):
        if self.pair is None:
            return False
        return self.bank == self.pair.bank

    @property
    def equal_loan_balance(self):
        if self.pair is None:
            return False
        return self.loan_balance == self.pair.loan_balance

    @property
    def equal_amount_paid(self):
        if self.pair is None:
            return False
        return self.amount_paid == self.pair.amount_paid

    @property
    def equal_gst_paid(self):
        if self.pair is None:
            return False
        return self.gst_paid == self.pair.gst_paid

    @property
    def equal_total_amount_paid(self):
        if self.pair is None:
            return False
        return self.total_amount_paid == self.pair.total_amount_paid

    @property
    def equal_comments(self):
        if self.pair is None:
            return False
        return self.comments == self.pair.comments
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.reference_id.encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
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

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left'):
        col = 0
        if side == 'right':
            col = 10

        worksheet.write(row, col, element.commission_type)
        worksheet.write(row, col + 1, element.client)
        worksheet.write(row, col + 2, element.reference_id)

        format_ = fmt_error if not element.equal_bank else None
        worksheet.write(row, col + 3, element.bank, format_)

        format_ = fmt_error if not element.equal_loan_balance else None
        worksheet.write(row, col + 4, element.loan_balance, format_)

        format_ = fmt_error if not element.equal_amount_paid else None
        worksheet.write(row, col + 5, element.amount_paid, format_)

        format_ = fmt_error if not element.equal_gst_paid else None
        worksheet.write(row, col + 6, element.gst_paid, format_)

        format_ = fmt_error if not element.equal_total_amount_paid else None
        worksheet.write(row, col + 7, element.total_amount_paid, format_)

        format_ = fmt_error if not element.equal_comments else None
        worksheet.write(row, col + 8, element.comments, format_)

        errors = []
        line = element.row_number
        if element.pair is not None:
            if not element.equal_bank:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Bank does not match', line, element.bank, element.pair.bank))
            if not element.equal_loan_balance:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Loan Balance does not match', line, element.loan_balance, element.pair.loan_balance))
            if not element.equal_amount_paid:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Amount Paid does not match', line, element.amount_paid, element.pair.amount_paid))
            if not element.equal_gst_paid:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Amount does not match', line, element.gst_paid, element.pair.gst_paid))
            if not element.equal_total_amount_paid:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Total Amount Paid does not match', line, element.total_amount_paid, element.pair.total_amount_paid))
            if not element.equal_comments:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Total Amount Paid does not match', line, element.comments, element.pair.comments))
        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line))

        return errors


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
