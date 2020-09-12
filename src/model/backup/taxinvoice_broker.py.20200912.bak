import os
import pdb
import numpy
import hashlib

import pandas

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_BROKER, new_error,
                                  get_header_format, get_error_format)
from src import utils as u
from src.utils import bcolors

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

        dataframe_info = dataframe.replace(numpy.nan, '', regex=True)
        dataframe_broker_info = dataframe_info.iloc[2:5, 0:2]

        account_info = dataframe_info.iloc[len(dataframe_info.index) - 1][1]
        account_info_parts = str(account_info).split(':')

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
        dataframe_rows = dataframe_rows.replace(numpy.nan, '', regex=True)

        for index, row in dataframe_rows.iterrows():
            invoice_row = BrokerInvoiceRow(
                row['Commission Type'], row['Client'], row['Commission Ref ID'], row['Bank'],
                row['Loan Balance'], row['Amount Paid'], row['GST Paid'],
                row['Total Amount Paid'], row['Comments'], index + 2)
            self.__add_datarow(invoice_row)

    def process_comparison(self, margin=0.000001):
        if self.pair is None:
            return None
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        workbook = self.create_workbook(OUTPUT_DIR_BROKER)
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

        if not self.equal_from:
            self.summary_errors.append(new_error(
                self.filename, self.pair.filename, 'From does not match', '', '', self.from_, self.pair.from_))
        if not self.equal_to:
            self.summary_errors.append(new_error(
                self.filename, self.pair.filename, 'To does not match', '', '', self.to, self.pair.to))
        if not self.equal_abn:
            self.summary_errors.append(new_error(
                self.filename, self.pair.filename, 'ABN does not match', '', '', self.abn, self.pair.abn))
        if not self.equal_bsb:
            self.summary_errors.append(new_error(
                self.filename, self.pair.filename, 'BSB does not match', '', '', self.bsb, self.pair.bsb))
        if not self.equal_account:
            self.summary_errors.append(new_error(
                self.filename, self.pair.filename, 'Account does not match', '', '', self.account, self.pair.account))

        row += 2

        for index, item in enumerate(HEADER_BROKER):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.datarows.keys() - set(self.datarows.keys()))

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in self.datarows.keys():
            self_row = self.datarows[key_full]
            self_row.margin = margin

            pair_row = self.find_pair_row(self_row)
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.datarows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += BrokerInvoiceRow.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += BrokerInvoiceRow.write_row(worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += BrokerInvoiceRow.write_row(
                worksheet, self, self.pair.datarows[key], row, fmt_error, 'right', write_errors=False)
            row += 1

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def find_pair_row(self, row):
        # Match by full_key
        pair_row = self.pair.datarows.get(row.key_full, None)
        if pair_row is not None:
            return pair_row

        # We want to match by similarity before matching by the key
        # Match by similarity
        for _, item in self.pair.datarows.items():
            if row.equals(item):
                return item

        # Match by key
        for _, item in self.pair.datarows.items():
            if row.key == item.key:
                return item

        # Return None if nothing found
        return None

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[:-6]  # Remove process ID and date stamp
        filename_forkey = ''.join(filename_parts)

        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()

    def __add_datarow(self, row):
        if row.key_full in self.datarows.keys():  # If the row already exists
            self.datarows_count[row.key_full] += 1  # Increment row count for that key_full
            row.key_full = row._generate_key(self.datarows_count[row.key_full])  # Generate new key_full for the record
            self.datarows[row.key_full] = row  # Add row to the list
        else:
            self.datarows_count[row.key_full] = 0  # Start counter
            self.datarows[row.key_full] = row  # Add row to the list

    @property
    def equal_from(self):
        if self.pair is None:
            return False
        return u.sanitize(self.from_) == u.sanitize(self.pair.from_)

    @property
    def equal_to(self):
        if self.pair is None:
            return False
        return u.sanitize(self.to) == u.sanitize(self.pair.to)

    @property
    def equal_abn(self):
        if self.pair is None:
            return False
        return u.sanitize(self.abn) == u.sanitize(self.pair.abn)

    @property
    def equal_bsb(self):
        if self.pair is None:
            return False
        return u.sanitize(self.bsb) == u.sanitize(self.pair.bsb)

    @property
    def equal_account(self):
        if self.pair is None:
            return False
        return u.sanitize(self.account) == u.sanitize(self.pair.account)


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
        self._key_full = self._generate_key_full()

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

    @key_full.setter
    def key_full(self, k):
        self._key_full = k

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
        bank_a = u.bank_fullname(self.bank)
        bank_b = u.bank_fullname(self.pair.bank)
        return u.sanitize(bank_a) == u.sanitize(bank_b)

    @property
    def equal_loan_balance(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.loan_balance, self.pair.loan_balance, self.margin)

    @property
    def equal_amount_paid(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.amount_paid, self.pair.amount_paid, self.margin)

    @property
    def equal_gst_paid(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst_paid, self.pair.gst_paid, self.margin)

    @property
    def equal_total_amount_paid(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.total_amount_paid, self.pair.total_amount_paid, self.margin)

    @property
    def equal_comments(self):
        if self.pair is None:
            return False
        return u.sanitize(self.comments) == u.sanitize(self.pair.comments)
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.commission_type).encode(ENCODING))
        sha.update(u.sanitize(self.client).encode(ENCODING))
        sha.update(self.reference_id.encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.reference_id.encode(ENCODING))
        # sha.update(self.bank.encode(ENCODING))
        sha.update(self.loan_balance.encode(ENCODING))
        sha.update(self.amount_paid.encode(ENCODING))
        sha.update(self.gst_paid.encode(ENCODING))
        sha.update(self.total_amount_paid.encode(ENCODING))
        sha.update(self.comments.encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != BrokerInvoiceRow:
            return False

        return (
            u.sanitize(self.commission_type) == u.sanitize(obj.commission_type)
            and u.sanitize(self.client) == u.sanitize(obj.client)
            and u.sanitize(self.reference_id) == u.sanitize(obj.reference_id)
            and u.sanitize(u.bank_fullname(self.bank)) == u.sanitize(u.bank_fullname(obj.bank))
            and self.compare_numbers(self.loan_balance, obj.loan_balance, self.margin)
            and self.compare_numbers(self.amount_paid, obj.amount_paid, self.margin)
            and self.compare_numbers(self.gst_paid, obj.gst_paid, self.margin)
            and self.compare_numbers(self.total_amount_paid, obj.total_amount_paid, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
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
        line_a = element.row_number
        description = f"Reference ID: {element.reference_id}"
        if element.pair is not None:
            line_b = element.pair.row_number
            if write_errors:
                if not element.equal_bank:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Bank does not match', line_a, line_b, element.bank, element.pair.bank))

                if not element.equal_loan_balance:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Loan Balance does not match', line_a, line_b, element.loan_balance, element.pair.loan_balance))

                if not element.equal_amount_paid:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Amount Paid does not match', line_a, line_b, element.amount_paid, element.pair.amount_paid))

                if not element.equal_gst_paid:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Amount does not match', line_a, line_b, element.gst_paid, element.pair.gst_paid))

                if not element.equal_total_amount_paid:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Total Amount Paid does not match', line_a, line_b, element.total_amount_paid, element.pair.total_amount_paid))

                if not element.equal_comments:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Total Amount Paid does not match', line_a, line_b, element.comments, element.pair.comments))
        else:
            if write_errors:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line_a, '', value_a=description))
            else:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', '', line_a, value_b=description))

        return errors


def read_files_broker(dir_: str, files: list) -> dict:
    keys = {}
    counter = 1
    for file in files:
        print(f'Parsing {counter} of {len(files)} files from {bcolors.BLUE}{dir_}{bcolors.ENDC}', end='\r')
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = BrokerTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
        counter += 1
    print()
    return keys
