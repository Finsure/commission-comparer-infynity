
import numpy
import pandas
import xlrd
import copy
import hashlib

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, new_error, OUTPUT_DIR_EXEC_SUMMARY,
                                  get_header_format, get_error_format)
from src import utils as u
from src.utils import bcolors

HEADER_LENDER = ['Bank', 'Bank Detailed Name', 'Settlement Amount', 'Commission Amount (Excl GST)',
                 'GST', 'Commission Amount Inc GST']


class ExecutiveSummary(TaxInvoice):

    # The tabs list is mapped FROM Infynity TO LoanKit files
    # 'Infynity Tab': 'Loankit tab where that information is located'
    TABS = {
        'Branch Summary Report': 'Branch Summary Report',
        'Branch Fee Summary Report': 'Branch Summary Report'
    }

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.datarows_branch_summary = {}
        self.datarows_branch_fee_summary = {}
        self.datarows_broker_summary = {}
        self.datarows_broker_fee_summary = {}
        self.datarows_lender_upfront = {}
        self.datarows_lender_trail = {}
        self.datarows_lender_vbi = {}
        self.summary_errors = []  # List of errors found during the comparison
        self.pair = None
        self.margin = 0  # margin of error acceptable for numeric comprisons
        self.parse()

    def __add_datarow(self, datarows_dict, counter_dict, row):
        if row.key_full in datarows_dict.keys():  # If the row already exists
            counter_dict[row.key_full] += 1  # Increment row count for that key_full
            row.key_full = row._generate_key(counter_dict[row.key_full])  # Generate new key_full for the record
            datarows_dict[row.key_full] = row  # Add row to the list
        else:
            counter_dict[row.key_full] = 0  # Start counter
            datarows_dict[row.key_full] = row  # Add row to the list

    def parse(self):
        xl = pandas.ExcelFile(self.full_path)
        self.datarows_lender_upfront = self.parse_lender(xl, 'Lender Upfront Records')
        self.datarows_lender_trail = self.parse_lender(xl, 'Lender Trail Records')
        self.datarows_lender_vbi = self.parse_lender(xl, 'Lender VBI Records')
        self.datarows_branch_summary = self.parse_branch(xl, 'Branch Summary Report')
        self.datarows_branch_fee_summary = self.parse_branch(xl, 'Branch Fee Summary Report')
        self.datarows_broker_summary = self.parse_broker(xl, 'Broker Summary Report')
        self.datarows_broker_fee_summary = self.parse_broker(xl, 'Broker Fee Summary Report')

    def parse_lender(self, xl, tab):
        df = xl.parse(tab)
        df = df.dropna(how='all')
        df = self.general_replaces(df)
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])

        rows_counter = {}
        rows = {}
        for index, row in df.iterrows():
            lsum_row = LenderExecutiveSummaryRow(
                row['Bank'], row['Bank Detailed Name'], row['Settlement Amount'],
                row['Commission Amount (Excl GST)'], row['GST'], row['Commission Amount Incl. GST'],
                index)
            self.__add_datarow(rows, rows_counter, lsum_row)

        return rows

    def parse_branch(self, xl, tab):
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.dropna(how='all')  # remove rows that don't have any value
            df = self.general_replaces(df)
            df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header

            # TODO: Remove the code below once received the updated version of the report
            if 'Compliance Fee GST' in df.columns:
                df['Compliance GST'] = df['Compliance Fee GST']
                del df['Compliance Fee GST']
            # TODO: Remove the code below once received the updated version of the report
            if 'Compliance Fee Excl. GST' in df.columns:
                df['Compliance Excl. GST'] = df['Compliance Fee Excl. GST']
                del df['Compliance Fee Excl. GST']
            # TODO: Remove the code below once received the updated version of the report
            if 'Compliance Fee Incl. GST' in df.columns:
                df['Compliance Incl. GST'] = df['Compliance Fee Incl. GST']
                del df['Compliance Fee Incl. GST']
            # TODO: Remove the code below once received the updated version of the report
            if 'Conference fee' in df.columns:
                df['Conference fee GST'] = df['Conference fee']
                del df['Conference fee']
            # TODO: Remove the code below once received the updated version of the report
            if 'Fee Adjustment' in df.columns:
                df['Fee Adjustment GST'] = df['Fee Adjustment']
                del df['Fee Adjustment']
            # TODO: Remove the code below once received the updated version of the report
            if 'Goldfields Loan Repayment Finsure' in df.columns:
                df['Goldfields Loan Repayment Finsure GST'] = df['Goldfields Loan Repayment Finsure']
                del df['Goldfields Loan Repayment Finsure']
            # TODO: Remove the code below once received the updated version of the report
            if 'RP Data' in df.columns:
                df['RP Data GST'] = df['RP Data']
                del df['RP Data']
            # TODO: Remove the code below once received the updated version of the report
            if 'Software Fee' in df.columns:
                df['Software Fee GST'] = df['Software Fee']
                del df['Software Fee']

            for index, row in df.iterrows():
                drow = df.loc[df['ID'] == row['ID']].to_dict(orient='records')[0]
                drow['line'] = index
                if drow['ID'] != 'Total':
                    drow['ID'] = int(drow['ID'])
                rows[drow['ID']] = drow
        except Exception:  # Exception if tab is not found
            pass
        return rows

    def parse_broker(self, xl, tab):
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.loc[:, ~df.columns.duplicated()]  # Remove duplicate columns
            df = df.dropna(how='all')  # remove rows that don't have any value
            df = self.general_replaces(df)
            if tab in ['Broker Summary Report']:
                df = df.rename(columns=df.iloc[1]).drop(df.index[0]).drop(df.index[1])  # Make first row the table header
            else:
                df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header

            if 'Broker Name (ID)' in list(df):
                df['Broker Name'] = ''
                df['Broker ID'] = ''
                df['Branch Name'] = ''
                df['Branch ID'] = ''
                for index, row in df.iterrows():
                    try:
                        row['Broker Name'] = row['Broker Name (ID)'].rsplit('(', 1)[0].strip()
                        row['Broker ID'] = row['Broker Name (ID)'].rsplit('(', 1)[1][:-1]
                        row['Branch Name'] = row['Branch Name (ID)'].rsplit('(', 1)[0].strip()
                        row['Branch ID'] = row['Branch Name (ID)'].rsplit('(', 1)[1][:-1]
                    except IndexError:
                        row['Broker Name'] = 'Total'
                        row['Broker ID'] = 'Total'
                        row['Branch Name'] = 'Total'
                        row['Branch ID'] = 'Total'

                    df.loc[index].at['Broker Name'] = row['Broker Name']
                    df.loc[index].at['Broker ID'] = row['Broker ID']
                    df.loc[index].at['Branch Name'] = row['Branch Name']
                    df.loc[index].at['Branch ID'] = row['Branch ID']
                df.drop(['Broker Name (ID)'], axis=1)
                df.drop(['Branch Name (ID)'], axis=1)

            field_id = 'Broker ID'
            for index, row in df.iterrows():
                drow = df.loc[df[field_id] == row[field_id]].to_dict(orient='records')[0]
                drow['line'] = index
                rows[drow[field_id]] = drow
        except xlrd.biffh.XLRDError:
            pass
        return rows

    def process_comparison(self, margin=0.000001):
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        if self.pair is None:
            return None

        workbook = self.create_workbook(OUTPUT_DIR_EXEC_SUMMARY)
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        self.process_lender(workbook, 'Lender Upfront Records', self.datarows_lender_upfront,
                            self.pair.datarows_lender_upfront, fmt_table_header, fmt_error)

        self.process_lender(workbook, 'Lender Trail Records', self.datarows_lender_trail,
                            self.pair.datarows_lender_trail, fmt_table_header, fmt_error)

        self.process_lender(workbook, 'Lender VBI Records', self.datarows_lender_vbi,
                            self.pair.datarows_lender_vbi, fmt_table_header, fmt_error)

        self.process_generic(workbook, 'Branch Summary Report', self.datarows_branch_summary,
                             self.pair.datarows_branch_summary, fmt_table_header, fmt_error)

        self.process_generic(workbook, 'Branch Fee Summary Report', self.datarows_branch_fee_summary,
                             self.pair.datarows_branch_summary, fmt_table_header, fmt_error)

        self.process_generic(workbook, 'Broker Summary Report', self.datarows_broker_summary,
                             self.pair.datarows_broker_summary, fmt_table_header, fmt_error)

        self.process_generic(workbook, 'Broker Fee Summary Report', self.datarows_broker_fee_summary,
                             self.pair.datarows_broker_summary, fmt_table_header, fmt_error)

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def process_generic(self, workbook, tab, dict_a, dict_b, fmt_table_header, fmt_error):
        worksheet = workbook.add_worksheet(tab)

        # This return an arbitrary element from the dictionary so we can get the headers
        header = copy.copy(next(iter(dict_a.values())))
        del header['line']

        row = 0
        col_a = 0
        col_b = len(header.keys()) + 1

        for index, item in enumerate(header.keys()):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(dict_b.keys()) - set(dict_a.keys())

        for key in dict_a.keys():
            self_row = dict_a[key]
            pair_row = dict_b.get(key, None)

            self.summary_errors += comapre_dicts(
                worksheet, row, self_row, pair_row, self.margin, self.filename, self.pair.filename,
                fmt_error, tab)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += comapre_dicts(
                worksheet, row, None, dict_b[key], self.margin,
                self.filename, self.pair.filename, fmt_error, tab)
            row += 1

    def process_lender(self, workbook, tab, datarows, datarows_pair, fmt_table_header, fmt_error):
        worksheet = workbook.add_worksheet(tab)
        row = 0
        col_a = 0
        col_b = len(HEADER_LENDER) + 1

        for index, item in enumerate(HEADER_LENDER):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(datarows_pair.keys()) - set(datarows.keys())

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in datarows.keys():
            self_row = datarows[key_full]
            self_row.margin = self.margin

            pair_row = self.find_pair_row(datarows_pair, self_row)
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del datarows_pair[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = self.margin
                pair_row.pair = self_row
                self.summary_errors += LenderExecutiveSummaryRow.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += LenderExecutiveSummaryRow.write_row(worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += LenderExecutiveSummaryRow.write_row(
                worksheet, self, datarows_pair[key], row, fmt_error, 'right', write_errors=False)
            row += 1

    def find_pair_row(self, datarows_pair, row):
        # Match by full_key
        pair_row = datarows_pair.get(row.key_full, None)
        if pair_row is not None:
            return pair_row

        # We want to match by similarity before matching by the key
        # Match by similarity
        for _, item in datarows_pair.items():
            if row.equals(item):
                return item

        # Match by key
        for _, item in datarows_pair.items():
            if row.key == item.key:
                return item

        # Return None if nothing found
        return None

    def new_error(self, msg, line_a='', line_b='', value_a='', value_b='', tab=''):
        return new_error(self.filename, self.pair.filename, msg, line_a, line_b, value_a, value_b, tab)

    def general_replaces(self, df):
        df = df.replace(numpy.nan, '', regex=True)  # remove rows that don't have any value
        df = df.replace(' Inc GST', ' Incl. GST', regex=True)
        df = df.replace(' Exc GST', ' Excl. GST', regex=True)
        df = df.replace('Pmt ', 'Payment ', regex=True)
        return df

    def parse_broker_name(self, val):
        if len(val) == 0:
            return val
        return val.split('(')[0].strip()


class LenderExecutiveSummaryRow(InvoiceRow):

    def __init__(self, bank, bank_detailed_name, settlement_amount, commission_amount_exc_gst, gst,
                 commission_amount_inc_gst, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

        self.bank = bank
        self.bank_detailed_name = bank_detailed_name
        self.settlement_amount = settlement_amount
        self.commission_amount_exc_gst = commission_amount_exc_gst
        self.gst = gst
        self.commission_amount_inc_gst = commission_amount_inc_gst

        self.row_number = row_number

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
        return u.sanitize(self.bank) == u.sanitize(self.pair.bank)

    @property
    def equal_bank_detailed_name(self):
        if self.pair is None:
            return False
        return u.sanitize(self.bank_detailed_name) == u.sanitize(self.pair.bank_detailed_name)

    @property
    def equal_settlement_amount(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.settlement_amount, self.pair.settlement_amount, self.margin)

    @property
    def equal_commission_amount_exc_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission_amount_exc_gst, self.pair.commission_amount_exc_gst, self.margin)

    @property
    def equal_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst, self.pair.gst, self.margin)

    @property
    def equal_commission_amount_inc_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission_amount_inc_gst, self.pair.commission_amount_inc_gst, self.margin)
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.bank).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(self.bank.encode(ENCODING))
        sha.update(self.bank_detailed_name.encode(ENCODING))
        sha.update(str(self.settlement_amount).encode(ENCODING))
        sha.update(str(self.commission_amount_exc_gst).encode(ENCODING))
        sha.update(str(self.gst).encode(ENCODING))
        sha.update(str(self.commission_amount_inc_gst).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != LenderExecutiveSummaryRow:
            return False

        return (
            u.sanitize(self.bank) == u.sanitize(obj.bank)
            and u.sanitize(self.bank_detailed_name) == u.sanitize(obj.bank_detailed_name)
            and self.compare_numbers(self.settlement_amount, obj.settlement_amount, self.margin)
            and self.compare_numbers(self.commission_amount_exc_gst, obj.commission_amount_exc_gst, self.margin)
            and self.compare_numbers(self.gst, obj.gst, self.margin)
            and self.compare_numbers(self.commission_amount_inc_gst, obj.commission_amount_inc_gst, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = len(HEADER_LENDER) + 1

        worksheet.write(row, col, element.bank)

        format_ = fmt_error if not element.equal_bank_detailed_name else None
        worksheet.write(row, col + 1, element.bank_detailed_name)

        format_ = fmt_error if not element.equal_settlement_amount else None
        worksheet.write(row, col + 2, element.settlement_amount, format_)

        format_ = fmt_error if not element.equal_commission_amount_exc_gst else None
        worksheet.write(row, col + 3, element.commission_amount_exc_gst, format_)

        format_ = fmt_error if not element.equal_gst else None
        worksheet.write(row, col + 4, element.gst, format_)

        format_ = fmt_error if not element.equal_commission_amount_inc_gst else None
        worksheet.write(row, col + 5, element.commission_amount_inc_gst, format_)

        errors = []
        line_a = element.row_number
        description = f"Bank: {element.bank}"
        if element.pair is not None:
            line_b = element.pair.row_number
            if write_errors:
                if not element.equal_bank_detailed_name:
                    msg = 'Detailed Bank Name does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.bank_detailed_name, element.pair.bank_detailed_name))

                if not element.equal_settlement_amount:
                    msg = 'Settlement Amount does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.settlement_amount, element.pair.settlement_amount))

                if not element.equal_commission_amount_exc_gst:
                    msg = 'Commission Amount (Excl GST) does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.commission_amount_exc_gst, element.pair.commission_amount_exc_gst))

                if not element.equal_gst:
                    msg = 'Amount does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.gst, element.pair.gst))

                if not element.equal_commission_amount_inc_gst:
                    msg = 'Total Amount Paid does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.commission_amount_inc_gst, element.pair.commission_amount_inc_gst))

        else:
            if write_errors:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line_a, '', value_a=description))
            else:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', '', line_a, value_b=description))

        return errors


def comapre_dicts(worksheet, row, row_a, row_b, margin, filename_a, filename_b, fmt_error, tab):
    errors = []
    if row_b is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', row_a['line'], '', tab=tab))
        return errors
    elif row_a is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', '', row_b['line'], tab=tab))
        return errors

    col_a = 0
    col_b = len(row_a.keys())  # + 1

    for index, column in enumerate(row_a.keys()):
        if column == 'line':  # if we evere remove this condition don't forget to add + 1 to 2 lines above
            continue

        val_a = str(row_a[column])
        try:
            val_a = u.money_to_float(val_a)
        except ValueError:
            pass

        if row_b is not None:
            val_b = str(row_b[column]) if row_b.get(column, None) is not None else None
        else:
            val_b = None

        if val_b is None:
            errors.append(new_error(filename_a, filename_b, f'No corresponding column ({column}) in commission file', tab=tab))
            worksheet.write(row, col_a, val_a, fmt_error)
            col_a += 1
            col_b += 1
            continue

        try:
            val_b = u.money_to_float(val_b)
        except ValueError:
            pass

        fmt = None
        if not compare_values(val_a, val_b, margin):
            fmt = fmt_error
            errors.append(new_error(
                filename_a, filename_b, f'Value of {column} do not match', row_a['line'],
                row_b['line'], val_a, val_b, tab=tab))

        worksheet.write(row, col_a, row_a[column], fmt)
        worksheet.write(row, col_b, row_b[column], fmt)
        col_a += 1
        col_b += 1

    return errors


def compare_values(val_a, val_b, margin):
    if type(val_a) == float and type(val_b) == float:
        return u.compare_numbers(val_a, val_b, margin)
    else:
        return u.sanitize(val_a) == u.sanitize(val_b)


def read_file_exec_summary(file: str):
    print(f'Parsing executive summary file {bcolors.BLUE}{file}{bcolors.ENDC}')
    filename = file.split('/')[-1]
    dir_ = '/'.join(file.split('/')[:-1]) + '/'
    return ExecutiveSummary(dir_, filename)
