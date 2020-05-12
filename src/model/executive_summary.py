
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
HEADER_EXECUTIVE_SUMMARY = ['Description', 'Value']
HEADER_REFERRER = ['Branch ID', 'Branch Company Name', 'Referrer Name', 'Opening Balance',
                   'Commission Amount Paid Incl. GST', 'Total Banked Amount', 'Closing Balance']
HEADER_DE = ['Aggregator', 'Aggregator BSB No#', 'Aggregator Acct No#', 'Branch ID', 'Agent Type', 'Company Name',
             'Bank Account Name', 'BSB No#', 'Account No#', 'Amount Banked']


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
        self.datarows_executive_summary = {}
        self.datarows_referrer = {}
        self.datarows_de_file_entries = {}
        self.datarows_de_file_notpaid = {}
        self.datarows_fee = {}
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
        self.datarows_executive_summary = self.parse_executive_summary(xl, 'Executive Summary Report')
        self.datarows_referrer = self.parse_referrer(xl, 'Referrer Summary Report')
        self.datarows_de_file_entries = self.parse_de(xl, 'DE File Entries')
        self.datarows_de_file_notpaid = self.parse_de(xl, 'DE File - Amount Not Paid')
        self.datarows_fee = self.parse_executive_summary(xl, 'Fee Summary Report')

    def parse_referrer(self, xl, tab):
        df = xl.parse(tab)
        df = df[3:len(df)]
        df = df.dropna(how='all')
        df = self.general_replaces(df)
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])

        replaces = {
            'Opening Carried Forward Balance': 'Opening Balance',
            'Closing Carried Forward Balance': 'Closing Balance',
            'Payment': 'Commission Amount Paid Incl. GST'
        }
        df = self.replace_keys(replaces, df)

        if 'Branch Name (ID)' in list(df.columns):
            df['Branch ID'] = ''
            df['Branch Company Name'] = ''
            df['Referrer Name'] = ''
            for index, row in df.iterrows():
                try:
                    row['Branch ID'] = row['Branch Name (ID)'].rsplit('(', 1)[1][:-1]
                    row['Branch Company Name'] = row['Branch Name (ID)'].rsplit('(', 1)[0].strip()
                    row['Referrer Name'] = row['Referrer Name (ID)'].rsplit('(', 1)[0].strip()
                except IndexError:
                    row['Branch ID'] = 'Total'
                    row['Branch Company Name'] = 'Total'
                    row['Referrer Name'] = 'Total'
                df.loc[index].at['Branch ID'] = row['Branch ID']
                df.loc[index].at['Branch Company Name'] = row['Branch Company Name']
                df.loc[index].at['Referrer Name'] = row['Referrer Name']
            df = df.drop(['Branch Name (ID)'], axis=1)
            df = df.drop(['Referrer Name (ID)'], axis=1)
            df = df.drop(['Invoice Number'], axis=1)
        else:
            df = df.drop(['Referrer Key'], axis=1)
            df = df.drop(['Commission Amount Paid Excl. GST'], axis=1)
            df = df.drop(['Commission Amount Paid GST'], axis=1)

        rows = {}
        rows_counter = {}
        for index, row in df.iterrows():
            rsum_row = ReferrerExecutiveSummaryRow(
                row['Branch ID'], row['Branch Company Name'], row['Referrer Name'], row['Opening Balance'],
                row['Commission Amount Paid Incl. GST'], row['Total Banked Amount'], row['Closing Balance'], index)
            self.__add_datarow(rows, rows_counter, rsum_row)

        return rows

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
                row['Commission Amount (Excl GST)'], row['GST'], row['Commission Amount Incl. GST'], index)
            self.__add_datarow(rows, rows_counter, lsum_row)

        return rows

    def parse_branch(self, xl, tab):
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.dropna(how='all')  # remove rows that don't have any value
            df = self.general_replaces(df)
            df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header

            replaces = {
                'Upfront Rec Excl. GST': 'Upfront Commission Excl. GST',
                'Upfront Records GST': 'Upfront Commission GST',
                'Upfront Records Incl. GST': 'Upfront Commission Incl. GST',
                'Trail Rec Excl. GST': 'Trail Commission Excl. GST',
                'Trail Records GST': 'Trail Commission GST',
                'Trail Records Incl. GST': 'Trail Commission Incl. GST',
                'VBI Rec Excl. GST': 'VBI Commission Excl. GST',
                'VBI Records GST': 'VBI Commission GST',
                'VBI Records Incl. GST': 'VBI Commission Incl. GST',
                'Total Commission': 'Total Commission Received',
                'ID': 'Branch ID',
                'Opening CFB': 'Branch Opening Carried Forward Balance'
            }
            df = self.replace_keys(replaces, df)

            columns_to_remove = [
                'Brokers Opening Carried Forward Balance Incl. GST',
                'Brokers Upfront Amount Calculated Incl. GST',
                'Brokers Trail Amount Calculated Incl. GST',
                'Brokers VBI Amount Calculated Incl. GST',
                'Brokers Fee Charged Incl. GST',
                'Brokers Amount Calculated Incl. GST',
                'Brokers Closing Carried Forward Balance Incl. GST',
                'Referrers Amount Calculated Incl. GST',
                'Total Branch Fee Excl. GST',
                'Total Branch Fee Charge GST',
                'Total Branch Fee Charge Incl. GST',
                'Branch Closing Carried Forward Balance Incl. GST',
                'Amount Retained by Branch Incl. GST',
                'Total Commission Paid To Branch'
            ]
            for column in columns_to_remove:
                try:
                    del df[column]
                except KeyError:
                    pass

            for index, row in df.iterrows():
                drow = df.loc[df['Branch ID'] == row['Branch ID']].to_dict(orient='records')[0]
                drow['line'] = index
                if drow['Branch ID'] not in ['Total', '']:
                    drow['Branch ID'] = int(drow['Branch ID'])
                rows[drow['Branch ID']] = drow
        except xlrd.biffh.XLRDError:  # Exception if tab is not found
            print(f"{bcolors.YELLOW}No sheet named {tab} found in {bcolors.BLUE}{self.full_path}{bcolors.ENDC}")
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
                df = df.drop(['Broker Name (ID)'], axis=1)
                df = df.drop(['Branch Name (ID)'], axis=1)

            replaces = {
                'Opening Carried Forward Balance': 'Opening Carried Forward Balance Incl. GST',
                'Total Banked Amount': 'Amount Banked',
                'Closing Carried Forward Balance': 'Closing Carried Forward Balance Incl. GST'
            }
            df = self.replace_keys(replaces, df)

            field_id = 'Broker ID'
            for index, row in df.iterrows():
                drow = df.loc[df[field_id] == row[field_id]].to_dict(orient='records')[0]
                drow['line'] = index
                rows[drow[field_id]] = drow
        except xlrd.biffh.XLRDError:
            print(f"{bcolors.YELLOW}No sheet named {tab} found in {bcolors.BLUE}{self.full_path}{bcolors.ENDC}")
        return rows

    def parse_executive_summary(self, xl, tab):
        rows = {}
        df = xl.parse(tab)
        df = df.dropna(how='all')  # remove rows that don't have any value
        df = self.general_replaces(df)
        rows_counter = {}
        rows = {}
        counter = 0
        for index, row in df.iterrows():
            execsum_row = ExecutiveSummaryRow(df.iloc[counter, 0], df.iloc[counter, 1], index)
            self.__add_datarow(rows, rows_counter, execsum_row)
            counter += 1
        return rows

    def parse_de(self, xl, tab):
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.dropna(how='all')
            df = self.general_replaces(df)
            if df.columns[0] != 'Num#':
                df = df.rename(columns=df.iloc[0]).drop(df.index[0])
                df = df.drop(['Aggregator ABN No#'], axis=1)
                df = df.drop(['ABN'], axis=1)
                df = df.drop(['GST Registered'], axis=1)
                df = df.drop(['Commission Email'], axis=1)
                df = df.drop(['Mobile No#'], axis=1)
            else:
                df = df.drop(['Num#'], axis=1)

            replaces = {
                'Finsure BSB No#': 'Aggregator BSB No#',
                'Finsure Acct No#': 'Aggregator Acct No#'
            }
            df = self.replace_keys(replaces, df)

            rows_counter = {}
            rows = {}
            for index, row in df.iterrows():
                lsum_row = DEExecutiveSummaryRow(
                    row['Aggregator'], row['Aggregator BSB No#'], row['Aggregator Acct No#'], row['Branch ID'],
                    row['Agent Type'], row['Company Name'], row['Bank Account Name'], row['BSB No#'],
                    row['Account No#'], row['Amount Banked'], index)
                self.__add_datarow(rows, rows_counter, lsum_row)
        except xlrd.biffh.XLRDError:
            print(f"{bcolors.YELLOW}No sheet named {tab} found in {bcolors.BLUE}{self.full_path}{bcolors.ENDC}")
        return rows

    def replace_keys(self, replaces: dict, df):
        for key in replaces.keys():
            if key in df.columns:
                df[replaces[key]] = df[key]
                del df[key]
        return df

    def process_comparison(self, margin=0.000001):
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        if self.pair is None:
            return None

        workbook = self.create_workbook(OUTPUT_DIR_EXEC_SUMMARY)

        self.process_specific(
            workbook, 'Executive Summary Report',
            self.datarows_executive_summary, self.pair.datarows_executive_summary,
            ExecutiveSummaryRow, HEADER_EXECUTIVE_SUMMARY)

        self.process_specific(
            workbook, 'Fee Summary Report',
            self.datarows_fee, self.pair.datarows_fee,
            ExecutiveSummaryRow, HEADER_EXECUTIVE_SUMMARY)

        self.process_specific(
            workbook, 'Lender Upfront Records',
            self.datarows_lender_upfront, self.pair.datarows_lender_upfront,
            LenderExecutiveSummaryRow, HEADER_LENDER)

        self.process_specific(
            workbook, 'Lender Trail Records',
            self.datarows_lender_trail, self.pair.datarows_lender_trail,
            LenderExecutiveSummaryRow, HEADER_LENDER)

        self.process_specific(
            workbook, 'Lender VBI Records',
            self.datarows_lender_vbi, self.pair.datarows_lender_vbi,
            LenderExecutiveSummaryRow, HEADER_LENDER)

        self.process_generic(
            workbook, 'Branch Summary Report',
            self.datarows_branch_summary, self.pair.datarows_branch_summary)

        self.process_generic(
            workbook, 'Branch Fee Summary Report',
            self.datarows_branch_fee_summary, self.pair.datarows_branch_summary)

        self.process_generic(
            workbook, 'Broker Summary Report',
            self.datarows_broker_summary, self.pair.datarows_broker_summary)

        self.process_generic(
            workbook, 'Broker Fee Summary Report',
            self.datarows_broker_fee_summary, self.pair.datarows_broker_summary)

        self.process_specific(
            workbook, 'Referrer Summary Report',
            self.datarows_referrer, self.pair.datarows_referrer,
            ReferrerExecutiveSummaryRow, HEADER_REFERRER)

        self.process_specific(
            workbook, 'DE File Entries',
            self.datarows_de_file_entries, self.pair.datarows_de_file_entries,
            DEExecutiveSummaryRow, HEADER_DE)

        self.process_specific(
            workbook, 'DE File - Amount Not Paid',
            self.datarows_de_file_notpaid, self.pair.datarows_de_file_notpaid,
            DEExecutiveSummaryRow, HEADER_DE)

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def process_generic(self, workbook, tab, dict_a, dict_b):
        worksheet = workbook.add_worksheet(tab)
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

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

    def process_specific(self, workbook, tab, datarows, datarows_pair, cls, header):
        worksheet = workbook.add_worksheet(tab)
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        row = 0
        col_a = 0
        col_b = len(header) + 1

        for index, item in enumerate(header):
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
                self.summary_errors += cls.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += cls.write_row(worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += cls.write_row(
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


class ExecutiveSummaryRow(InvoiceRow):

    def __init__(self, description, value, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

        self.description = description
        self.value = value

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
    def equal_description(self):
        if self.pair is None:
            return False
        return u.sanitize(self.description) == u.sanitize(self.pair.description)

    @property
    def equal_value(self):
        if self.pair is None:
            return False

        # Fix for header that is being compared in the Fee Tab
        if self.value == 'Amounts':
            return True

        return self.compare_numbers(self.value, self.pair.value, self.margin)
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.description).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(self.description.encode(ENCODING))
        sha.update(str(self.value).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != ExecutiveSummaryRow:
            return False

        return u.sanitize(self.description) == u.sanitize(obj.description) and self.compare_numbers(self.value, obj.value, self.margin)

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = len(HEADER_EXECUTIVE_SUMMARY) + 1

        worksheet.write(row, col, element.description)

        format_ = fmt_error if not element.equal_value else None
        worksheet.write(row, col + 1, element.value, format_)

        errors = []
        line_a = element.row_number
        description = f"{element.description}"
        if element.pair is not None:
            line_b = element.pair.row_number
            if write_errors:
                if not element.equal_value:
                    msg = 'Value does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.value, element.pair.value))
        else:
            if write_errors:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line_a, '', value_a=description))
            else:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', '', line_a, value_b=description))

        return errors


class ReferrerExecutiveSummaryRow(InvoiceRow):

    def __init__(self, branch_id, branch_name, referrer_name, opening_balance, commission_paid, total_amount_banked,
                 closing_balance, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

        self.branch_id = branch_id
        self.branch_name = branch_name
        self.referrer_name = referrer_name
        self.opening_balance = opening_balance
        self.commission_paid = commission_paid
        self.total_amount_banked = total_amount_banked
        self.closing_balance = closing_balance

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
    def equal_branch_id(self):
        if self.pair is None:
            return False
        return u.sanitize(self.branch_id) == u.sanitize(self.pair.branch_id)

    @property
    def equal_branch_name(self):
        if self.pair is None:
            return False
        return u.sanitize(self.branch_name) == u.sanitize(self.pair.branch_name)

    @property
    def equal_referrer_name(self):
        if self.pair is None:
            return False
        return u.sanitize(self.referrer_name) == u.sanitize(self.pair.referrer_name)

    @property
    def equal_opening_balance(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.opening_balance, self.pair.opening_balance, self.margin)

    @property
    def equal_commission_paid(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission_paid, self.pair.commission_paid, self.margin)

    @property
    def equal_total_amount_banked(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.total_amount_banked, self.pair.total_amount_banked, self.margin)

    @property
    def equal_closing_balance(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.closing_balance, self.pair.closing_balance, self.margin)
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.branch_id).encode(ENCODING))
        sha.update(u.sanitize(self.branch_name).encode(ENCODING))
        sha.update(u.sanitize(self.referrer_name).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.branch_id).encode(ENCODING))
        sha.update(u.sanitize(self.branch_name).encode(ENCODING))
        sha.update(u.sanitize(self.referrer_name).encode(ENCODING))
        sha.update(u.sanitize(str(self.opening_balance)).encode(ENCODING))
        sha.update(u.sanitize(str(self.commission_paid)).encode(ENCODING))
        sha.update(u.sanitize(str(self.total_amount_banked)).encode(ENCODING))
        sha.update(u.sanitize(str(self.closing_balance)).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != LenderExecutiveSummaryRow:
            return False

        return (
            u.sanitize(self.branch_id) == u.sanitize(obj.branch_id)
            and u.sanitize(self.branch_name) == u.sanitize(obj.branch_name)
            and u.sanitize(self.referrer_name) == u.sanitize(obj.referrer_name)
            and self.compare_numbers(self.opening_balance, obj.opening_balance, self.margin)
            and self.compare_numbers(self.commission_paid, obj.commission_paid, self.margin)
            and self.compare_numbers(self.total_amount_banked, obj.total_amount_banked, self.margin)
            and self.compare_numbers(self.closing_balance, obj.closing_balance, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = len(HEADER_REFERRER) + 1

        worksheet.write(row, col, element.branch_id)
        worksheet.write(row, col + 1, element.branch_name)
        worksheet.write(row, col + 2, element.referrer_name)

        format_ = fmt_error if not element.equal_opening_balance else None
        worksheet.write(row, col + 3, element.opening_balance, format_)

        format_ = fmt_error if not element.equal_commission_paid else None
        worksheet.write(row, col + 4, element.commission_paid, format_)

        format_ = fmt_error if not element.equal_total_amount_banked else None
        worksheet.write(row, col + 5, element.total_amount_banked, format_)

        format_ = fmt_error if not element.equal_closing_balance else None
        worksheet.write(row, col + 6, element.closing_balance, format_)

        errors = []
        line_a = element.row_number
        description = f"Referrer name: {element.referrer_name}"
        if element.pair is not None:
            line_b = element.pair.row_number
            if write_errors:
                if not element.equal_opening_balance:
                    msg = 'Opening Balance does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.opening_balance, element.pair.opening_balance))

                if not element.equal_commission_paid:
                    msg = 'Commission Amount Paid Incl. GST does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.commission_paid, element.pair.commission_paid))

                if not element.equal_total_amount_banked:
                    msg = 'Total Banked Amount does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.total_amount_banked, element.pair.total_amount_banked))

                if not element.equal_closing_balance:
                    msg = 'Closing Balance does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.closing_balance, element.pair.closing_balance))
        else:
            if write_errors:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line_a, '', value_a=description))
            else:
                errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', '', line_a, value_b=description))

        return errors


class DEExecutiveSummaryRow(InvoiceRow):

    def __init__(self, aggregator, aggregator_bsb_number, aggregator_acc_number, branch_id, agent_type,
                 company_name, bank_account_name, bsb, account, amount_banked, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

        self.aggregator = aggregator
        self.aggregator_bsb_number = aggregator_bsb_number
        self.aggregator_acc_number = aggregator_acc_number
        self.branch_id = branch_id
        self.agent_type = agent_type
        self.company_name = company_name
        self.bank_account_name = bank_account_name
        self.bsb = bsb
        self.account = account
        self.amount_banked = amount_banked

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
    def equal_aggregator(self):
        if self.pair is None:
            return False
        return u.sanitize(self.aggregator) == u.sanitize(self.pair.aggregator)

    @property
    def equal_aggregator_bsb_number(self):
        if self.pair is None:
            return False
        return u.sanitize(self.aggregator_bsb_number) == u.sanitize(self.pair.aggregator_bsb_number)

    @property
    def equal_aggregator_acc_number(self):
        if self.pair is None:
            return False
        return u.sanitize(self.aggregator_acc_number) == u.sanitize(self.pair.aggregator_acc_number)

    @property
    def equal_branch_id(self):
        if self.pair is None:
            return False
        return u.sanitize(self.branch_id) == u.sanitize(self.pair.branch_id)

    @property
    def equal_agent_type(self):
        if self.pair is None:
            return False
        return u.sanitize(self.agent_type) == u.sanitize(self.pair.agent_type)

    @property
    def equal_company_name(self):
        if self.pair is None:
            return False
        return u.sanitize(self.company_name) == u.sanitize(self.pair.company_name)

    @property
    def equal_bank_account_name(self):
        if self.pair is None:
            return False
        return u.sanitize(self.bank_account_name) == u.sanitize(self.pair.bank_account_name)

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

    @property
    def equal_amount_banked(self):
        if self.pair is None:
            return False
        return u.compare_numbers(self.amount_banked, self.pair.amount_banked, self.margin)
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.bsb).encode(ENCODING))
        sha.update(u.sanitize(self.account).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.aggregator).encode(ENCODING))
        sha.update(u.sanitize(self.aggregator_bsb_number).encode(ENCODING))
        sha.update(u.sanitize(self.aggregator_acc_number).encode(ENCODING))
        sha.update(u.sanitize(self.branch_id).encode(ENCODING))
        sha.update(u.sanitize(self.agent_type).encode(ENCODING))
        sha.update(u.sanitize(self.company_name).encode(ENCODING))
        sha.update(u.sanitize(self.bank_account_name).encode(ENCODING))
        sha.update(u.sanitize(self.bsb).encode(ENCODING))
        sha.update(u.sanitize(self.account).encode(ENCODING))
        sha.update(u.sanitize(str(self.amount_banked)).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != LenderExecutiveSummaryRow:
            return False

        return (
            u.sanitize(self.aggregator) == u.sanitize(obj.aggregator)
            and u.sanitize(self.aggregator_bsb_number) == u.sanitize(obj.aggregator_bsb_number)
            and u.sanitize(self.aggregator_acc_number) == u.sanitize(obj.aggregator_acc_number)
            and u.sanitize(self.branch_id) == u.sanitize(obj.branch_id)
            and u.sanitize(self.agent_type) == u.sanitize(obj.agent_type)
            and u.sanitize(self.company_name) == u.sanitize(obj.company_name)
            and u.sanitize(self.bank_account_name) == u.sanitize(obj.bank_account_name)
            and u.sanitize(self.bsb) == u.sanitize(obj.bsb)
            and u.sanitize(self.account) == u.sanitize(obj.account)
            and u.compare_numbers(self.amount_banked, obj.amount_banked, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = len(HEADER_DE) + 1

        format_ = fmt_error if not element.equal_aggregator else None
        worksheet.write(row, col, element.aggregator, format_)

        format_ = fmt_error if not element.equal_aggregator_bsb_number else None
        worksheet.write(row, col + 1, element.aggregator_bsb_number, format_)

        format_ = fmt_error if not element.equal_aggregator_acc_number else None
        worksheet.write(row, col + 2, element.aggregator_acc_number, format_)

        format_ = fmt_error if not element.equal_branch_id else None
        worksheet.write(row, col + 3, element.branch_id, format_)

        format_ = fmt_error if not element.equal_agent_type else None
        worksheet.write(row, col + 4, element.agent_type, format_)

        format_ = fmt_error if not element.equal_company_name else None
        worksheet.write(row, col + 5, element.company_name, format_)

        format_ = fmt_error if not element.equal_bank_account_name else None
        worksheet.write(row, col + 6, element.bank_account_name, format_)

        worksheet.write(row, col + 7, element.bsb)
        worksheet.write(row, col + 8, element.account)

        format_ = fmt_error if not element.equal_amount_banked else None
        worksheet.write(row, col + 9, element.amount_banked, format_)

        errors = []
        line_a = element.row_number
        description = f"Bank Account Name: {element.bank_account_name}"
        if element.pair is not None:
            line_b = element.pair.row_number
            if write_errors:

                if not element.equal_aggregator:
                    msg = 'Aggregator does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.aggregator, element.pair.aggregator))

                if not element.equal_aggregator_bsb_number:
                    msg = 'Aggregator BSB N# does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.aggregator_bsb_number, element.pair.aggregator_bsb_number))

                if not element.equal_aggregator_acc_number:
                    msg = 'Agregator Acct No# does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.aggregator_acc_number, element.pair.aggregator_acc_number))

                if not element.equal_branch_id:
                    msg = 'Branch ID does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.branch_id, element.pair.branch_id))

                if not element.equal_agent_type:
                    msg = 'Agent Type does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.agent_type, element.pair.agent_type))

                if not element.equal_company_name:
                    msg = 'Company Name does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.company_name, element.pair.company_name))

                if not element.equal_bank_account_name:
                    msg = 'Bank Account Name does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.bank_account_name, element.pair.bank_account_name))

                if not element.equal_amount_banked:
                    msg = 'Amount Banked does not match'
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, msg, line_a, line_b, element.amount_banked, element.pair.amount_banked))

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
