import os
import numpy
import hashlib

import pandas
import xlsxwriter
from xlrd.biffh import XLRDError

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_BRANCH, new_error,
                                  get_header_format, get_error_format)

from src import utils as u

HEADER_VBI = ['Broker', 'Lender', 'Client', 'Ref #', 'Settled Loan',
              'Settlement Date', 'Commission', 'GST', 'Fee/Commission Split',
              'Fees GST', 'Remitted/Net', 'Paid To Broker', 'Paid To Referrer', 'Retained']

HEADER_UPFRONT = HEADER_VBI

HEADER_TRAIL = ['Broker', 'Lender', 'Client', 'Ref #', 'Loan Balance',
                'Settlement Date', 'Commission', 'GST', 'Fee/Commission Split',
                'Fees GST', 'Remitted/Net', 'Paid To Broker', 'Paid To Referrer', 'Retained']

HEADER_TAXINVOICE = ['Description', 'Amount', 'Gst', 'Total', 'Comments']

HEADER_RCTI = ['Description', 'Amount', 'Gst', 'Total']

HEADER_SUMMARY = HEADER_RCTI

HEADER_SUMMARY_SHORTENED = ['Description', 'Amount']

TAB_SUMMARY = 'Summary'
TAB_RCTI = 'RCTI'
TAB_TAX_INVOICE = 'Tax Invoice'
TAB_UPFRONT_DATA = 'Upfront Data'
TAB_TRAIL_DATA = 'Trail Data'
TAB_VBI_DATA = 'Vbi Data'


class BranchTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.pair = None

        # VBI Data tab fields
        self.vbi_data_rows = {}
        self.vbi_data_rows_count = {}

        # Trail Data tab fields
        self.trail_data_rows = {}
        self.trail_data_rows_count = {}

        # Upfront Data tab fields
        self.upfront_data_rows = {}
        self.upfront_data_rows_count = {}

        # Tax Invoice tab fields
        self.tax_invoice_data_rows_a = {}
        self.tax_invoice_data_rows_a_count = {}
        self.tax_invoice_data_rows_b = {}
        self.tax_invoice_data_rows_b_count = {}
        self.tax_invoice_from = ''
        self.tax_invoice_from_abn = ''
        self.tax_invoice_to = ''
        self.tax_invoice_to_abn = ''

        # RCTI tab fields
        self.rcti_data_rows = {}
        self.rcti_data_rows_count = {}
        self.rcti_from = ''
        self.rcti_from_abn = ''
        self.rcti_to = ''
        self.rcti_to_abn = ''

        # Summary_tab_fields
        self.summary_summary = {}
        self.summary_summary_count = {}
        self.summary_ptbff = {}  # payment to brokers from finsure
        self.summary_ptbff_count = {}
        self.summary_mobbtb = {}  # money owned by brokers to branch
        self.summary_mobbtb_count = {}
        self.summary_ptrff = {}  # payment to referrers from finsure
        self.summary_ptrff_count = {}
        self.summary_mobrtb = {}  # money owned by referrers to branch
        self.summary_mobrtb_count = {}
        self.summary_from = ''
        self.summary_to = ''

        self.summary_errors = []
        self._key = self._generate_key()
        self.parse()

    def parse(self):
        self.parse_tab_vbi_data(TAB_VBI_DATA)
        self.parse_tab_trail_data()
        self.parse_tab_vbi_data(TAB_UPFRONT_DATA)  # VBI and Upfront have the same strurcture, therefore we can reuse.
        self.parse_tab_tax_invoice()
        self.parse_tab_rcti()
        self.parse_tab_summary()

    def parse_tab_vbi_data(self, tab):
        try:
            df = pandas.read_excel(self.full_path, sheet_name=tab)
            df = df.dropna(how='all')
            df = df.replace(numpy.nan, '', regex=True)
            df = df.replace('--', ' ', regex=True)
            if df.columns[0] != 'Broker':
                df = df.rename(columns=df.iloc[0]).drop(df.index[0])

            for index, row in df.iterrows():
                vbidatarow = VBIDataRow(
                    row['Broker'],
                    row['Lender'],
                    row['Client'],
                    row['Ref #'],
                    # row['Referrer'],
                    float(row['Settled Loan']),
                    row['Settlement Date'],
                    float(row['Commission']),
                    float(row['GST']),
                    float(row['Fee/Commission Split']),
                    float(row['Fees GST']),
                    float(row['Remitted/Net']),
                    float(row['Paid To Broker']),
                    float(row['Paid To Referrer']),
                    float(row['Retained']),
                    index)
                if tab == TAB_VBI_DATA:
                    self.__add_datarow(self.vbi_data_rows, self.vbi_data_rows_count, vbidatarow)
                elif tab == TAB_UPFRONT_DATA:
                    self.__add_datarow(self.upfront_data_rows, self.upfront_data_rows_count, vbidatarow)
        except XLRDError:
            pass

    def parse_tab_trail_data(self):
        df = pandas.read_excel(self.full_path, sheet_name=TAB_TRAIL_DATA)
        df = df.dropna(how='all')
        df = df.replace(numpy.nan, '', regex=True)
        df = df.replace('--', ' ', regex=True)
        if df.columns[0] != 'Broker':
            df = df.rename(columns=df.iloc[0]).drop(df.index[0])

        for index, row in df.iterrows():
            traildatarow = TrailDataRow(
                row['Broker'],
                row['Lender'],
                row['Client'],
                row['Ref #'],
                # row['Referrer'],
                float(row['Loan Balance']),
                row['Settlement Date'],
                float(row['Commission']),
                float(row['GST']),
                float(row['Fee/Commission Split']),
                float(row['Fees GST']),
                float(row['Remitted/Net']),
                float(row['Paid To Broker']),
                float(row['Paid To Referrer']),
                float(row['Retained']),
                index)
            self.__add_datarow(self.trail_data_rows, self.trail_data_rows_count, traildatarow)

    def parse_tab_tax_invoice(self):
        df = pandas.read_excel(self.full_path, sheet_name=TAB_TAX_INVOICE)
        if df.iloc[1]['Tax Invoice Summary'] == 'Date:':
            df = df.drop(index=1)

        df = df.replace(' ', numpy.nan, regex=False)
        df = df.dropna(how='all')
        df = df.replace(numpy.nan, '', regex=True)

        self.tax_invoice_from = df.iloc[0][1].strip()
        self.tax_invoice_from_abn = df.iloc[1][1].strip()
        self.tax_invoice_to = df.iloc[2][1].strip()
        self.tax_invoice_to_abn = df.iloc[3][1].strip()

        section1_end = None
        section2_start = None
        section2_end = None

        current_section = 1
        index = 0
        for i, row in df.iterrows():
            if row[0].lower().endswith('software fee breakdown'):
                current_section = 2
                section2_start = index + 1
            elif row[0].lower() == 'total':
                if current_section == 1:
                    section1_end = index + 1
                elif current_section == 2:
                    section2_end = index + 1
            index += 1

        df_a = df[5:section1_end]
        df_b = df[section2_start:section2_end]

        for index, row in df_a.iterrows():
            invoicerow = TaxInvoiceDataRow(row[0], row[1], row[2], row[3], row[4], index)
            self.__add_datarow(self.tax_invoice_data_rows_a, self.tax_invoice_data_rows_a_count, invoicerow)

        if section2_start is not None:
            for index, row in df_b.iterrows():
                invoicerow = TaxInvoiceDataRow(' '.join(row[0].split()), row[1], row[2], row[3], row[4], index)
                self.__add_datarow(self.tax_invoice_data_rows_b, self.tax_invoice_data_rows_b_count, invoicerow)

    def parse_tab_rcti(self):
        df = pandas.read_excel(self.full_path, sheet_name=TAB_RCTI)
        df = df.replace(' ', numpy.nan, regex=False)
        df = df.dropna(how='all')
        df = df.replace(numpy.nan, '', regex=True)

        self.rcti_from = str(df.iloc[1][1]).strip()
        self.rcti_from_abn = str(df.iloc[2][1]).strip()
        self.rcti_to = str(df.iloc[3][1]).strip()
        self.rcti_to_abn = str(df.iloc[4][1]).strip()

        df = df[7:len(df)]

        for index, row in df.iterrows():
            rctirow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
            self.__add_datarow(self.rcti_data_rows, self.rcti_data_rows_count, rctirow)

    def parse_tab_summary(self):
        df = pandas.read_excel(self.full_path, sheet_name=TAB_SUMMARY)
        df = df.replace('  ', '', regex=False)
        df = df.replace(' ', numpy.nan, regex=False)
        df = df.dropna(how='all')
        df = df.replace(numpy.nan, '', regex=True)

        if df.iloc[0][0].strip() == 'Date:':
            df = df.drop(index=1)

        self.summary_from = df.iloc[1][1].strip()
        self.summary_to = df.iloc[2][1].strip()

        # Firstly we need to find out what are each section's start and end indexes
        df1_start = None
        df1_end = None
        df2_start = None
        df2_end = None
        df3_start = None
        df3_end = None
        df4_start = None
        df4_end = None
        df5_start = None
        df5_end = None

        current_df = 0
        index = 0
        for i, row in df.iterrows():
            if row[0].lower() == 'carried forward balance':
                current_df = 1
                df1_start = index
            elif row[0].lower().startswith('payment to brokers from'):
                current_df = 2
                df2_start = index + 1
            elif row[0].lower().startswith('payment to referrers from'):
                current_df = 3
                df3_start = index + 1
            elif row[0].lower() == 'money owed by brokers to branch':
                current_df = 4
                df4_start = index + 1
            elif row[0].lower() == 'money owed by referrers to branch':
                current_df = 5
                df5_start = index + 1

            elif row[0].lower() == '# of admin ids':
                df1_end = index + 1

            elif row[0].lower() == 'total':
                if current_df == 2:
                    df2_end = index + 1
                elif current_df == 3:
                    df3_end = index + 1
                elif current_df == 4:
                    df4_end = index + 1
                elif current_df == 5:
                    df5_end = index + 1

            index += 1

        # Now we have each df's dataframe
        df1 = df[df1_start:df1_end]
        de2 = df[df2_start:df2_end]
        df3 = df[df3_start:df3_end]
        df4 = df[df4_start:df4_end]
        df5 = df[df5_start:df5_end]

        # Iterate through each section and create the rows.
        # In this case we can use the RCTIDataRow bc the data matches it HURRAY!!!
        for index, row in df1.iterrows():
            summaryrow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
            self.__add_datarow(self.summary_summary, self.summary_summary_count, summaryrow)

        if df2_start is not None:
            for index, row in de2.iterrows():
                summaryrow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
                self.__add_datarow(self.summary_ptbff, self.summary_ptbff_count, summaryrow)

        if df3_start is not None:
            for index, row in df3.iterrows():
                summaryrow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
                self.__add_datarow(self.summary_mobbtb, self.summary_mobbtb_count, summaryrow)

        if df4_start is not None:
            for index, row in df4.iterrows():
                summaryrow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
                self.__add_datarow(self.summary_ptrff, self.summary_ptrff_count, summaryrow)

        if df5_start is not None:
            for index, row in df5.iterrows():
                summaryrow = RCTIDataRow(row[0], row[1], row[2], row[3], index)
                self.__add_datarow(self.summary_mobrtb, self.summary_mobrtb_count, summaryrow)

    # OH GOD WHY?
    def process_comparison(self, margin=0.000001):
        """
            Runs the comparison of the file with its own pair.
            If the comaprison is successfull it creates a DETAILED file and returns the
            Summary information.
        """
        if self.pair is None:
            return None
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        workbook = self.create_workbook()
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        # region Summary Section
        worksheet_summary = workbook.add_worksheet(TAB_SUMMARY)
        row = 0
        col_a = 0
        col_b = 5

        format_ = fmt_error if not self.equal_summary_from else None
        worksheet_summary.write(row, col_a, 'From')
        worksheet_summary.write(row, col_a + 1, self.summary_from, format_)
        row += 1
        format_ = fmt_error if not self.equal_summary_to else None
        worksheet_summary.write(row, col_a, 'To')
        worksheet_summary.write(row, col_a + 1, self.summary_to, format_)
        row += 1

        if self.pair is not None:
            row = 0
            format_ = fmt_error if not self.equal_summary_from else None
            worksheet_summary.write(row, col_b, 'From')
            worksheet_summary.write(row, col_b + 1, self.pair.summary_from, format_)
            row += 1
            format_ = fmt_error if not self.equal_summary_to else None
            worksheet_summary.write(row, col_b, 'To')
            worksheet_summary.write(row, col_b + 1, self.pair.summary_to, format_)
            row += 1

            if not self.equal_summary_from:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'From does not match', '', self.summary_from, self.pair.summary_from, tab=TAB_SUMMARY))
            if not self.equal_summary_to:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'To does not match', '', self.summary_to, self.pair.summary_to, tab=TAB_SUMMARY))
            row += 1

        sections = [self.summary_summary, self.summary_ptbff, self.summary_mobbtb, self.summary_ptrff, self.summary_mobrtb]
        sections_pairs = [self.pair.summary_summary, self.pair.summary_ptbff, self.pair.summary_mobbtb, self.pair.summary_ptrff, self.pair.summary_mobrtb]
        use_header = HEADER_SUMMARY
        for sec_index, section in enumerate(sections):
            for index, item in enumerate(use_header):
                worksheet_summary.write(row, col_a + index, item, fmt_table_header)
                worksheet_summary.write(row, col_b + index, item, fmt_table_header)
            row += 1

            keys_unmatched = set(sections_pairs[sec_index].keys() - set(section.keys()))

            ignore_last_two = sec_index > 0

            for key in section.keys():
                self_row = section[key]
                pair_row = self.find_pair_row(self_row, sections_pairs[sec_index])

                self_row.margin = margin
                self_row.pair = pair_row

                if pair_row is not None:
                    # delete from pair list so it doesn't get matched again
                    del sections_pairs[sec_index][pair_row.key_full]
                    # Remove the key from the keys_unmatched if it is there
                    if pair_row.key_full in keys_unmatched:
                        keys_unmatched.remove(pair_row.key_full)

                    pair_row.margin = margin
                    pair_row.pair = self_row
                    self.summary_errors += RCTIDataRow.write_row(
                        worksheet_summary, self, pair_row, row, fmt_error, TAB_SUMMARY, 'right', ignore_last_two, write_errors=False)

                self.summary_errors += RCTIDataRow.write_row(
                    worksheet_summary, self, self_row, row, fmt_error, TAB_SUMMARY, ignore_last_two=ignore_last_two)
                row += 1

            for key in keys_unmatched:
                self.summary_errors += RCTIDataRow.write_row(
                    worksheet_summary, self, sections_pairs[sec_index][key], row, fmt_error, TAB_SUMMARY, 'right', ignore_last_two)
                row += 1

            use_header = HEADER_SUMMARY_SHORTENED
            row += 2
        # endregion

        # region RCTI Section
        worksheet_rcti = workbook.add_worksheet(TAB_RCTI)
        row = 0
        col_a = 0
        col_b = 5

        format_ = fmt_error if not self.equal_rcti_from else None
        worksheet_rcti.write(row, col_a, 'From')
        worksheet_rcti.write(row, col_a + 1, self.rcti_from, format_)
        row += 1
        format_ = fmt_error if not self.equal_rcti_from_abn else None
        worksheet_rcti.write(row, col_a, 'From ABN')
        worksheet_rcti.write(row, col_a + 1, self.rcti_from_abn, format_)
        row += 1
        format_ = fmt_error if not self.equal_rcti_to else None
        worksheet_rcti.write(row, col_a, 'To')
        worksheet_rcti.write(row, col_a + 1, self.rcti_to, format_)
        row += 1
        format_ = fmt_error if not self.equal_rcti_to_abn else None
        worksheet_rcti.write(row, col_a, 'To ABN')
        worksheet_rcti.write(row, col_a + 1, self.rcti_to_abn, format_)

        if self.pair is not None:
            row = 0
            format_ = fmt_error if not self.equal_rcti_from else None
            worksheet_rcti.write(row, col_b, 'From')
            worksheet_rcti.write(row, col_b + 1, self.pair.rcti_from, format_)
            row += 1
            format_ = fmt_error if not self.equal_rcti_from_abn else None
            worksheet_rcti.write(row, col_b, 'From ABN')
            worksheet_rcti.write(row, col_b + 1, self.pair.rcti_from_abn, format_)
            row += 1
            format_ = fmt_error if not self.equal_rcti_to else None
            worksheet_rcti.write(row, col_b, 'To')
            worksheet_rcti.write(row, col_b + 1, self.pair.rcti_to, format_)
            row += 1
            format_ = fmt_error if not self.equal_rcti_to_abn else None
            worksheet_rcti.write(row, col_b, 'To ABN')
            worksheet_rcti.write(row, col_b + 1, self.pair.rcti_to_abn, format_)

            if not self.equal_rcti_from:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'From does not match', '', self.rcti_from, self.pair.rcti_from, tab=TAB_RCTI))
            if not self.equal_rcti_from_abn:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'From ABN does not match', '', self.rcti_from_abn, self.pair.rcti_from_abn, tab=TAB_RCTI))
            if not self.equal_rcti_to:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'To does not match', '', self.rcti_to, self.pair.rcti_to, tab=TAB_RCTI))
            if not self.equal_rcti_to_abn:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'To ABN does not match', '', self.rcti_to_abn, self.pair.rcti_to_abn, tab=TAB_RCTI))

        row += 2

        for index, item in enumerate(HEADER_RCTI):
            worksheet_rcti.write(row, col_a + index, item, fmt_table_header)
            worksheet_rcti.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.rcti_data_rows.keys() - set(self.rcti_data_rows.keys()))

        for key in self.rcti_data_rows.keys():
            self_row = self.rcti_data_rows[key]
            pair_row = self.find_pair_row(self_row, self.pair.rcti_data_rows)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.rcti_data_rows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += RCTIDataRow.write_row(
                    worksheet_rcti, self, pair_row, row, fmt_error, TAB_RCTI, 'right', write_errors=False)

            self.summary_errors += RCTIDataRow.write_row(worksheet_rcti, self, self_row, row, fmt_error, TAB_RCTI)
            row += 1

        for key in keys_unmatched:
            self.summary_errors += RCTIDataRow.write_row(
                worksheet_rcti, self, self.pair.rcti_data_rows[key], row, fmt_error, TAB_RCTI, 'right')
            row += 1
        # endregion

        # region Tax Invoice Section
        worksheet_tax_invoice = workbook.add_worksheet(TAB_TAX_INVOICE)
        row = 0
        col_a = 0
        col_b = 6

        format_ = fmt_error if not self.equal_tax_invoice_from else None
        worksheet_tax_invoice.write(row, col_a, 'From')
        worksheet_tax_invoice.write(row, col_a + 1, self.tax_invoice_from, format_)
        row += 1
        format_ = fmt_error if not self.equal_tax_invoice_from_abn else None
        worksheet_tax_invoice.write(row, col_a, 'From ABN')
        worksheet_tax_invoice.write(row, col_a + 1, self.tax_invoice_from_abn, format_)
        row += 1
        format_ = fmt_error if not self.equal_tax_invoice_to else None
        worksheet_tax_invoice.write(row, col_a, 'To')
        worksheet_tax_invoice.write(row, col_a + 1, self.tax_invoice_to, format_)
        row += 1
        format_ = fmt_error if not self.equal_tax_invoice_to_abn else None
        worksheet_tax_invoice.write(row, col_a, 'To ABN')
        worksheet_tax_invoice.write(row, col_a + 1, self.tax_invoice_to_abn, format_)

        if self.pair is not None:
            row = 0
            format_ = fmt_error if not self.equal_tax_invoice_from else None
            worksheet_tax_invoice.write(row, col_b, 'From')
            worksheet_tax_invoice.write(row, col_b + 1, self.pair.tax_invoice_from, format_)
            row += 1
            format_ = fmt_error if not self.equal_tax_invoice_from_abn else None
            worksheet_tax_invoice.write(row, col_b, 'From ABN')
            worksheet_tax_invoice.write(row, col_b + 1, self.pair.tax_invoice_from_abn, format_)
            row += 1
            format_ = fmt_error if not self.equal_tax_invoice_to else None
            worksheet_tax_invoice.write(row, col_b, 'To')
            worksheet_tax_invoice.write(row, col_b + 1, self.pair.tax_invoice_to, format_)
            row += 1
            format_ = fmt_error if not self.equal_tax_invoice_to_abn else None
            worksheet_tax_invoice.write(row, col_b, 'To ABN')
            worksheet_tax_invoice.write(row, col_b + 1, self.pair.tax_invoice_to_abn, format_)

            if not self.equal_tax_invoice_from:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'From does not match', '', self.tax_invoice_from, self.pair.tax_invoice_from, tab=TAB_TAX_INVOICE))
            if not self.equal_tax_invoice_from_abn:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'From ABN does not match', '', self.tax_invoice_from_abn, self.pair.tax_invoice_from_abn, tab=TAB_TAX_INVOICE))
            if not self.equal_tax_invoice_to:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'To does not match', '', self.tax_invoice_to, self.pair.tax_invoice_to, tab=TAB_TAX_INVOICE))
            if not self.equal_tax_invoice_to_abn:
                self.summary_errors.append(new_error(
                    self.filename, self.pair.filename, 'To ABN does not match', '', self.tax_invoice_to_abn, self.pair.tax_invoice_to_abn, tab=TAB_TAX_INVOICE))

        row += 2

        # Part A
        for index, item in enumerate(HEADER_TAXINVOICE):
            worksheet_tax_invoice.write(row, col_a + index, item, fmt_table_header)
            worksheet_tax_invoice.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.tax_invoice_data_rows_a.keys() - set(self.tax_invoice_data_rows_a.keys()))

        for key_full in self.tax_invoice_data_rows_a.keys():
            self_row = self.tax_invoice_data_rows_a[key_full]
            pair_row = self.find_pair_row(self_row, self.pair.tax_invoice_data_rows_a)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.tax_invoice_data_rows_a[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += TaxInvoiceDataRow.write_row(
                    worksheet_tax_invoice, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += TaxInvoiceDataRow.write_row(worksheet_tax_invoice, self, self_row, row, fmt_error)
            row += 1

        for key in keys_unmatched:
            self.summary_errors += TaxInvoiceDataRow.write_row(
                worksheet_tax_invoice, self, self.pair.tax_invoice_data_rows_a[key], row, fmt_error, 'right')
            row += 1
        row += 2

        # Part B
        for index, item in enumerate(HEADER_TAXINVOICE):
            worksheet_tax_invoice.write(row, col_a + index, item, fmt_table_header)
            worksheet_tax_invoice.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.tax_invoice_data_rows_b.keys() - set(self.tax_invoice_data_rows_b.keys()))

        for key_full in self.tax_invoice_data_rows_b.keys():
            self_row = self.tax_invoice_data_rows_b[key_full]
            pair_row = self.find_pair_row(self_row, self.pair.tax_invoice_data_rows_b)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.tax_invoice_data_rows_b[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += TaxInvoiceDataRow.write_row(
                    worksheet_tax_invoice, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += TaxInvoiceDataRow.write_row(worksheet_tax_invoice, self, self_row, row, fmt_error)
            row += 1

        for key in keys_unmatched:
            self.summary_errors += TaxInvoiceDataRow.write_row(
                worksheet_tax_invoice, self, self.pair.tax_invoice_data_rows_b[key], row, fmt_error, 'right')
            row += 1
        # endregion

        # region Upfront Data Section
        worksheet_upfront = workbook.add_worksheet(TAB_UPFRONT_DATA)
        row = 0
        col_a = 0
        col_b = 16

        # Write headers to Upfront tab
        for index, item in enumerate(HEADER_UPFRONT):
            worksheet_upfront.write(row, col_a + index, item, fmt_table_header)
            worksheet_upfront.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.upfront_data_rows.keys() - set(self.upfront_data_rows.keys()))

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in self.upfront_data_rows.keys():
            self_row = self.upfront_data_rows[key_full]
            pair_row = self.find_pair_row(self_row, self.pair.upfront_data_rows)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.upfront_data_rows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += VBIDataRow.write_row(
                    worksheet_upfront, self, pair_row, row, fmt_error, TAB_UPFRONT_DATA, 'right', write_errors=False)

            self.summary_errors += VBIDataRow.write_row(worksheet_upfront, self, self_row, row, fmt_error, TAB_UPFRONT_DATA)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += VBIDataRow.write_row(
                worksheet_upfront, self, self.pair.upfront_data_rows[key], row, fmt_error, TAB_UPFRONT_DATA, 'right')
            row += 1
        # endregion

        # region Trail Data Section
        worksheet_trail = workbook.add_worksheet(TAB_TRAIL_DATA)
        row = 0
        col_a = 0
        col_b = 16

        # Write headers to Trail tab
        for index, item in enumerate(HEADER_TRAIL):
            worksheet_trail.write(row, col_a + index, item, fmt_table_header)
            worksheet_trail.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.trail_data_rows.keys() - set(self.trail_data_rows.keys()))

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in self.trail_data_rows.keys():
            self_row = self.trail_data_rows[key_full]
            pair_row = self.find_pair_row(self_row, self.pair.trail_data_rows)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.trail_data_rows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += TrailDataRow.write_row(
                    worksheet_trail, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += TrailDataRow.write_row(worksheet_trail, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += TrailDataRow.write_row(
                worksheet_trail, self, self.pair.trail_data_rows[key], row, fmt_error, 'right')
            row += 1
        # endregion

        # region Vbi Data Section
        worksheet_vbi = workbook.add_worksheet(TAB_VBI_DATA)
        row = 0
        col_a = 0
        col_b = 16

        # Write headers to VBI tab
        for index, item in enumerate(HEADER_VBI):
            worksheet_vbi.write(row, col_a + index, item, fmt_table_header)
            worksheet_vbi.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.vbi_data_rows.keys() - set(self.vbi_data_rows.keys()))

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in self.vbi_data_rows.keys():
            self_row = self.vbi_data_rows[key_full]
            pair_row = self.pair.vbi_data_rows.get(key, None)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.vbi_data_rows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += VBIDataRow.write_row(
                    worksheet_vbi, self, pair_row, row, fmt_error, TAB_VBI_DATA, 'right', write_errors=False)

            self.summary_errors += VBIDataRow.write_row(worksheet_vbi, self, self_row, row, fmt_error, TAB_VBI_DATA)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += VBIDataRow.write_row(
                worksheet_vbi, self, self.pair.vbi_data_rows[key], row, fmt_error, TAB_VBI_DATA, 'right')
            row += 1
        # endregion

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def __add_datarow(self, datarows_dict, counter_dict, row):
        if row.key_full in datarows_dict.keys():  # If the row already exists
            counter_dict[row.key_full] += 1  # Increment row count for that key_full
            row.key_full = row._generate_key(counter_dict[row.key_full])  # Generate new key_full for the record
            datarows_dict[row.key_full] = row  # Add row to the list
        else:
            counter_dict[row.key_full] = 0  # Start counter
            datarows_dict[row.key_full] = row  # Add row to the list

    # region Properties
    @property
    def equal_tax_invoice_from(self):
        if self.pair is None:
            return False
        return self.tax_invoice_from == self.pair.tax_invoice_from

    @property
    def equal_tax_invoice_from_abn(self):
        if self.pair is None:
            return False
        return self.tax_invoice_from_abn == self.pair.tax_invoice_from_abn

    @property
    def equal_tax_invoice_to(self):
        if self.pair is None:
            return False
        return self.tax_invoice_to == self.pair.tax_invoice_to

    @property
    def equal_tax_invoice_to_abn(self):
        if self.pair is None:
            return False
        return self.tax_invoice_to_abn == self.pair.tax_invoice_to_abn

    @property
    def equal_rcti_from(self):
        if self.pair is None:
            return False
        return self.rcti_from == self.pair.rcti_from

    @property
    def equal_rcti_from_abn(self):
        if self.pair is None:
            return False
        return self.rcti_from_abn == self.pair.rcti_from_abn

    @property
    def equal_rcti_to(self):
        if self.pair is None:
            return False
        return self.rcti_to == self.pair.rcti_to

    @property
    def equal_rcti_to_abn(self):
        if self.pair is None:
            return False
        return self.rcti_to_abn == self.pair.rcti_to_abn

    @property
    def equal_summary_from(self):
        if self.pair is None:
            return False
        return self.summary_from == self.pair.summary_from

    @property
    def equal_summary_to(self):
        if self.pair is None:
            return False
        return self.summary_to == self.pair.summary_to
    # endregion

    def find_pair_row(self, row, pair_datarows):
        # Match by full_key
        pair_row = pair_datarows.get(row.key_full, None)
        if pair_row is not None:
            return pair_row

        # We want to match by similarity before matching by the key
        # Match by similarity
        for _, item in pair_datarows.items():
            if row.equals(item):
                return item

        # Match by key
        for _, item in pair_datarows.items():
            if row.key == item.key:
                return item

        # Return None if nothing found
        return None

    def create_workbook(self):
        suffix = '' if self.filename.endswith('.xlsx') else '.xlsx'
        return xlsxwriter.Workbook(OUTPUT_DIR_BRANCH + 'DETAILED_' + self.filename + suffix)

    def _generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[0:5]
        filename_forkey = ''.join(filename_parts)

        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()


class VBIDataRow(InvoiceRow):

    def __init__(self, broker, lender, client, ref_no, settled_loan, settlement_date,
                 commission, gst, commission_split, fees_gst, remitted, paid_to_broker,
                 paid_to_referrer, retained, document_row=None):
        InvoiceRow.__init__(self)

        self.broker = broker.strip()
        self.lender = lender.strip()
        self.client = client.strip()
        self.ref_no = str(ref_no).strip().split('.')[0]
        self.referrer = ''
        self.settled_loan = settled_loan
        self.settlement_date = settlement_date
        self.commission = commission
        self.gst = gst
        self.commission_split = commission_split
        self.fees_gst = fees_gst
        self.remitted = remitted
        self.paid_to_broker = paid_to_broker
        self.paid_to_referrer = paid_to_referrer
        self.retained = retained

        self._pair = None
        self._margin = 0
        self._document_row = document_row + 2 if document_row is not None else document_row

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
    def document_row(self):
        return self._document_row

    @property
    def equal_referrer(self):
        if self.pair is None:
            return False
        return u.sanitize(self.referrer) == u.sanitize(self.pair.referrer)

    @property
    def equal_settled_loan(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.settled_loan, self.pair.settled_loan, self.margin)

    @property
    def equal_settlement_date(self):
        if self.pair is None:
            return False
        return u.sanitize(self.settlement_date) == u.sanitize(self.pair.settlement_date)

    @property
    def equal_commission(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission, self.pair.commission, self.margin)

    @property
    def equal_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst, self.pair.gst, self.margin)

    @property
    def equal_commission_split(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission_split, self.pair.commission_split, self.margin)

    @property
    def equal_fees_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.fees_gst, self.pair.fees_gst, self.margin)

    @property
    def equal_remitted(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.remitted, self.pair.remitted, self.margin)

    @property
    def equal_paid_to_broker(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.paid_to_broker, self.pair.paid_to_broker, self.margin)

    @property
    def equal_paid_to_referrer(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.paid_to_referrer, self.pair.paid_to_referrer, self.margin)

    @property
    def equal_retained(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.retained, self.pair.retained, self.margin)

    @property
    def equal_all(self):
        return (
            self.equal_referrer()
            and self.equal_settled_loan()
            and self.equal_settlement_date()
            and self.equal_commission()
            and self.equal_gst()
            and self.equal_commission_split()
            and self.equal_fees_gst()
            and self.equal_remitted()
            and self.equal_paid_to_broker()
            and self.equal_paid_to_referrer()
            and self.equal_retained()
        )
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.broker).encode(ENCODING))
        sha.update(u.sanitize(self.lender).encode(ENCODING))
        sha.update(u.sanitize(self.client).encode(ENCODING))
        sha.update(self.ref_no.lower().encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))

        return sha.hexdigest()

    def _generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(str(self.broker).encode(ENCODING))
        sha.update(str(self.lender).encode(ENCODING))
        sha.update(str(self.client).encode(ENCODING))
        sha.update(str(self.ref_no).encode(ENCODING))
        # sha.update(str(self.referrer).encode(ENCODING))
        sha.update(str(self.settled_loan).encode(ENCODING))
        sha.update(str(self.settlement_date).encode(ENCODING))
        sha.update(str(self.commission).encode(ENCODING))
        sha.update(str(self.gst).encode(ENCODING))
        sha.update(str(self.commission_split).encode(ENCODING))
        sha.update(str(self.fees_gst).encode(ENCODING))
        sha.update(str(self.remitted).encode(ENCODING))
        sha.update(str(self.paid_to_broker).encode(ENCODING))
        sha.update(str(self.paid_to_referrer).encode(ENCODING))
        sha.update(str(self.retained).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != VBIDataRow:
            return False

        return (
            u.sanitize(self.broker) == u.sanitize(obj.broker)
            and u.sanitize(self.lender) == u.sanitize(obj.lender)
            and u.sanitize(self.client) == u.sanitize(obj.client)
            and u.sanitize(self.ref_no) == u.sanitize(obj.ref_no)
            and self.settlement_date == obj.settlement_date
            and self.compare_numbers(self.settled_loan, obj.settled_loan, self.margin)
            and self.compare_numbers(self.commission, obj.commission, self.margin)
            and self.compare_numbers(self.gst, obj.gst, self.margin)
            and self.compare_numbers(self.commission_split, obj.commission_split, self.margin)
            and self.compare_numbers(self.fees_gst, obj.fees_gst, self.margin)
            and self.compare_numbers(self.remitted, obj.remitted, self.margin)
            and self.compare_numbers(self.paid_to_broker, obj.paid_to_broker, self.margin)
            and self.compare_numbers(self.paid_to_referrer, obj.paid_to_referrer, self.margin)
            and self.compare_numbers(self.retained, obj.retained, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, tabname, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = 16

        worksheet.write(row, col, element.broker)
        worksheet.write(row, col + 1, element.lender)
        worksheet.write(row, col + 2, element.client)
        worksheet.write(row, col + 3, element.ref_no)
        # format_ = fmt_error if not element.equal_referrer else None
        # worksheet.write(row, col + 4, element.referrer, format_)
        format_ = fmt_error if not element.equal_settled_loan else None
        worksheet.write(row, col + 4, element.settled_loan, format_)
        format_ = fmt_error if not element.equal_settlement_date else None
        worksheet.write(row, col + 5, element.settlement_date, format_)
        format_ = fmt_error if not element.equal_commission else None
        worksheet.write(row, col + 6, element.commission, format_)
        format_ = fmt_error if not element.equal_gst else None
        worksheet.write(row, col + 7, element.gst, format_)
        format_ = fmt_error if not element.equal_commission_split else None
        worksheet.write(row, col + 8, element.commission_split, format_)
        format_ = fmt_error if not element.equal_fees_gst else None
        worksheet.write(row, col + 9, element.fees_gst, format_)
        format_ = fmt_error if not element.equal_remitted else None
        worksheet.write(row, col + 10, element.remitted, format_)
        format_ = fmt_error if not element.equal_paid_to_broker else None
        worksheet.write(row, col + 11, element.paid_to_broker, format_)
        format_ = fmt_error if not element.equal_paid_to_referrer else None
        worksheet.write(row, col + 12, element.paid_to_referrer, format_)
        format_ = fmt_error if not element.equal_retained else None
        worksheet.write(row, col + 13, element.retained, format_)

        errors = []
        line = element.document_row
        if element.pair is not None:
            if write_errors:
                if not element.equal_settled_loan:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Settled Loan does not match', line, element.settled_loan, element.pair.settled_loan, tab=tabname))
                if not element.equal_settlement_date:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Settlement Date does not match', line, element.settlement_date, element.pair.settlement_date, tab=tabname))
                if not element.equal_commission:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Commission does not match', line, element.commission, element.pair.commission, tab=tabname))
                if not element.equal_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'GST does not match', line, element.gst, element.pair.gst, tab=tabname))
                if not element.equal_commission_split:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Commission Split does not match', line, element.commission_split, element.pair.commission_split, tab=tabname))
                if not element.equal_fees_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Fees GST does not match', line, element.fees_gst, element.pair.fees_gst, tab=tabname))
                if not element.equal_remitted:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Remitted does not match', line, element.remitted, element.pair.remitted, tab=tabname))
                if not element.equal_paid_to_broker:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Paid to Broker does not match', line, element.paid_to_broker, element.pair.paid_to_broker, tab=tabname))
                if not element.equal_paid_to_referrer:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Paid to Referrer does not match', line, element.paid_to_referrer, element.pair.paid_to_referrer, tab=tabname))
                if not element.equal_retained:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Retained does not match', line, element.retained, element.pair.retained, tab=tabname))
        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line, tab=tabname))

        return errors


class TrailDataRow(InvoiceRow):

    def __init__(self, broker, lender, client, ref_no, loan_balance, settlement_date,
                 commission, gst, commission_split, fees_gst, remitted, paid_to_broker,
                 paid_to_referrer, retained, document_row=None):
        InvoiceRow.__init__(self)

        self.broker = broker.strip()
        self.lender = lender.strip()
        self.client = client.strip()
        self.ref_no = str(ref_no).strip()
        self.referrer = ''
        self.loan_balance = loan_balance
        self.settlement_date = settlement_date
        self.commission = commission
        self.gst = gst
        self.commission_split = commission_split
        self.fees_gst = fees_gst
        self.remitted = remitted
        self.paid_to_broker = paid_to_broker
        self.paid_to_referrer = paid_to_referrer
        self.retained = retained

        self._pair = None
        self._margin = 0
        self._document_row = document_row + 2 if document_row is not None else document_row

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
    def document_row(self):
        return self._document_row

    @property
    def equal_lender(self):
        if self.pair is None:
            return False
        return u.sanitize(self.lender) == u.sanitize(self.pair.lender)

    @property
    def equal_referrer(self):
        if self.pair is None:
            return False
        return u.sanitize(self.referrer) == u.sanitize(self.pair.referrer)

    @property
    def equal_loan_balance(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.loan_balance, self.pair.loan_balance, self.margin)

    @property
    def equal_settlement_date(self):
        if self.pair is None:
            return False
        return self.settlement_date == self.pair.settlement_date

    @property
    def equal_commission(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission, self.pair.commission, self.margin)

    @property
    def equal_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst, self.pair.gst, self.margin)

    @property
    def equal_commission_split(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.commission_split, self.pair.commission_split, self.margin)

    @property
    def equal_fees_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.fees_gst, self.pair.fees_gst, self.margin)

    @property
    def equal_remitted(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.remitted, self.pair.remitted, self.margin)

    @property
    def equal_paid_to_broker(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.paid_to_broker, self.pair.paid_to_broker, self.margin)

    @property
    def equal_paid_to_referrer(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.paid_to_referrer, self.pair.paid_to_referrer, self.margin)

    @property
    def equal_retained(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.retained, self.pair.retained, self.margin)

    @property
    def equal_all(self):
        return (
            self.equal_referrer()
            and self.equal_loan_balance()
            and self.equal_settlement_date()
            and self.equal_commission()
            and self.equal_gst()
            and self.equal_commission_split()
            and self.equal_fees_gst()
            and self.equal_remitted()
            and self.equal_paid_to_broker()
            and self.equal_paid_to_referrer()
            and self.equal_retained()
        )
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.broker).encode(ENCODING))
        sha.update(u.sanitize(self.client).encode(ENCODING))
        sha.update(self.ref_no.lower().encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(str(self.broker).encode(ENCODING))
        sha.update(str(self.lender).encode(ENCODING))
        sha.update(str(self.client).encode(ENCODING))
        sha.update(str(self.ref_no).encode(ENCODING))
        # sha.update(str(self.referrer).encode(ENCODING))
        sha.update(str(self.loan_balance).encode(ENCODING))
        sha.update(str(self.settlement_date).encode(ENCODING))
        sha.update(str(self.commission).encode(ENCODING))
        sha.update(str(self.gst).encode(ENCODING))
        sha.update(str(self.commission_split).encode(ENCODING))
        sha.update(str(self.fees_gst).encode(ENCODING))
        sha.update(str(self.remitted).encode(ENCODING))
        sha.update(str(self.paid_to_broker).encode(ENCODING))
        sha.update(str(self.paid_to_referrer).encode(ENCODING))
        sha.update(str(self.retained).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != TrailDataRow:
            return False

        return (
            u.sanitize(self.broker) == u.sanitize(obj.broker)
            and u.sanitize(self.lender) == u.sanitize(obj.lender)
            and u.sanitize(self.client) == u.sanitize(obj.client)
            and u.sanitize(self.ref_no) == u.sanitize(obj.ref_no)
            and self.compare_numbers(self.loan_balance, obj.loan_balance, self.margin)
            and self.settlement_date == obj.settlement_date
            and self.compare_numbers(self.commission, obj.commission, self.margin)
            and self.compare_numbers(self.gst, obj.gst, self.margin)
            and self.compare_numbers(self.commission_split, obj.commission_split, self.margin)
            and self.compare_numbers(self.fees_gst, obj.fees_gst, self.margin)
            and self.compare_numbers(self.remitted, obj.remitted, self.margin)
            and self.compare_numbers(self.paid_to_broker, obj.paid_to_broker, self.margin)
            and self.compare_numbers(self.paid_to_referrer, obj.paid_to_referrer, self.margin)
            and self.compare_numbers(self.retained, obj.retained, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = 16

        worksheet.write(row, col, element.broker)
        format_ = fmt_error if not element.equal_lender else None
        worksheet.write(row, col + 1, element.lender, format_)
        worksheet.write(row, col + 2, element.client)
        worksheet.write(row, col + 3, element.ref_no)
        # format_ = fmt_error if not element.equal_referrer else None
        # worksheet.write(row, col + 4, element.referrer, format_)
        format_ = fmt_error if not element.equal_loan_balance else None
        worksheet.write(row, col + 4, element.loan_balance, format_)
        format_ = fmt_error if not element.equal_settlement_date else None
        worksheet.write(row, col + 5, element.settlement_date, format_)
        format_ = fmt_error if not element.equal_commission else None
        worksheet.write(row, col + 6, element.commission, format_)
        format_ = fmt_error if not element.equal_gst else None
        worksheet.write(row, col + 7, element.gst, format_)
        format_ = fmt_error if not element.equal_commission_split else None
        worksheet.write(row, col + 8, element.commission_split, format_)
        format_ = fmt_error if not element.equal_fees_gst else None
        worksheet.write(row, col + 9, element.fees_gst, format_)
        format_ = fmt_error if not element.equal_remitted else None
        worksheet.write(row, col + 10, element.remitted, format_)
        format_ = fmt_error if not element.equal_paid_to_broker else None
        worksheet.write(row, col + 11, element.paid_to_broker, format_)
        format_ = fmt_error if not element.equal_paid_to_referrer else None
        worksheet.write(row, col + 12, element.paid_to_referrer, format_)
        format_ = fmt_error if not element.equal_retained else None
        worksheet.write(row, col + 13, element.retained, format_)

        errors = []
        line = element.document_row
        if element.pair is not None:
            if write_errors:
                if not element.equal_lender:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename,  'Lender does not match', line, element.lender, element.pair.lender, tab=TAB_TRAIL_DATA))
                if not element.equal_loan_balance:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Loan Balance does not match', line, element.loan_balance, element.pair.loan_balance, tab=TAB_TRAIL_DATA))
                if not element.equal_settlement_date:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Settlement Date does not match', line, element.settlement_date, element.pair.settlement_date, tab=TAB_TRAIL_DATA))
                if not element.equal_commission:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Commission does not match', line, element.commission, element.pair.commission, tab=TAB_TRAIL_DATA))
                if not element.equal_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'GST does not match', line, element.gst, element.pair.gst, tab=TAB_TRAIL_DATA))
                if not element.equal_commission_split:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Commission Split does not match', line, element.commission_split, element.pair.commission_split, tab=TAB_TRAIL_DATA))
                if not element.equal_fees_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Fees GST does not match', line, element.fees_gst, element.pair.fees_gst, tab=TAB_TRAIL_DATA))
                if not element.equal_remitted:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Remitted does not match', line, element.remitted, element.pair.remitted, tab=TAB_TRAIL_DATA))
                if not element.equal_paid_to_broker:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Paid to Broker does not match', line, element.paid_to_broker, element.pair.paid_to_broker, tab=TAB_TRAIL_DATA))
                if not element.equal_paid_to_referrer:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Paid to Referrer does not match', line, element.paid_to_referrer, element.pair.paid_to_referrer, tab=TAB_TRAIL_DATA))
                if not element.equal_retained:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Retained does not match', line, element.retained, element.pair.retained, tab=TAB_TRAIL_DATA))
        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line, tab=TAB_TRAIL_DATA))

        return errors


class TaxInvoiceDataRow(InvoiceRow):

    def __init__(self, description, amount, gst, total, comments, document_row=None):
        InvoiceRow.__init__(self)

        self.description = ' '.join(description.strip().split())
        self.amount = float(amount) if amount not in ('', ' ') else 0
        self.gst = float(gst) if gst != '' else 0
        self.total = float(total) if total != '' else 0
        self.comments = comments.lower()

        self._pair = None
        self._margin = 0
        self._document_row = document_row + 2 if document_row is not None else document_row

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
    def document_row(self):
        return self._document_row

    @property
    def equal_description(self):
        if self.pair is None:
            return False
        return u.sanitize(self.description) == u.sanitize(self.pair.description)

    @property
    def equal_amount(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.amount, self.pair.amount, self.margin)

    @property
    def equal_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst, self.pair.gst, self.margin)

    @property
    def equal_total(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.total, self.pair.total, self.margin)

    @property
    def equal_comments(self):
        if self.pair is None:
            return False
        return u.sanitize(self.comments) == u.sanitize(self.pair.comments)

    @property
    def equal_all(self):
        return (
            self.equal_description()
            and self.equal_amount()
            and self.equal_gst()
            and self.equal_total()
            and self.equal_comments()
        )
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.description).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(str(self.description).encode(ENCODING))
        sha.update(str(self.amount).encode(ENCODING))
        sha.update(str(self.gst).encode(ENCODING))
        sha.update(str(self.total).encode(ENCODING))
        sha.update(str(self.comments).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != TaxInvoiceDataRow:
            return False

        return (
            u.sanitize(self.description) == u.sanitize(obj.description)
            and u.sanitize(self.comments) == u.sanitize(obj.comments)
            and self.compare_numbers(self.amount, obj.amount, self.margin)
            and self.compare_numbers(self.gst, obj.gst, self.margin)
            and self.compare_numbers(self.total, obj.total, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
        col = 0
        if side == 'right':
            col = 6

        worksheet.write(row, col, element.description)
        format_ = fmt_error if not element.equal_amount else None
        worksheet.write(row, col + 1, element.amount, format_)
        format_ = fmt_error if not element.equal_gst else None
        worksheet.write(row, col + 2, element.gst, format_)
        format_ = fmt_error if not element.equal_total else None
        worksheet.write(row, col + 3, element.total, format_)
        format_ = fmt_error if not element.equal_comments else None
        worksheet.write(row, col + 4, element.comments, format_)

        errors = []
        line = element.document_row
        if element.pair is not None:
            if write_errors:
                if not element.equal_amount:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Amount does not match', line, element.amount, element.pair.amount, tab=TAB_TAX_INVOICE))
                if not element.equal_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'GST does not match', line, element.gst, element.pair.gst, tab=TAB_TAX_INVOICE))
                if not element.equal_total:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Total does not match', line, element.total, element.pair.total, tab=TAB_TAX_INVOICE))
                if not element.equal_comments:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Comments does not match', line, element.comments, element.pair.comments, tab=TAB_TAX_INVOICE))
        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line, tab=TAB_TAX_INVOICE))

        return errors


class RCTIDataRow(InvoiceRow):

    def __init__(self, description, amount, gst, total, document_row=None):
        InvoiceRow.__init__(self)

        self.description = ' '.join(description.strip().split())
        self.amount = float(amount) if amount != '' and amount != ' ' else 0
        self.gst = float(gst) if gst != '' and gst != ' ' else 0
        self.total = float(total) if total != '' and gst != ' ' else 0

        self._pair = None
        self._margin = 0
        self._document_row = document_row + 2 if document_row is not None else document_row

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
    def document_row(self):
        return self._document_row

    @property
    def equal_description(self):
        if self.pair is None:
            return False
        return u.sanitize(self.description) == u.sanitize(self.pair.description)

    @property
    def equal_amount(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.amount, self.pair.amount, self.margin)

    @property
    def equal_gst(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.gst, self.pair.gst, self.margin)

    @property
    def equal_total(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.total, self.pair.total, self.margin)

    @property
    def equal_all(self):
        return (
            self.equal_description()
            and self.equal_amount()
            and self.equal_gst()
            and self.equal_total()
        )
    # endregion

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.description).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def _generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(str(self.description).encode(ENCODING))
        sha.update(str(self.amount).encode(ENCODING))
        sha.update(str(self.gst).encode(ENCODING))
        sha.update(str(self.total).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != RCTIDataRow:
            return False

        return (
            u.sanitize(self.description) == u.sanitize(obj.description)
            and self.compare_numbers(self.amount, obj.amount, self.margin)
            and self.compare_numbers(self.gst, obj.gst, self.margin)
            and self.compare_numbers(self.total, obj.total, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, tab, side='left', ignore_last_two=False, write_errors=True):
        col = 0
        if side == 'right':
            col = 5

        worksheet.write(row, col, element.description)
        format_ = fmt_error if not element.equal_amount else None
        worksheet.write(row, col + 1, element.amount, format_)

        if not ignore_last_two:
            format_ = fmt_error if not element.equal_gst else None
            worksheet.write(row, col + 2, element.gst, format_)
            format_ = fmt_error if not element.equal_total else None
            worksheet.write(row, col + 3, element.total, format_)

        errors = []
        line = element.document_row
        if element.pair is not None:
            if write_errors:
                if not element.equal_amount:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Amount does not match', line, element.amount, element.pair.amount, tab=tab))
                if not element.equal_gst:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'GST does not match', line, element.gst, element.pair.gst, tab=tab))
                if not element.equal_total:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Total does not match', line, element.total, element.pair.total, tab=tab))
        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line, tab=tab))

        return errors


def read_files_branch(dir_: str, files: list) -> dict:
    keys = {}
    counter = 1
    for file in files:
        print(f'Parsing {counter} of {len(files)} files from {dir_}', end='\r')
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = BranchTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
        counter += 1
    print()
    return keys
