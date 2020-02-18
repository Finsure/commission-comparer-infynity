import os
import numpy
import hashlib

import pandas
import xlsxwriter

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_BRANCH_PID, new_error)


class BranchTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.pair = None
        self.vbi_data_rows = {}
        self.summary_errors = []
        self._key = self.__generate_key()
        self.parse()

    def parse(self):
        self.parse_tab_vbi_data()

    def parse_tab_vbi_data(self):
        vbi_dataframe = pandas.read_excel(self.full_path, sheet_name='Vbi Data')
        vbi_dataframe = vbi_dataframe.dropna(how='all')
        vbi_dataframe = vbi_dataframe.replace(numpy.nan, '', regex=True)
        vbi_dataframe = vbi_dataframe.replace('--', ' ', regex=True)
        if vbi_dataframe.columns[0] != 'Broker':
            vbi_dataframe = vbi_dataframe.rename(columns=vbi_dataframe.iloc[0]).drop(vbi_dataframe.index[0])

        for index, row in vbi_dataframe.iterrows():
            vbidatarow = VBIDataRow(
                row['Broker'],
                row['Lender'],
                row['Client'],
                row['Ref #'],
                row['Referrer'],
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
            self.vbi_data_rows[vbidatarow.key] = vbidatarow

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
        fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
        fmt_error = workbook.add_format({'font_color': 'red'})

        # Vbi Data Section
        worksheet_vbi = workbook.add_worksheet('Vbi Data')
        row = 0
        col_a = 0
        col_b = 16

        # Write headers to VBI tab
        header_vbi = ['Broker', 'Lender', 'Client', 'Ref #', 'Referrer', 'Settled Loan',
                      'Settlement Date', 'Commission', 'GST', 'Fee/Commission Split',
                      'Fees GST', 'Remitted/Net', 'Paid To Broker', 'Paid To Referrer', 'Retained']
        for index, item in enumerate(header_vbi):
            worksheet_vbi.write(row, col_a + index, item, fmt_table_header)
            worksheet_vbi.write(row, col_b + index, item, fmt_table_header)
        row += 1

        for key in self.vbi_data_rows.keys():
            self_row = self.vbi_data_rows[key]
            pair_row = self.pair.vbi_data_rows.get(key, None)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += VBIDataRow.write_row(
                    self.full_path, worksheet_vbi, pair_row, row, fmt_error, 'right')

            self.summary_errors += VBIDataRow.write_row(
                self.pair.full_path, worksheet_vbi, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        alone_keys_infynity = set(self.pair.vbi_data_rows.keys() - set(self.vbi_data_rows.keys()))
        for key in alone_keys_infynity:
            self.summary_errors += VBIDataRow.write_row(
                self.pair.full_path, worksheet_vbi, self.pair.vbi_data_rows[key], row, fmt_error, 'right')

        workbook.close()

        return self.summary_errors

    def create_workbook(self):
        suffix = '' if self.filename.endswith('.xlsx') else '.xlsx'
        return xlsxwriter.Workbook(OUTPUT_DIR_BRANCH_PID + 'DETAILED_' + self.filename + suffix)

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[0:5]
        filename_forkey = '_'.join(filename_parts)

        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()


class VBIDataRow(InvoiceRow):

    def __init__(self, broker, lender, client, ref_no, referrer, settled_loan, settlement_date,
                 commission, gst, commission_split, fees_gst, remitted, paid_to_broker,
                 paid_to_referrer, retained, document_row=None):
        InvoiceRow.__init__(self)

        if type(ref_no) is float:
            ref_no = int(ref_no)

        self.broker = broker
        self.lender = lender
        self.client = client
        self.ref_no = ref_no
        self.referrer = referrer
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
        self._document_row = document_row

        self._key = self.__generate_key()
        self._key_full = self.__generate_key_full()

    @property
    def key(self):
        return self._key

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
    def document_row(self):
        return self._document_row

    @property
    def equal_referrer(self):
        if self.pair is None:
            return False
        return self.referrer == self.pair.referrer

    @property
    def equal_settled_loan(self):
        if self.pair is None:
            return False
        return self.compare_numbers(self.settled_loan, self.pair.settled_loan, self.margin)

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

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(str(self.broker).encode(ENCODING))
        sha.update(str(self.lender).encode(ENCODING))
        sha.update(str(self.client).encode(ENCODING))
        sha.update(str(self.ref_no).encode(ENCODING))
        return sha.hexdigest()

    def __generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(str(self.broker).encode(ENCODING))
        sha.update(str(self.lender).encode(ENCODING))
        sha.update(str(self.client).encode(ENCODING))
        sha.update(str(self.ref_no).encode(ENCODING))
        sha.update(str(self.referrer).encode(ENCODING))
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

    @staticmethod
    def write_row(filename, worksheet, element, row, fmt_error, side='left'):
        col = 0
        if side == 'right':
            col = 16

        worksheet.write(row, col, element.broker)
        worksheet.write(row, col + 1, element.lender)
        worksheet.write(row, col + 2, element.client)
        worksheet.write(row, col + 3, element.ref_no)
        format_ = fmt_error if not element.equal_referrer else None
        worksheet.write(row, col + 4, element.referrer, format_)
        format_ = fmt_error if not element.equal_settled_loan else None
        worksheet.write(row, col + 5, element.settled_loan, format_)
        format_ = fmt_error if not element.equal_settlement_date else None
        worksheet.write(row, col + 6, element.settlement_date, format_)
        format_ = fmt_error if not element.equal_commission else None
        worksheet.write(row, col + 7, element.commission, format_)
        format_ = fmt_error if not element.equal_gst else None
        worksheet.write(row, col + 8, element.gst, format_)
        format_ = fmt_error if not element.equal_commission_split else None
        worksheet.write(row, col + 9, element.commission_split, format_)
        format_ = fmt_error if not element.equal_fees_gst else None
        worksheet.write(row, col + 10, element.fees_gst, format_)
        format_ = fmt_error if not element.equal_remitted else None
        worksheet.write(row, col + 11, element.remitted, format_)
        format_ = fmt_error if not element.equal_paid_to_broker else None
        worksheet.write(row, col + 12, element.paid_to_broker, format_)
        format_ = fmt_error if not element.equal_paid_to_referrer else None
        worksheet.write(row, col + 13, element.paid_to_referrer, format_)
        format_ = fmt_error if not element.equal_retained else None
        worksheet.write(row, col + 14, element.retained, format_)

        errors = []
        tabname = 'Vbi Data'
        line = element.document_row
        if element.pair is not None:
            if not element.equal_referrer:
                errors.append(new_error(
                    filename, 'Referrer does not match', line, element.referrer, element.pair.referrer, tab=tabname))
            if not element.equal_settled_loan:
                errors.append(new_error(
                    filename, 'Settled Loan does not match', line, element.settled_loan, element.pair.settled_loan, tab=tabname))
            if not element.equal_settlement_date:
                errors.append(new_error(
                    filename, 'Settlement Date does not match', line, element.settlement_date, element.pair.settlement_date, tab=tabname))
            if not element.equal_commission:
                errors.append(new_error(
                    filename, 'Commission does not match', line, element.commission, element.pair.commission, tab=tabname))
            if not element.equal_gst:
                errors.append(new_error(
                    filename, 'GST does not match', line, element.gst, element.pair.gst, tab=tabname))
            if not element.equal_commission_split:
                errors.append(new_error(
                    filename, 'Commission Split does not match', line, element.commission_split, element.pair.commission_split, tab=tabname))
            if not element.equal_fees_gst:
                errors.append(new_error(
                    filename, 'Fees GST does not match', line, element.fees_gst, element.pair.fees_gst, tab=tabname))
            if not element.equal_remitted:
                errors.append(new_error(
                    filename, 'Remitted does not match', line, element.remitted, element.pair.remitted, tab=tabname))
            if not element.equal_paid_to_broker:
                errors.append(new_error(
                    filename, 'Paid to Broker does not match', line, element.paid_to_broker, element.pair.paid_to_broker, tab=tabname))
            if not element.equal_paid_to_referrer:
                errors.append(new_error(
                    filename, 'Paid to Referrer does not match', line, element.paid_to_referrer, element.pair.paid_to_referrer, tab=tabname))
            if not element.equal_retained:
                errors.append(new_error(
                    filename, 'Retained does not match', line, element.retained, element.pair.retained, tab=tabname))
        else:
            errors.append(new_error(filename, 'No corresponding row in commission file', line, tab=tabname))

        return errors


def result_invoice_branch():
    return {
        'filename': '',
        'file': '',
        'has_pair': False,
        'overall': False,
        'tab_vbi_data': {
            'rows': [],
        }
    }


def read_files_branch(dir_: str, files: list) -> dict:
    keys = {}
    for file in files:
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = BranchTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return keys
