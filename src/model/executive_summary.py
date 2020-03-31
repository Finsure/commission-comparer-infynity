
import numpy
import pandas
import xlrd
import copy

from src.model.taxinvoice import (TaxInvoice, new_error, OUTPUT_DIR_EXEC_SUMMARY, get_header_format,
                                  get_error_format)
from src import utils as u
from src.utils import bcolors


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
        self.summary_errors = []
        self.pair = None
        self.margin = 0
        self.parse()

    def parse(self):
        xl = pandas.ExcelFile(self.full_path)
        self.datarows_branch_summary = self.parse_branch(xl, 'Branch Summary Report')
        self.datarows_branch_fee_summary = self.parse_branch(xl, 'Branch Fee Summary Report')
        self.datarows_broker_summary = self.parse_broker(xl, 'Broker Summary Report')
        self.datarows_broker_fee_summary = self.parse_broker(xl, 'Broker Fee Summary Report')

    def parse_branch(self, xl, tab):
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.dropna(how='all')  # remove rows that don't have any value
            df = self.general_replaces(df)
            df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header
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
        field_id = 'Broker Name (ID)'
        rows = {}
        try:
            df = xl.parse(tab)
            df = df.dropna(how='all')  # remove rows that don't have any value
            df = self.general_replaces(df)
            if tab in ['Broker Summary Report']:
                df = df.rename(columns=df.iloc[1]).drop(df.index[0]).drop(df.index[1])  # Make first row the table header
            else:
                df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header

            if 'Broker ID' in list(df):
                df['Broker Name (ID)'] = df['Broker Name'] + ' (' + df['Broker ID'] + ')'
                del df['Broker ID']
                del df['Broker Name']

            for index, row in df.iterrows():
                drow = df.loc[df[field_id] == row[field_id]].to_dict(orient='records')[0]
                drow['line'] = index
                if drow[field_id] != 'Total':
                    drow[field_id] = u.sanitize(drow[field_id])
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

        self.process_generic(
            workbook, 'Branch Summary Report',
            self.datarows_branch_summary,
            self.pair.datarows_branch_summary,
            fmt_table_header, fmt_error)

        self.process_generic(
            workbook, 'Branch Fee Summary Report',
            self.datarows_branch_fee_summary,
            self.pair.datarows_branch_summary,
            fmt_table_header, fmt_error)

        self.process_generic(
            workbook, 'Broker Summary Report',
            self.datarows_broker_summary,
            self.pair.datarows_broker_summary,
            fmt_table_header, fmt_error)

        self.process_generic(
            workbook, 'Broker Fee Summary Report',
            self.datarows_broker_fee_summary,
            self.pair.datarows_broker_summary,
            fmt_table_header, fmt_error)

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

    def new_error(self, msg, line_a='', line_b='', value_a='', value_b='', tab=''):
        return new_error(self.filename, self.pair.filename, msg, line_a, line_b, value_a, value_b, tab)

    def general_replaces(self, df):
        df = df.replace(numpy.nan, '', regex=True)  # remove rows that don't have any value
        df = df.replace('Incl.', 'Inc', regex=True)
        df = df.replace('Excl.', 'Exc', regex=True)
        # df = df.replace('Payment', 'Pmt', regex=True)
        # df = df.replace('Amount', 'Amt', regex=True)

        return df


def comapre_dicts(worksheet, row, row_a, row_b, margin, filename_a, filename_b, fmt_error, tab):
    errors = []
    if row_b is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', row_a['line'], '', tab=tab))
        return errors
    elif row_a is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', '', row_b['line'], tab=tab))
        return errors

    col_a = 0
    col_b = len(row_a.keys()) + 1

    for index, column in enumerate(row_a.keys()):
        if column == 'line':
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
