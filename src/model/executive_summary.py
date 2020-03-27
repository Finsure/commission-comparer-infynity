
import numpy
import pandas

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
        self.summary_errors = []
        self.pair = None
        self.margin = 0
        self.parse()

    def parse(self):
        xl = pandas.ExcelFile(self.full_path)
        self.parse_branch_summary_report(xl)

    def parse_branch_summary_report(self, excel_file):
        df = excel_file.parse('Branch Summary Report')
        df = df.dropna(how='all')  # remove rows that don't have any value
        df = df.replace(numpy.nan, '', regex=True)  # remove rows that don't have any value
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header
        for index, row in df.iterrows():
            drow = df.loc[df['ID'] == row['ID']].to_dict(orient='records')[0]
            drow['line'] = index
            if drow['ID'] != 'Total':
                drow['ID'] = int(drow['ID'])
            self.datarows_branch_summary[drow['ID']] = drow

    def process_comparison(self, margin=0.000001):
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        if self.pair is None:
            return None

        workbook = self.create_workbook(OUTPUT_DIR_EXEC_SUMMARY)
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        current_tab = 'Branch Summary Report'
        worksheet = workbook.add_worksheet(current_tab)

        # This return an arbitrary element from the dictionary so we can get the headers
        header = next(iter(self.datarows_branch_summary.values()))

        row = 0
        col_a = 0
        col_b = len(header.keys()) + 1

        for index, item in enumerate(header.keys()):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1


        keys_unmatched = set(self.pair.datarows_branch_summary.keys()) - set(self.datarows_branch_summary.keys())

        for key in self.datarows_branch_summary.keys():
            self_row = self.datarows_branch_summary[key]
            pair_row = self.pair.datarows_branch_summary.get(key, None)

            self.summary_errors += comapre_dicts(
                worksheet, row, self_row, pair_row, self.margin, self.filename, self.pair.filename,
                fmt_error, current_tab)

            row += 1

        workbook.close()

    def new_error(self, msg, line_a='', line_b='', value_a='', value_b='', tab=''):
        return new_error(self.filename, self.pair.filename, msg, line_a, line_b, value_a, value_b, tab)


def comapre_dicts(worksheet, row, row_a, row_b, margin, filename_a, filename_b, fmt_error, tab):
    errors = []
    col_a = 0
    col_b = len(row_a.keys()) + 1

    if row_b is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', row_a['line'], '', tab=tab))
    elif row_a is None:
        errors.append(new_error(filename_a, filename_b, 'No corresponding row in commission file', '', row_b['line'], tab=tab))

    for index, column in enumerate(row_a.keys()):

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
    if (type(val_a) == float) and (type(val_b) == float):
        return u.compare_numbers(val_a, val_b, margin)
    else:
        return u.sanitize(val_a) == u.sanitize(val_b)


def read_file_exec_summary(file: str):
    print(f'Parsing executive summary file {bcolors.BLUE}{file}{bcolors.ENDC}')
    filename = file.split('/')[-1]
    dir_ = '/'.join(file.split('/')[:-1]) + '/'
    return ExecutiveSummary(dir_, filename)
