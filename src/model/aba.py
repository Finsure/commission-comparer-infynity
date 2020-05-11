from src.model.taxinvoice import TaxInvoice, new_error, OUTPUT_DIR_ABA, get_error_format
from src.utils import bcolors
import src.utils as u


class ABAFile(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.pair = None
        self.datarows = {}
        self.summary_errors = []
        self.parse()

    def parse(self):
        file = open(self.full_path, 'r')

        for index, line in enumerate(file.readlines()):
            if line.startswith('0'):
                aba_line = self.parse_line_type_0(line)
                key = u.sanitize(''.join(aba_line))
                self.datarows[key] = aba_line
            elif line.startswith('1'):
                aba_line = self.parse_line_type_1(line)
                key = u.sanitize(aba_line[7])
                self.datarows[key] = aba_line
            elif line.startswith('7'):
                aba_line = self.parse_line_type_7(line)
                key = u.sanitize(''.join(aba_line))
                self.datarows[key] = aba_line
            else:
                msg = f'There is an invalid ABA line on line {index}'
                error = new_error(self.filename, self.pair.filename, msg)
                self.summary_errors.append(error)

    def parse_line_type_0(self, line):
        return [
            line[0],
            line[1:17],
            line[18:19],
            line[20:22],
            line[23:29],
            line[30:55],
            line[56:61],
            line[62:73],
            line[74:79],
            line[80:119]
        ]

    def parse_line_type_1(self, line):
        return [
            line[0],
            line[1:7],
            line[8:16],
            line[17],
            line[18:19],
            line[20:29],
            line[30:61],
            line[62:79],
            line[80:86],
            line[87:95],
            line[96:111],
            line[112:119]
        ]

    def parse_line_type_7(self, line):
        return [
            line[0],
            line[1:7],
            line[8:19],
            line[20:29],
            line[30:39],
            line[40:49],
            line[50:73],
            line[74:79],
            line[80:120]
        ]

    def process_comparison(self):
        if self.pair is None:
            return None

        workbook = self.create_workbook(OUTPUT_DIR_ABA)
        worksheet = workbook.add_worksheet('ABA Comparison Results')
        fmt_error = get_error_format(workbook)

        row = 0
        col_a = 0
        col_b = 13

        keys_unmatched = set(self.pair.datarows.keys()) - set(self.datarows.keys())

        for key in self.datarows.keys():
            self_row = self.datarows[key]
            pair_row = self.pair.datarows.get(key, None)

            if pair_row is None:
                worksheet.write_row(row, col_a, self_row, fmt_error)
                error = new_error(self.filename, self.pair.filename, f'No match found for row', '', '', ' '.join(self_row))
                self.summary_errors.append(error)
            else:
                for index, value in enumerate(self_row):
                    if value != pair_row[index]:
                        print(value, pair_row[index])
                    format_ = None if u.sanitize(value) == u.sanitize(pair_row[index]) else fmt_error
                    worksheet.write(row, index, value, format_)
                    worksheet.write(row, index + col_b, pair_row[index], format_)

                    if value != pair_row[index]:
                        column = self.get_column(value[0], index)
                        error = new_error(self.filename, self.pair.filename, f'Values of {column} does not match', '', '', value, pair_row[index])
                        self.summary_errors.append(error)

            row += 1

        for key in keys_unmatched:
            worksheet.write_row(row, col_b, self.pair.datarows[key], fmt_error)
            error = new_error(self.filename, self.pair.filename, f'No match found for row', '', '', '', ' '.join(self.pair.datarows[key]))
            self.summary_errors.append(error)

            row += 1

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def get_column(self, type_, index):
        c0 = ["Record Type", "Blank1", "Reel Sequence Number", "Name of User's Financial Institution",
                "Blank2", "Name of Use supplying file1", "Name of Use supplying file2",
                "Description of entries on file", "Date to be processed", "Blank3"]

        c1 = ["Record Type 1", "Bank/State/Branch Number", "Account number to be credited/debited",
                "Indicator", "Transaction Code", "Amount", "Title of Account to be credited/debited",
                "Lodgement Reference", "Trace Record", "(Account number)", "Name of Remitter", "Amount of Withholding Tax"]

        c7 = ["Record Type 7", "BSB Format Filler", "Blank1", "File (User) Net Total Amount", "File (User) Credit Total Amount",
                "File (User) Debit Total Amount", "Blank2", "File (user) count of Records Type 1", "Blank3"]

        if type_ == 0:
            return c0[index]
        if type_ == 1:
            return c1[index]
        if type_ == 7:
            return c7[index]


def read_file_aba(file: str):
    print(f'Parsing executive summary file {bcolors.BLUE}{file}{bcolors.ENDC}')
    filename = file.split('/')[-1]
    dir_ = '/'.join(file.split('/')[:-1]) + '/'
    return ABAFile(dir_, filename)
