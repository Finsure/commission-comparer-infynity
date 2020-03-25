
import pandas

from src.model.taxinvoice import TaxInvoice
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
        self.tabs_contents = {}
        self.parse()

    def parse(self):
        xl = pandas.ExcelFile(self.full_path)
        for sn in xl.sheet_names:
            df = xl.parse(sn)
            df = df.dropna(how='all')  # remove rows that don't have any value

            if sn == 'Branch Summary Report' or sn == 'Branch Fee Summary Report':
                df = df.rename(columns=df.iloc[0]).drop(df.index[0])  # Make first row the table header
                # print(df.to_dict())

            self.tabs_contents[sn] = df.to_dict()



def read_file_exec_summary(file: str):
    print(f'Parsing executive summary file {bcolors.BLUE}{file}{bcolors.ENDC}')
    filename = file.split('/')[-1]
    dir_ = '/'.join(file.split('/')[:-1]) + '/'
    return ExecutiveSummary(dir_, filename)
