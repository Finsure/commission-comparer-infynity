import os

import click

from src.model import TaxInvoice, create_summary, create_all_datailed_report, PID

from src.utils import merge_lists


# Constants
AMOUNT_OF_FILES = 'Amount of Files'
INFYNITY = 'Infynity'
LOANKIT = 'Loankit'

SUMMARY = 'Summary'

SUMMARY_REPORT = {
    AMOUNT_OF_FILES: {
        INFYNITY: 0,
        LOANKIT: 0
    },
    SUMMARY: []
}

RED = '\033[91m'
ENDC = '\033[0m'
OKGREEN = '\033[92m'


@click.command()
@click.option('-l', '--loose', type=int, default=0, help='Margin of error for a comparison to be considered correct.')
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def compare_referrer_rcti(loose, loankit_dir, infynity_dir):
    """ A CLI for comparing the commission files between two directories """

    print("Starting referrer files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)
    loankit_files = os.listdir(loankit_dir)
    infynity_files = os.listdir(infynity_dir)

    invoices = {
        LOANKIT: _read_files(loankit_dir, loankit_files),
        INFYNITY: _read_files(infynity_dir, infynity_files)
    }

    keys_all = merge_lists(invoices[LOANKIT].keys(), invoices[INFYNITY].keys())

    results = []

    for key in keys_all:
        invoice_lkt = invoices[LOANKIT].get(key, None)
        invoice_inf = invoices[INFYNITY].get(key, None)

        # Chek if its possible to do a comparison
        if invoice_lkt is not None:
            results.append(invoice_lkt.compare_to(invoice_inf, loose))
        elif invoice_inf is not None:
            results.append(invoice_inf.compare_to(invoice_lkt, loose))

    # print(results)
    print("Creating summary...")
    create_summary(results)
    print("Creating detailed reports...")
    create_all_datailed_report(results)
    print("Finished.")


def _read_files(dir_: str, files: list) -> dict:
    results = {}
    for filename in files:
        try:
            ti = TaxInvoice(dir_, filename)
            results[ti.key()] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return results


def new_summary_row():
    return {
        'Transaction Line Number': "",
        'Key': "",
        'Infynity Amount Paid': "",
        'Loankit Amount Paid': "",
        'Infynity GST Paid': "",
        'Loankit GST Paid': "",
        'Infynity Total Amount Paid': "",
        'Loankit Total Amount Paid': ""
    }


if __name__ == '__main__':
    compare_referrer_rcti()

# python cli.py -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Loankit/Sent/" "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Infynity/Sent/"

# python app.py --help
