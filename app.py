import os
import click

from src.model import TaxInvoice


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


@click.command()
@click.option('-l', '--loose', type=int, default=0, help='Margin of error for a comparison to be considered correct.')
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def compare_referrer_rcti(loose, loankit_dir, infynity_dir):
    """ A CLI for comparing the commission files between two directories """

    loankit_files = os.listdir(loankit_dir)
    infynity_files = os.listdir(infynity_dir)

    invoices = {
        INFYNITY: _read_files(infynity_dir, infynity_files),
        LOANKIT: _read_files(loankit_dir, loankit_files)
    }

    keys_all = _merge_lists(invoices[LOANKIT].keys(), invoices[INFYNITY].keys())

    for key in keys_all:
        invoice_inf = invoices[INFYNITY].get(key, None)
        invoice_lkt = invoices[LOANKIT].get(key, None)

        # Chek if its possible to do a comparison
        if invoice_inf is not None:
            print('Not matching file for: ' + invoice_inf.get_full_path())
            continue
        if invoice_lkt is not None:
            print('Not matching file for: ' + invoice_lkt.get_full_path())
            continue



        comparison = compare()





    # print(invoices)


def _merge_lists(l1: list, l2: list) -> list:
    difference = list(set(l2) - set(l1))
    return list(l1) + difference


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

# python app.py -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Loankit/Sent/" "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Infynity/Sent/"

# python app.py --help
