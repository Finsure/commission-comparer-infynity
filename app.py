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

    files_loankit = os.listdir(loankit_dir)
    files_infynity = os.listdir(infynity_dir)

    # print("Loankit commission files:")
    # for filename in files_loankit:
    #     print(filename)

    # print("\n")
    # print("Infynity commission files:")
    # for filename in files_infynity:
    #     print(filename)

    print(TaxInvoice(loankit_dir, files_loankit[0]).serialize())


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
