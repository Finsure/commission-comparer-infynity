import os

import click

from src.model import (ReferrerTaxInvoice, BrokerTaxInvoice, create_summary,
                       create_all_datailed_report, PID)

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

# Just a few colors to use in the console logs.
RED = '\033[91m'
ENDC = '\033[0m'
OKGREEN = '\033[92m'


@click.group()
def rcti():
    pass


@click.command('compare_referrer')
@click.option('-l', '--loose', type=int, default=0, help='Margin of error for a comparison to be considered correct.')
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_referrer(loose, loankit_dir, infynity_dir):
    """ A CLI for comparing the commission files between two directories """

    print("Starting referrer files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)
    loankit_files = os.listdir(loankit_dir)
    infynity_files = os.listdir(infynity_dir)

    invoices = {
        LOANKIT: _read_files_referrer(loankit_dir, loankit_files),
        INFYNITY: _read_files_referrer(infynity_dir, infynity_files)
    }

    # A list with all keys generated in both dicts
    keys_all = merge_lists(invoices[LOANKIT].keys(), invoices[INFYNITY].keys())

    results = []

    for key in keys_all:
        invoice_lkt = invoices[LOANKIT].get(key, None)
        invoice_inf = invoices[INFYNITY].get(key, None)

        # Check if its possible to do a comparison
        if invoice_lkt is not None:
            results.append(invoice_lkt.compare_to(invoice_inf, loose))
        elif invoice_inf is not None:
            results.append(invoice_inf.compare_to(invoice_lkt, loose))

    print("Creating summary...")
    create_summary(results)
    print("Creating detailed reports...")
    create_all_datailed_report(results)
    print("Finished.")


@click.command('compare_broker')
@click.option('-l', '--loose', type=int, default=0, help='Margin of error for a comparison to be considered correct.')
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcit_compare_broker(loose, loankit_dir, infynity_dir):
    print("Starting broker files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)
    loankit_files = os.listdir(loankit_dir)
    infynity_files = os.listdir(infynity_dir)

    invoices = {
        LOANKIT: _read_files_broker(loankit_dir, loankit_files),
        INFYNITY: _read_files_broker(infynity_dir, infynity_files)
    }

    # A list with all keys generated in both dicts
    keys_all = merge_lists(invoices[LOANKIT].keys(), invoices[INFYNITY].keys())


@click.command('compare_branch')
def rcti_compare_branch(loose, loankit_dir, infynity_dir):
    pass


# Add subcomands to the CLI
rcti.add_command(rcti_compare_referrer)
rcti.add_command(rcit_compare_broker)
rcti.add_command(rcti_compare_branch)


def _read_files_referrer(dir_: str, files: list) -> dict:
    keys = {}
    for filename in files:
        try:
            ti = ReferrerTaxInvoice(dir_, filename)
            keys[ti.key()] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return keys


def _read_files_broker(dir_: str, files: list) -> dict:
    keys = {}
    for filename in files:
        try:
            ti = BrokerTaxInvoice(dir_, filename)
            keys[ti.key] = ti
            break
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return keys


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
    rcti()

# python cli.py compare_referrer -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Loankit/Sent/" "/Users/petrosschilling/dev/commission-comparer-infynity/Referrers/Infynity/Sent/"

# python cli.py compare_broker -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/Brokers/Loankit/" "/Users/petrosschilling/dev/commission-comparer-infynity/Brokers/Infynity/"

# python app.py --help
