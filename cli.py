import os

import click

from src.model.taxinvoice import create_detailed_dir, create_summary_dir, PID
from src.model.taxinvoice_referrer import (create_summary_referrer, create_detailed_referrer,
                                           read_files_referrer)
from src.model.taxinvoice_broker import (create_summary_broker, create_detailed_broker,
                                         read_files_broker)
from src.utils import merge_lists, OKGREEN, ENDC


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

DESC_LOOSE = 'Margin of error for a comparison between two numbers to be considered correct.'

@click.group()
def rcti():
    pass




@click.command('compare_referrer')
@click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_referrer(loose, loankit_dir, infynity_dir):
    """ A CLI for comparing the commission files between two directories """

    print("Starting referrer files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)
    loankit_files = os.listdir(loankit_dir)
    infynity_files = os.listdir(infynity_dir)

    invoices = {
        LOANKIT: read_files_referrer(loankit_dir, loankit_files),
        INFYNITY: read_files_referrer(infynity_dir, infynity_files)
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

    create_summary_dir()
    create_detailed_dir()

    print("Creating summary...", end='')
    create_summary_referrer(results)
    print(OKGREEN + ' OK' + ENDC)

    print("Creating detailed reports...", end='')
    for result in results:
        create_detailed_referrer(result)
    print(OKGREEN + ' OK' + ENDC)

    print("Finished.")


@click.command('compare_broker')
@click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
@click.argument('loankit_dir', required=True, type=click.Path(exists=True))
@click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcit_compare_broker(loose, loankit_dir, infynity_dir):
    print("Starting broker files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)

    loankit_files = list_files(loankit_dir)
    infynity_files = list_files(infynity_dir)

    invoices = {
        LOANKIT: read_files_broker(loankit_dir, loankit_files),
        INFYNITY: read_files_broker(infynity_dir, infynity_files)
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

    create_summary_dir()
    create_detailed_dir()

    print("Creating summary...", end='')
    create_summary_broker(results)
    print(OKGREEN + ' OK' + ENDC)

    print("Creating detailed reports...", end='')
    for result in results:
        create_detailed_broker(result)
    print(OKGREEN + ' OK' + ENDC)


@click.command('compare_branch')
def rcti_compare_branch(loose, loankit_dir, infynity_dir):
    pass


# Add subcommands to the CLI
rcti.add_command(rcti_compare_referrer)
rcti.add_command(rcit_compare_broker)
rcti.add_command(rcti_compare_branch)


def list_files(dir_: str) -> list:
    files = []
    with os.scandir(dir_) as it:
        for entry in it:
            if not entry.name.startswith('.') and not entry.name.startswith('~') and entry.is_file():
                files.append(entry.name)
    return files


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
