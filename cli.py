import os

import click
import xlsxwriter

from src.model.taxinvoice import create_detailed_dir, create_summary_dir, new_error, write_errors, PID, OUTPUT_DIR_SUMMARY_PID
from src.model.taxinvoice_referrer import (create_summary_referrer, create_detailed_referrer,
                                           read_files_referrer)
from src.model.taxinvoice_broker import (create_summary_broker, create_detailed_broker,
                                         read_files_broker)
from src.model.taxinvoice_branch import (read_files_branch)
from src.utils import merge_lists, OKGREEN, ENDC


# Constants
AMOUNT_OF_FILES = 'Amount of Files'
INFYNITY = 'Infynity'
LOANKIT = 'Loankit'
SUMMARY = 'Summary'

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


# @click.command('compare_branch')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_dir', required=True, type=click.Path(exists=True))
# @click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_branch(loose, loankit_dir, infynity_dir):
    print("Starting branch files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)

    files_loankit = list_files(loankit_dir)
    files_infynity = list_files(infynity_dir)

    invoices_loankit = read_files_branch(loankit_dir, files_loankit)
    invoices_infynity = read_files_branch(infynity_dir, files_infynity)

    create_summary_dir()
    create_detailed_dir()

    summary_errors = []

    # Set each invoice pair
    for key in invoices_loankit.keys():
        if invoices_infynity.get(key, None) is not None:
            invoices_loankit[key].pair = invoices_infynity[key]
            invoices_infynity[key].pair = invoices_loankit[key]
        else:
            # Log in the summary files that don't have a match
            msg = 'No corresponding commission file found'
            error = new_error(invoices_loankit[key].filename, msg)
            summary_errors.append(error)

    # Fin all Infynity files that don't have a match
    alone_keys_infynity = set(invoices_infynity.keys()) - set(invoices_loankit.keys())
    for key in alone_keys_infynity:
        msg = 'No corresponding commission file found'
        error = new_error(invoices_infynity[key].filename, msg)
        summary_errors.append(error)

    for key in invoices_loankit.keys():
        errors = invoices_loankit[key].process_comparison(loose)
        summary_errors = summary_errors + errors

    # Create summary based on errors
    workbook = xlsxwriter.Workbook(OUTPUT_DIR_SUMMARY_PID + 'branch_rcti_summary.xlsx')
    worksheet = workbook.add_worksheet('Summary')
    row = 0
    col = 0
    fmt_title = workbook.add_format({'font_size': 20, 'bold': True})
    fmt_table_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
    worksheet.merge_range('A1:I1', 'Commission Branch RCTI Summary', fmt_title)
    row += 1
    worksheet.write(row, col, 'Number of issues: ' + str(len(summary_errors)))
    row += 2
    worksheet = write_errors(summary_errors, worksheet, row, col, fmt_table_header)
    workbook.close()

    print(OKGREEN + 'OK' + ENDC)


# Add subcommands to the CLI
rcti.add_command(rcti_compare_referrer)
rcti.add_command(rcit_compare_broker)
# rcti.add_command(rcti_compare_branch)


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
    # rcti()
    rcti_compare_branch(
        0.5,
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/branch/',
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/branch/')

# SIMULATE REFERRER
# python cli.py compare_referrer -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/referrer/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/referrer/"

# SIMULATE BROKER
# python cli.py compare_broker -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/broker/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/broker/"

# SIMULATE BRANCH
# python cli.py compare_branch -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/branch/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/branch/"

# python app.py --help
