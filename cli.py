import os

import click
import xlsxwriter

from src.model.taxinvoice import (create_dirs, new_error, write_errors, get_header_format,
                                  get_title_format, PID, OUTPUT_DIR_SUMMARY)
from src.model.taxinvoice_referrer import read_files_referrer
from src.model.taxinvoice_broker import read_files_broker
from src.model.taxinvoice_branch import read_files_branch
from src.utils import OKGREEN, ENDC


# Constants
INFYNITY = 'Infynity'
LOANKIT = 'Loankit'

DESC_LOOSE = 'Margin of error for a comparison between two numbers to be considered correct.'


@click.group()
def rcti():
    pass


# @click.command('compare_referrer')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_dir', required=True, type=click.Path(exists=True))
# @click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_referrer(loose, loankit_dir, infynity_dir):
    """ A CLI for comparing the commission files between two directories """

    print("Starting referrer files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)
    loankit_files = list_files(loankit_dir)
    infynity_files = list_files(infynity_dir)

    invoices_loankit = read_files_referrer(loankit_dir, loankit_files)
    invoices_infynity = read_files_referrer(infynity_dir, infynity_files)

    run_comparison(
        invoices_loankit,
        invoices_infynity,
        loose,
        'referrer_rcti_summary',
        'Commission Referrer RCTI Summary',
        loankit_dir,
        infynity_dir)

    print(OKGREEN + ' DONE' + ENDC)


# @click.command('compare_broker')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_dir', required=True, type=click.Path(exists=True))
# @click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_broker(loose, loankit_dir, infynity_dir):
    print("Starting broker files comparison...")
    print('This Process ID (PID) is: ' + OKGREEN + PID + ENDC)

    files_loankit = list_files(loankit_dir)
    files_infynity = list_files(infynity_dir)

    invoices_loankit = read_files_broker(loankit_dir, files_loankit)
    invoices_infynity = read_files_broker(infynity_dir, files_infynity)

    run_comparison(
        invoices_loankit,
        invoices_infynity,
        loose,
        'broker_rcti_summary',
        'Commission Broker RCTI Summary',
        loankit_dir,
        infynity_dir)

    print(OKGREEN + 'DONE' + ENDC)


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

    run_comparison(
        invoices_loankit,
        invoices_infynity,
        loose,
        'branch_rcti_summary',
        'Commission Branch RCTI Summary',
        loankit_dir,
        infynity_dir)

    print(OKGREEN + 'DONE' + ENDC)


def run_comparison(files_a, files_b, margin, summary_filname, summary_title, filepath_a, filepath_b):
    create_dirs()

    summary_errors = []

    # Set each invoice pair
    for key in files_a.keys():
        if files_b.get(key, None) is not None:
            files_a[key].pair = files_b[key]
            files_b[key].pair = files_a[key]
        else:
            # Log in the summary files that don't have a match
            msg = 'No corresponding commission file found'
            error = new_error(files_a[key].filename, '', msg)
            summary_errors.append(error)

    # Find all Infynity files that don't have a match
    alone_keys_infynity = set(files_b.keys()) - set(files_a.keys())
    for key in alone_keys_infynity:
        msg = 'No corresponding commission file found'
        error = new_error('', files_b[key].filename, msg)
        summary_errors.append(error)

    counter = 1
    for key in files_a.keys():
        print(f'Processing {counter} of {len(files_a)} files', end='\r')
        errors = files_a[key].process_comparison(margin)
        if errors is not None:
            summary_errors = summary_errors + errors
        counter += 1
    print()

    # Create summary based on errors
    workbook = xlsxwriter.Workbook(OUTPUT_DIR_SUMMARY + summary_filname + '.xlsx')
    worksheet = workbook.add_worksheet('Summary')
    fmt_title = get_title_format(workbook)
    fmt_table_header = get_header_format(workbook)
    worksheet.merge_range('A1:I1', summary_title, fmt_title)
    row = 1
    col = 0
    worksheet.write(row, col, 'Number of issues: ' + str(len(summary_errors)))
    row += 2
    worksheet = write_errors(summary_errors, worksheet, row, col, fmt_table_header, filepath_a, filepath_b)
    workbook.close()


# Add subcommands to the CLI
# rcti.add_command(rcti_compare_referrer)
# rcti.add_command(rcti_compare_broker)
# rcti.add_command(rcti_compare_branch)


def list_files(dir_: str) -> list:
    files = []
    with os.scandir(dir_) as it:
        for entry in it:
            if not entry.name.startswith('.') and not entry.name.startswith('~') and entry.is_file():
                files.append(entry.name)
    return files


if __name__ == '__main__':
    # rcti()
    # rcti_compare_referrer(
    #     0.0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/referrer/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/referrer/')
    rcti_compare_broker(
        0.0,
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/broker/',
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/broker/')
    # rcti_compare_branch(
    #     0.0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/branch/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/branch/')

# SIMULATE REFERRER
# python cli.py compare_referrer -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/referrer/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/referrer/"

# SIMULATE BROKER
# python cli.py compare_broker -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/broker/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/broker/"

# SIMULATE BRANCH
# python cli.py compare_branch -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/branch/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/branch/"

# python app.py --help
