import os

import click
import xlsxwriter

from src.model.taxinvoice import (create_dirs, new_error, write_errors, get_header_format,
                                  get_title_format, PID, OUTPUT_DIR_SUMMARY)
from src.model.taxinvoice_referrer import read_files_referrer
from src.model.taxinvoice_broker import read_files_broker
from src.model.taxinvoice_branch import read_files_branch
from src.model.executive_summary import read_file_exec_summary
from src.model.aba import read_file_aba
from src.utils import bcolors


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
    print_start_message('referrer')
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

    print_done_message()


# @click.command('compare_broker')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_dir', required=True, type=click.Path(exists=True))
# @click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_broker(loose, loankit_dir, infynity_dir):
    print_start_message('broker')
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

    print_done_message()


# @click.command('compare_branch')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_dir', required=True, type=click.Path(exists=True))
# @click.argument('infynity_dir', required=True, type=click.Path(exists=True))
def rcti_compare_branch(loose, loankit_dir, infynity_dir):
    print_start_message('branch')
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

    print_done_message()


# @click.command('compare_executive_summary')
# @click.option('-l', '--loose', type=float, default=0, help=DESC_LOOSE)
# @click.argument('loankit_file', required=True, type=click.File(exists=True))
# @click.argument('infynity_file', required=True, type=click.File(exists=True))
def rcti_compare_executive_summary(loose, loankit_file, infynity_file):
    print_start_message('executive summary')
    exec_summary_infynity = read_file_exec_summary(infynity_file)
    exec_summary_loankit = read_file_exec_summary(loankit_file)

    exec_summary_infynity.pair = exec_summary_loankit
    exec_summary_infynity.margin = loose
    create_dirs()
    summary_errors = exec_summary_infynity.process_comparison(margin=loose)

    # Create summary based on errors
    file = f"{OUTPUT_DIR_SUMMARY}{'Final Summary'}.xlsx"
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet('Summary')
    fmt_title = get_title_format(workbook)
    fmt_table_header = get_header_format(workbook)
    worksheet.merge_range('A1:I1', 'Summary', fmt_title)
    row = 1
    col = 0
    worksheet.write(row, col, f"Number of issues: {str(len(summary_errors))}")
    row += 2
    worksheet = write_errors(summary_errors, worksheet, row, col, fmt_table_header,
                             exec_summary_infynity.directory, exec_summary_loankit.directory)
    workbook.close()

    print_done_message()


# @click.command('compare_aba')
# @click.argument('loankit_file', required=True, type=click.File(exists=True))
# @click.argument('infynity_file', required=True, type=click.File(exists=True))
def rcti_compare_aba(loankit_file, infynity_file):
    print_start_message('aba files')
    aba_infynity = read_file_aba(infynity_file)
    aba_loankit = read_file_aba(loankit_file)

    aba_infynity.pair = aba_loankit
    create_dirs()
    summary_errors = aba_infynity.process_comparison()

    # Create summary based on errors
    file = f"{OUTPUT_DIR_SUMMARY}{'ABA Summary'}.xlsx"
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet('ABA Comparison Results')
    fmt_title = get_title_format(workbook)
    fmt_table_header = get_header_format(workbook)
    worksheet.merge_range('A1:I1', 'Summary', fmt_title)
    row = 1
    col = 0
    worksheet.write(row, col, f"Number of issues: {str(len(summary_errors))}")
    row += 2
    worksheet = write_errors(summary_errors, worksheet, row, col, fmt_table_header,
                             aba_infynity.directory, aba_loankit.directory)
    workbook.close()

    print_done_message()


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
    file = f"{OUTPUT_DIR_SUMMARY}{summary_filname}.xlsx"
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet('Summary')
    fmt_title = get_title_format(workbook)
    fmt_table_header = get_header_format(workbook)
    worksheet.merge_range('A1:I1', summary_title, fmt_title)
    row = 1
    col = 0
    worksheet.write(row, col, f"Number of issues: {str(len(summary_errors))}")
    row += 2
    worksheet = write_errors(summary_errors, worksheet, row, col, fmt_table_header, filepath_a, filepath_b)
    workbook.close()


# Add subcommands to the CLI
# rcti.add_command(rcti_compare_referrer)
# rcti.add_command(rcti_compare_broker)
# rcti.add_command(rcti_compare_branch)
# rcti.add_command(rcti_compare_executive_summary)

def print_start_message(type: str):
    print(f"{bcolors.BOLD}Starting {type} files comparison...{bcolors.ENDC}")
    print(f"This Process ID (PID) is: {bcolors.GREEN}{PID}{bcolors.ENDC}")


def print_done_message():
    print(f"{bcolors.GREEN}DONE{bcolors.ENDC}")


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
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/15457/referrer/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/15457/referrer/')
    # rcti_compare_broker(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/15457/broker/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/15457/broker/')
    # rcti_compare_branch(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/15457/branch/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/15457/branch/')
    # rcti_compare_executive_summary(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/15457/executive_summary/Finsure_ES_Report_17129_Sun_May_10_2020.xls',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/15457/executive_summary/Finsure_ES_Report_6602__Sun_May_10_2020.xlsx')
    # rcti_compare_aba(
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/15457/de_file/Finsure_DE_2020-05-10.txt',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/15457/de_file/Finsure_DE_File_6602__Sun_May_10_2020.txt')

    # rcti_compare_referrer(
    #     0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/10511/referrer/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/10511/referrer/')
    # rcti_compare_broker(
    #     0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/10511/broker/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/10511/broker/')
    # rcti_compare_branch(
    #     0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/10511/branch/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/10511/branch/')
    # rcti_compare_executive_summary(
    #     0,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/10511/executive_summary/Finsure_ES_Report_17076_Sun_May_10_2020.xls',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/10511/executive_summary/Finsure_ES_Report_6645__Sun_May_10_2020.xlsx')
    # rcti_compare_aba(
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/10511/de_file/Finsure_DE_2020-05-10.txt',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/10511/de_file/Finsure_DE_File_6645__Sun_May_10_2020.txt')

    # rcti_compare_referrer(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/7127/referrer/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/7127/referrer/')
    # rcti_compare_broker(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/7127/broker/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/7127/broker/')
    # rcti_compare_branch(
    #     1,
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/7127/branch/',
    #     '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/7127/branch/')
    rcti_compare_executive_summary(
        1,
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/7127/executive_summary/Finsure_ES_Report_17104_Sun_May_10_2020.xls',
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/7127/executive_summary/Finsure_ES_Report_6624__Sun_May_10_2020.xlsx')
    rcti_compare_aba(
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/7127/de_file/Finsure_DE_2020-05-10.txt',
        '/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/7127/de_file/Finsure_DE_File_6624__Sun_May_10_2020.txt')


# SIMULATE REFERRER
# python cli.py compare_referrer -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/referrer/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/referrer/"

# SIMULATE BROKER
# python cli.py compare_broker -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/broker/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/broker/"

# SIMULATE BRANCH
# python cli.py compare_branch -l 0 "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/loankit/branch/" "/Users/petrosschilling/dev/commission-comparer-infynity/inputs/infynity/branch/"

# python app.py --help
