import pwd
import grp
import sys
import os
import shutil
import re
import pdb
#import datetime
from datetime import datetime as dt
import time
import math
import pprint
import xlsxwriter
from decimal import *
from pathlib import Path
from pprint import pprint


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

def move_files_to_dir(main_dir, group, process_id):
    print(f"""Main DIR: {main_dir}""")
    print(f"""Group ID: {group}""")
    print(f"""Process ID: {process_id}""")
    #pdb.set_trace()
    org_directory = main_dir + "/" + group 
    new_directory = main_dir + "/" + group + "/" + process_id
    _build_dir(new_directory)
    print(f"""org_directory: {org_directory}| new directory: {new_directory}""")
    files = os.listdir(org_directory)
    for f in files:
       file_name = org_directory + '/' + f
       if os.path.isfile(file_name):
           move_branch_files(file_name, new_directory)
           move_broker_files(file_name, new_directory)
           move_referrer_files(file_name, new_directory)
           move_exec_summary(file_name, new_directory)
           move_de_file(group, file_name, new_directory)
           move_other_files(file_name, new_directory)
       #print(f"""File Name: {file_name} | Branch ID: {branch_id}""")
       #pdb.set_trace()

def move_other_files(filename, new_dir):
    extra_dir = new_dir + "/extra_files"
    search_pattern = '_Broker_Summary_Report_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        shutil.move(filename, extra_dir)
    search_pattern = 'release.txt'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        shutil.move(filename, extra_dir)
        #pdb.set_trace()

def move_branch_files(filename, new_dir):
    branch_dir = new_dir + "/branch"
    search_pattern = 'LoanKit_Branch_Summary_Report_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        os.remove(filename)
        print(f"""Remove File Name: {filename} | Branch Dir: {branch_dir}""")
    else:
        search_pattern = '_Branch_'
        if (re.search(search_pattern, filename, re.IGNORECASE)):
            branch_id = filename.split(search_pattern)[0].split('_')[-1]
            if filename.endswith("xls"):
                shutil.move(filename, branch_dir)
                print(f"""File Name: {filename} | Branch Dir: {branch_dir}| Branch ID: {branch_id}""")
            else:
                os.remove(filename)
                print(f"""Remove File Name: {filename} | Branch Dir: {branch_dir}| Branch ID: {branch_id}""")


def move_broker_files(filename, new_dir):
    broker_dir = new_dir + "/brokers"
    branch_id = ''
    search_pattern = '_Broker_RCTI_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        #branch_id = filename.split(search_pattern)[0].split('_')[-1]
        if filename.endswith("xls"):
            shutil.move(filename, broker_dir)
            print(f"""File Name: {filename} | Broker Dir: {broker_dir}| FileName: {filename}""")
        else:
            os.remove(filename)
            print(f"""Remove File Name: {filename} | Broker Dir: {broker_dir}| FileName: {filename}""")

def move_referrer_files(filename, new_dir):
    referrer_dir = new_dir + "/referrers"
    branch_id = ''
    search_pattern = '_Referrer_RCTI_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        #branch_id = filename.split(search_pattern)[0].split('_')[-1]
        shutil.move(filename, referrer_dir)

def move_exec_summary(filename, new_dir):
    exec_summ_dir = new_dir + "/executive_summary_report"
    branch_id = ''
    search_pattern = '_ES_Report'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        #branch_id = filename.split(search_pattern)[0].split('_')[-1]
        shutil.move(filename, exec_summ_dir)

def move_de_file(group, filename, new_dir):
    suffix = group.capitalize()
    de_file_dir = new_dir + "/de_file"
    branch_id = ''
    search_pattern = '_DE_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        #branch_id = filename.split(search_pattern)[0].split('_')[-1]
        if filename.endswith(".txt"):
            shutil.move(filename, de_file_dir)

def get_branch_id_org(filename):
    branch_id = ''
    today = dt.now()
    thisMonth = today.strftime('%Y%m')
    #search_pattern = '_Branch_'+ str(thisMonth) + '_'
    search_pattern = '_Branch_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        branch_id = filename.split(search_pattern)[0].split('_')[-1]
        search_pattern = '_Broker_RCTI_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        branch_id = filename.split(search_pattern)[0].split('_')[-1]
        search_pattern = '_' + str(thisMonth) + '_Referrer_RCTI_'
    if (re.search(search_pattern, filename, re.IGNORECASE)):
        branch_id = filename.split(search_pattern)[0].split('_')[-1]
    return branch_id

def _build_dir(directory):
    #referrer_name = replace_non_alphanumeric_character_with_underscore(self.referrer_scheme_list[branch_id][referrer_scheme_id]['referrer_name'])
    #referrer_company_name = replace_non_alphanumeric_character_with_underscore(self.referrer_scheme_list[branch_id][referrer_scheme_id]['company_name'])
    #branch_company_name = replace_non_alphanumeric_character_with_underscore(self.branch_dict[branch_id]['general']['company_name'])
    #(referrer_company_exists, file_name) = self._get_referrer_file_name(branch_id, branch_company_name, referrer_scheme_id, referrer_name, referrer_company_name)
    #suffix = self.suffix.lower()
    #data_folder = '/var/www/mystro.com/data'
    #commission_folder = '/commission_'+suffix
    #pre_processing_folder = f"/rcti/pre_processed/{self.run_date}/referrers"
    directory_branch = directory + "/branch"
    directory_broker = directory + "/brokers"
    directory_referrer = directory + "/referrers"
    directory_es = directory + "/executive_summary_report"
    directory_aba = directory + "/de_file"
    directory_other = directory + "/extra_files"
    _check_if_directory_exists_or_create(directory_branch)
    _check_if_directory_exists_or_create(directory_broker)
    _check_if_directory_exists_or_create(directory_referrer)
    _check_if_directory_exists_or_create(directory_es)
    _check_if_directory_exists_or_create(directory_aba)
    _check_if_directory_exists_or_create(directory_other)
    #file_full_path = directory + "/" + file_name
    #return (referrer_company_exists, referrer_company_name, referrer_name, branch_company_name, file_full_path)


def _check_if_directory_exists_or_create(path, owner='apache', group='apache', perms=0o755, force=False):
    uid = pwd.getpwnam(owner).pw_uid
    gid = grp.getgrnam(group).gr_gid
    realpath = os.path.abspath(path)
    path_exists = os.path.exists(realpath)
    if path_exists and force:
        if not os.path.isdir(realpath):
            log("Removing non-directory file {} prior to mkdir()".format(path))
            os.unlink(realpath)
            os.makedirs(realpath, perms)
    elif not path_exists:
        """Create a directory"""
        print("Making dir {} {}:{} {:o}".format(path, owner, group, perms))
        os.makedirs(realpath, perms)
        os.chown(realpath, uid, gid)
        os.chmod(realpath, perms)

if __name__ == '__main__':
    main_dir = "/home/qaisar/rcti_comparison/commission-comparer-infynity/inputs"
    #group = input("Please enter LoanKit or Finsure\n\n")
    #process_id = input("Please enter the process id\n\n")
    group = "loankit"
    process_id = "22369"
    move_files_to_dir(main_dir, group.lower(), process_id )

