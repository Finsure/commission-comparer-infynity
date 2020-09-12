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

def move_files_to_dir(pre_processing_dir, input_dir, run_date):
    print(f"""Pre Processing DIR: {pre_processing_dir}""")
    print(f"""INPUT DIR: {input_dir}""")
    print(f"""Run Date: {run_date}""")
    #pdb.set_trace()
    #org_directory = main_dir + "/" + group 
    new_directory = input_dir + "/" + run_date
    _build_dir(new_directory)
    print(f"""new directory: {new_directory}""")
    copy_branch_files(pre_processing_dir, run_date, new_directory)
    copy_broker_files(pre_processing_dir, run_date, new_directory)
    copy_referrer_files(pre_processing_dir, run_date, new_directory)
    copy_executive_summary_file(pre_processing_dir, run_date, new_directory)
    copy_de_file(pre_processing_dir, run_date, new_directory)

def copy_de_file(pre_processing_dir, run_date, new_dir):
    de_file_pre_processing_dir = f"""{pre_processing_dir}/{run_date}/de_file"""
    de_file_summ_dir = new_dir + "/de_file"
    files = os.listdir(de_file_pre_processing_dir)
    for f in files:
        filename = de_file_pre_processing_dir + '/' + f
        if os.path.isfile(filename):
            search_pattern = '_DE_'
            if (re.search(search_pattern, filename, re.IGNORECASE)):
                shutil.copy(filename, de_file_summ_dir)
                print(f"""File Name: {filename} | de_file_summ_dir: {de_file_summ_dir}| FileName: {filename}""")

def copy_executive_summary_file(pre_processing_dir, run_date, new_dir):
    exec_summ_pre_processing_dir = f"""{pre_processing_dir}/{run_date}/executive_summary_report"""
    exec_summ_dir = new_dir + "/executive_summary_report"
    files = os.listdir(exec_summ_pre_processing_dir)
    for f in files:
        filename = exec_summ_pre_processing_dir + '/' + f
        if os.path.isfile(filename):
            search_pattern = '_ES_Report_'
            if (re.search(search_pattern, filename, re.IGNORECASE)):
                shutil.copy(filename, exec_summ_dir)
                print(f"""File Name: {filename} | exec_summ Dir: {exec_summ_dir}| FileName: {filename}""")

def copy_referrer_files(pre_processing_dir, run_date, new_dir):
    referrer_pre_processing_dir = f"""{pre_processing_dir}/{run_date}/referrers"""
    referrer_dir = new_dir + "/referrers"
    files = os.listdir(referrer_pre_processing_dir)
    for f in files:
        filename = referrer_pre_processing_dir + '/' + f
        if os.path.isfile(filename):
            search_pattern = '_referrer_RCTI_'
            if (re.search(search_pattern, filename, re.IGNORECASE)):
                if filename.endswith("html"):
                    shutil.copy(filename, referrer_dir)
                    print(f"""File Name: {filename} | referrer Dir: {referrer_dir}| FileName: {filename}""")
                else:
                    pass
                    print(f"""Ignore File Name: {filename} | referrer Dir: {referrer_dir}| FileName: {filename}""")

def copy_broker_files(pre_processing_dir, run_date, new_dir):
    broker_pre_processing_dir = f"""{pre_processing_dir}/{run_date}/brokers/rcti"""
    broker_dir = new_dir + "/brokers"
    files = os.listdir(broker_pre_processing_dir)
    for f in files:
        filename = broker_pre_processing_dir + '/' + f
        if os.path.isfile(filename):
            search_pattern = '_Broker_RCTI_'
            if (re.search(search_pattern, filename, re.IGNORECASE)):
                if filename.endswith("xlsx"):
                    shutil.copy(filename, broker_dir)
                    print(f"""File Name: {filename} | Broker Dir: {broker_dir}| FileName: {filename}""")
                else:
                    pass
                    print(f"""Ignore File Name: {filename} | Broker Dir: {broker_dir}| FileName: {filename}""")

def copy_branch_files(pre_processing_dir, run_date, new_dir):
    branch_pre_processing_dir = f"""{pre_processing_dir}/{run_date}/branch/rcti""" 
    branch_dir = new_dir + "/branch"
    files = os.listdir(branch_pre_processing_dir)
    for f in files:
       filename = branch_pre_processing_dir + '/' + f
       if os.path.isfile(filename):
           search_pattern = '_Branch_'
           if (re.search(search_pattern, filename, re.IGNORECASE)):
               branch_id = filename.split(search_pattern)[0].split('_')[-1]
               if filename.endswith("xlsx"):
                   shutil.copy(filename, branch_dir)
                   print(f"""File Name: {filename} | Branch Dir: {branch_dir}| Branch ID: {branch_id}""")
               else:
                   pass
                   print(f"""Ignore File Name: {filename} | Branch Dir: {branch_dir}| Branch ID: {branch_id}""")


           #move_broker_files(file_name, new_directory)
           #move_referrer_files(file_name, new_directory)
           #move_exec_summary(file_name, new_directory)
           #move_de_file(group, file_name, new_directory)
           #move_other_files(file_name, new_directory)
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

def _build_dir(directory):
    directory_branch = directory + "/branch"
    directory_broker = directory + "/brokers"
    directory_referrer = directory + "/referrers"
    directory_es = directory + "/executive_summary_report"
    directory_aba = directory + "/de_file"
    _check_if_directory_exists_or_create(directory_branch)
    _check_if_directory_exists_or_create(directory_broker)
    _check_if_directory_exists_or_create(directory_referrer)
    _check_if_directory_exists_or_create(directory_es)
    _check_if_directory_exists_or_create(directory_aba)


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
    pre_processing_dir = "/var/www/mystro.com/data/commission_loankit/rcti/pre_processed"
    input_dir = "/home/qaisar/rcti_comparison/commission-comparer-infynity/inputs/infynity"
    run_date = "15034_Tue_Jul_14_2020"
    #group = input("Please enter LoanKit or Finsure\n\n")
    #process_id = input("Please enter the process id\n\n")
    #group = "loankit"
    move_files_to_dir(pre_processing_dir, input_dir, run_date) 

