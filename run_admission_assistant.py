# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 23:46:47 2023

@author: Yue
"""

import os
import argparse

import write_spreadsheet
import write_applicant_records

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Outputs recommended admission decisions given reviewers' feedback"
    )
    parser.add_argument(
        "-d", help="enter the path to the reviewer feedback spreadsheets"
    )
    args, spreadsheets_dir = parser.parse_known_args()
    if args.d is not None:
        data_path = args.d
    elif spreadsheets_dir:
        data_path = spreadsheets_dir[0]
    else:
        print("No argument provided")

    if os.path.isdir(data_path):
        write_applicant_records.main(data_path)
        write_spreadsheet.main(data_path)
    else:
        print("invalid spreadsheets directory.")
