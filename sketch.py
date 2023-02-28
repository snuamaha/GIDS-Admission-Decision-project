# -*- coding: utf-8 -*-
"""
Created on Mon Feb 27 09:26:01 2023

@author: Yue
"""

import os
import re
import logging

import numpy as np
import pandas as pd

from util_vars import decision_matrix, rename_columns, reviewer_items, selected_cols


DATA_PATH = r"./ROUND 1 Reviews/"
error_logger = logging.getLogger("error_logger")
error_logger.setLevel('WARNING')
file_handler = logging.FileHandler(filename=DATA_PATH+"errors.log", mode="w")
error_logger.addHandler(file_handler)

console_logger = logging.getLogger("console_logger")
console_logger.setLevel('INFO')
file_handler = logging.StreamHandler()
console_logger.addHandler(file_handler)

files = [
    'ROUND 1 - CHINA_He_Hangfeng.xlsx',
    'ROUND 1 - CHINA_Li_Dongmei.xlsx',
    'ROUND 1 - CHINA_LUO_Jiebo.xlsx',
    'ROUND 1 - INDIA_Kahng_Anson.xlsx'
]
files = os.listdir(DATA_PATH)
gpa_pattern = re.compile(r'\d.[\d]+')
dtype_map = {"GPA 1": str}

#%%
def process_spreadsheet(excel_file) -> dict:
    excel_file = DATA_PATH + file
    reviewer_name = " ".join(excel_file.rstrip(".xlsx").split("_")[1:])
    df = pd.read_excel(excel_file, header=1)
    df.columns = [header.strip() for header in df.columns] # remove leading and trailing spaces
    df = df.dropna(subset=["Ref"])
    df = df.rename(columns=rename_columns)
    df_sub = df[selected_cols]
    df_sub = df_sub.astype(dtype_map)
    df_sub["Reviewer Name"] = reviewer_name
    df_sub["GPA 1"] = df_sub["GPA 1"].apply(lambda x: float(gpa_pattern.match(x).group(0)) if gpa_pattern.match(x) else np.nan)
    df_sub = df_sub.set_index("Ref")
    record = df_sub.to_dict("index")
    return record

applicant_records = {}
for file in files:
    if not file.endswith(".xlsx"):
        error_logger.warning(f"{file} not in excel format.")
        continue
    excel_file = DATA_PATH + file
    console_logger.info(excel_file)
    record = process_spreadsheet(excel_file)
    
    for applicant_id, data in record.items():
        applicant_name = data["Name"]  
        
        # convert the columns in reviewer_items to lists
        for reviewer_item in reviewer_items:
            data[reviewer_item] = [data[reviewer_item]]
        reviewer_name = data["Reviewer Name"][0]
        rating = data["Rating"][0]
        # check rating is an int
        try:
            rating = int(rating)
            data["Rating"] = [rating]
        except ValueError:
            error_logger.warning(f"invalid rating, {rating}, for {applicant_name} by {reviewer_name}.")
            continue
            
        if applicant_id not in applicant_records:
            reviewers_needed = 2
            if re.search("high gpa", data["Reader 2 Name"], re.IGNORECASE) is not None or re.search("high gpa", data["Reader 1 Name"], re.IGNORECASE) is not None:
                reviewers_needed = 1
                review_case = "high_gpa"
                recommeneded_action = decision_matrix[review_case][rating]
            elif re.search("2nd rev.*?needed", data["Reader 2 Name"], re.IGNORECASE) is not None or re.search("2nd rev.*?needed", data["Reader 1 Name"], re.IGNORECASE) is not None:
                reviewers_needed = 1
                review_case = "2nd_rev_if_needed"
                recommeneded_action = decision_matrix[review_case][rating]
            elif re.search("missing", data["Reader 2 Name"], re.IGNORECASE) is not None or re.search("missing", data["Reader 1 Name"], re.IGNORECASE) is not None:
                reviewers_needed = 2
                review_case = "missing_2nd_rev"
                recommeneded_action = decision_matrix[review_case][rating]
            else:
                reviewers_needed = 2
                recommeneded_action = "Need 2nd Rev"
            applicant_records[applicant_id] = data
            applicant_records[applicant_id]["reviewer_count"] = 1
            applicant_records[applicant_id]["reviewers_needed"] = reviewers_needed
         
        # second reviewer's rating
        else:
            # no applicant should be assigned to more than two reviewers
            applicant_records[applicant_id]["reviewer_count"] += 1
            if applicant_records[applicant_id]["reviewer_count"] > applicant_records[applicant_id]["reviewers_needed"]:
                error_logger.warning(f"more reviewers than needed are assigned for {applicant_name}.")
                continue
            
            applicant_records[applicant_id]["Reviewer Name"].append(reviewer_name)
            applicant_records[applicant_id]["Rating"].append(rating)
            reviewer1_rating = applicant_records[applicant_id]["Rating"][0]
            recommeneded_action = decision_matrix[reviewer1_rating][rating]
            
        applicant_records[applicant_id]["Recommended Action"] = recommeneded_action
       