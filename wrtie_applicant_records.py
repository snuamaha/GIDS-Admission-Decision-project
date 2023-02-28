# -*- coding: utf-8 -*-
"""
Created on Mon Feb 27 09:26:01 2023

@author: Yue
"""

import os
import re
import json
import logging
from pathlib import Path

import numpy as np
import pandas as pd

from util_vars import decision_matrix, rename_columns, reviewer_items, selected_cols


DATA_PATH = r"./ROUND 1 Reviews/"
WRITE_PATH = DATA_PATH + "program_data/"
log_file_path = WRITE_PATH + "errors.log"
Path(WRITE_PATH).mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("logger")
logging.getLogger().setLevel(logging.INFO)
logger.propagate = False
formatter = logging.Formatter("%(levelname)s - %(message)s")
file_handler = logging.FileHandler(filename=log_file_path, mode="w")
file_handler.setFormatter(formatter)
#file_handler.setLevel(logging.WARNING)
logger.addHandler(file_handler)

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
    reviewer_name = " ".join(excel_file.rstrip(".xlsx").split("_")[1:])
    df = pd.read_excel(excel_file, header=1)
    df = df.astype("string") # convert all data to string
    df = df.fillna(np.nan).replace([np.nan], [None]) # replace all nan with None
    df.columns = [header.strip() for header in df.columns] # remove leading and trailing spaces
    df = df.dropna(subset=["Ref"])
    df = df.rename(columns=rename_columns)
    try:
        df_sub = df[selected_cols]
    except KeyError:
        logging.getLogger("logger").exception(f"{excel_file} not parsed successfully, maybe because the reviewer modified some column names.")
        return {}
    df_sub = df_sub.astype(dtype_map)
    df_sub["Reviewer Name"] = reviewer_name
    df_sub["GPA 1"] = df_sub["GPA 1"].apply(lambda x: float(gpa_pattern.match(x).group(0)) if gpa_pattern.match(x) else np.nan)
    df_sub = df_sub.set_index("Ref")
    return df_sub.to_dict("index")

applicant_records = {}
for file in files:
    if not file.endswith(".xlsx"):
        #print(file)
        logging.getLogger("logger").critical(f"{file} not in excel format.\n")
        continue
    excel_file = DATA_PATH + file
    logging.info(file)
    #print(excel_file)
    record = process_spreadsheet(excel_file)
    if not record: # if anything went wrong when parsing the excel file, skip
        continue
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
        except (ValueError, TypeError):
            logging.getLogger("logger").warning(f"invalid rating, {rating}, for {applicant_name} by {reviewer_name}.")
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
            applicant_records[applicant_id]["Reviewer Count"] = 1
            applicant_records[applicant_id]["Reviewers Needed"] = reviewers_needed
         
        # second reviewer's rating
        else:
            # no applicant should be assigned to more than two reviewers
            applicant_records[applicant_id]["Reviewer Count"] += 1
            if applicant_records[applicant_id]["Reviewer Count"] > applicant_records[applicant_id]["Reviewers Needed"]:
                logging.getLogger("logger").critical(f"more reviewers than needed are assigned for {applicant_name}.")
                continue
            
            applicant_records[applicant_id]["Reviewer Name"].append(reviewer_name)
            applicant_records[applicant_id]["Rating"].append(rating)
            reviewer1_rating = applicant_records[applicant_id]["Rating"][0]
            recommeneded_action = decision_matrix[reviewer1_rating][rating]
            
        applicant_records[applicant_id]["Recommended Action"] = recommeneded_action
    
with open(WRITE_PATH + "applicant_records.json", "w") as outfile1:
    json.dump(applicant_records, outfile1, indent=2)

handlers = logger.handlers
for handler in handlers:
    logger.removeHandler(handler)
    handler.close()
       