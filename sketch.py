# -*- coding: utf-8 -*-
"""
Created on Mon Feb 27 09:26:01 2023

@author: Yue
"""

import os
import re

import numpy as np
import pandas as pd

from decision_matrix import decision_matrix

DATA_PATH = r"./China Example/"

files = [
    'ROUND 1 - CHINA_He_Hangfeng.xlsx',
    'ROUND 1 - CHINA_Li_Dongmei.xlsx',
    'ROUND 1 - CHINA_LUO_Jiebo.xlsx'
]

gpa_pattern = re.compile(r'\d.[\d]+')
dtype_map = {"GPA 1": str}

#%% Samuel TODO ###
# refer to the request list at https://docs.google.com/document/d/1AuXrk_64MoJz_aQwfpzxk1XVGMSTGWbgeMT4Rnq2Su4/edit#
# add the columns below that you think we can add. Mark the ones that are not reasonable. 
rename_columns = {
        "On a scale of 1-5, do you think this student will succeed in our curriculum  (see RUBRIC) (1= Deny, 2= additional review needed,  3=waitlist,  4= Accept 5= Strong Accept and increase scholarship)": "Rating",
        "any notes that make this highlight this candidate": "Highlights",
        "Institution 1 GPA (4.0 Scale)": "GPA 1",
}


# these columns will be lists
reviewer_items = ["Reviewer Name", "Rating", "Highlights"]
selected_cols = [
    "Rating",
    "Name", 
    "GPA 1", 
    "Reader 2 Name",
    "Reader 1 Name",
    "Highlights",
]

#%%

applicant_records = {}
for file in files:
    excel_file = DATA_PATH + file
    reviewer_name = " ".join(excel_file.rstrip(".xlsx").split("_")[-2:])
    df = pd.read_excel(excel_file, header=1)
    df = df.rename(columns=rename_columns)
    df_sub = df[selected_cols]
    df_sub = df_sub.astype(dtype_map)
    df_sub["Reviewer Name"] = reviewer_name
    df_sub["GPA 1"] = df_sub["GPA 1"].apply(lambda x: float(gpa_pattern.match(x).group(0)) if gpa_pattern.match(x) else np.nan)
    df_sub = df_sub.set_index(["Name", "GPA 1"])
    record = df_sub.to_dict("index")
    for applicant, data in record.items():
        applicant_name, gpa = applicant
        
        # convert the columns in reviewer_items to lists
        for reviewer_item in reviewer_items:
            data[reviewer_item] = [data[reviewer_item]]
        reviewer_name = data["Reviewer Name"][0]
        # check rating is an int
        try:
            rating = int(data["Rating"][0])
        except ValueError:
            print(f"invalid rating for {applicant_name} by {reviewer_name}.")
            continue
            
        if applicant not in applicant_records:
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
            applicant_records[applicant] = data
            applicant_records[applicant]["reviewer_count"] = 1
            applicant_records[applicant]["reviewers_needed"] = reviewers_needed
         
        # second reviewer's rating
        else:
            # no applicant should be assigned to more than two reviewers
            applicant_records[applicant]["reviewer_count"] += 1
            if applicant_records[applicant]["reviewer_count"] > applicant_records[applicant]["reviewers_needed"]:
                print(f"more reviewers than needed are assigned for {applicant_name}.")
                continue
            
            applicant_records[applicant]["Reviewer Name"].append(reviewer_name)
            applicant_records[applicant]["Rating"].append(rating)
            reviewer1_rating = applicant_records[applicant]["Rating"][0]
            recommeneded_action = decision_matrix[reviewer1_rating][rating]
            
        applicant_records[applicant]["Recommended Action"] = recommeneded_action
        

ratings = []
reader2_comment = ""