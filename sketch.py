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

DATA_PATH = r"./ROUND 1 Reviews/"

files = [
    'ROUND 1 - CHINA_He_Hangfeng.xlsx',
    'ROUND 1 - CHINA_Li_Dongmei.xlsx',
    'ROUND 1 - CHINA_LUO_Jiebo.xlsx',
    'ROUND 1 - INDIA_Kahng_Anson.xlsx'
]
files = os.listdir(DATA_PATH)
gpa_pattern = re.compile(r'\d.[\d]+')
dtype_map = {"GPA 1": str}

#%% Samuel TODO ###
# refer to the request list at https://docs.google.com/document/d/1AuXrk_64MoJz_aQwfpzxk1XVGMSTGWbgeMT4Rnq2Su4/edit#
# add the columns below that you think we can add. Mark the ones that are not reasonable. 
rename_columns = {
        "On a scale of 1-5, do you think this student will succeed in our curriculum  (see RUBRIC) (1= Deny, 2= additional review needed,  3=waitlist,  4= Accept 5= Strong Accept and increase scholarship)": "Rating",
        "any notes that make this highlight this candidate": "Highlights",
        "Institution 1 GPA (4.0 Scale)": "GPA 1", "Institution 1 Level of Study": "Level of Study 1", "Institution 1 Name": "Institution Name 1", "School 1 Country": "School 1 Country","Institution 1 Degree": "Degree", "Institution 1 Date Conferred": "Date Conferred 1",       "Institution 1 Major": "Major 1", "Citizenship1": "Citizenship1", "Test Date": "Test Date", "Verified":"Verified", 
"TOEFL Total":"TOEFL Total", "TOEFL Speaking":"TOEFL Speaking", "IELTS Test Date":"IELTS Test Date", "IELTS Verified":"IELTS Verified", "IELTS Total": "IELTS Total", "IELTS Speaking":"IELTS Speaking", "GRE Test Date":"GRE Test Date", "GRE Verified":"GRE Verified", "GRE Quantitative": "GRE Quantitative", "GRE Quantitative Percentile": "Quantitative Percentile", "GRE Verbal":"GRE Verbal", "GRE Verbal Percentile":"Verbal Percentile", "GRE Analytical Writing": "Analytical Writing", "GRE Analytical Writing Percentile":"Analytical Writing Percentile", "Sex":"Sex", "Age":"Age", "Institution 2 Level of Study": "Level of Study 2","School 2 Country": "School 2 Country", "Institution 2 Name": "Institution Name 2", "Institution 2 Major":"Major 2", "Institution 2 Date Conferred": "Date Conferred 2", "Institution 2 GPA (4.0 Scale)": "GPA 2"
}


# these columns will be lists
reviewer_items = ["Reviewer Name", "Rating", "Highlights"]
selected_cols = [
    "Ref",
    "Rating",
    "Name", 
    "GPA 1", 
    "Reader 2 Name",
    "Reader 1 Name",
    "Highlights",
    "Level of Study 1",
    "Institution Name 1",
    "School 1 Country",
    "Degree 1", 
    "Date Conferred 1",
    "Major 1",
    "Citizenship1",
    "Test Date", 
    "Verified",
    "TOEFL Total", 
    "TOEFL Speaking",
    "IELTS Test Date",
    "IELTS Verified", 
    "IELTS Total", 
    "IELTS Speaking", 
    "GRE Test Date", 
    "GRE Verified",
    "GRE Quantitative", 
    "Quantitative Percentile", 
    "GRE Verbal", 
    "Verbal Percentile", 
    "Analytical Writing", 
    "Analytical Writing Percentile", 
    "Sex", 
    "Age", 
    "Level of Study 2",
    "School 2 Country", 
    "Institution Name 2",
    "Major 2", 
    "Date Conferred 2", 
    "GPA 2"
]

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

#%%

applicant_records = {}
for file in files:
    if not file.endswith(".xlsx"):
        print(f"{file} not in excel format.")
        continue
    excel_file = DATA_PATH + file
    print(excel_file)
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
            print(f"invalid rating, {rating}, for {applicant_name} by {reviewer_name}.")
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
                print(f"more reviewers than needed are assigned for {applicant_name}.")
                continue
            
            applicant_records[applicant_id]["Reviewer Name"].append(reviewer_name)
            applicant_records[applicant_id]["Rating"].append(rating)
            reviewer1_rating = applicant_records[applicant_id]["Rating"][0]
            recommeneded_action = decision_matrix[reviewer1_rating][rating]
            
        applicant_records[applicant_id]["Recommended Action"] = recommeneded_action
        