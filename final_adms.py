#!/usr/bin/env python
# coding: utf-8

import os
import pandas as pd
import numpy
import re


def read_reviewer_folder1(folder_path):
    all_files = os.listdir(folder_path)
    file_list = []
    for i, file_name in enumerate(all_files):
        if file_name.endswith('.xlsx'):
            try:
                file_path = os.path.join(folder_path, file_name)
                r1 = pd.read_excel(file_path, header=0)
                r2 = pd.read_excel(file_path, header=1)
                if 'Rating' in r1.columns and 'Name' in r2.columns and 'Institution 1 GPA (4.0 Scale)' in r2.columns:
                    rating = r1['Rating'][1:].to_numpy()
                    name = r2['Name'].to_numpy()
                    reader2 = r2['Reader 2 Name'].astype(str)
                    splits = file_name.split(' ')
                    round_num = splits[1].split('_')[0] if len(splits) > 1 else ""
                    verd = {'Applicant_Name': name,'Reader 2': reader2, f'Rating': rating}
                    verdict = pd.DataFrame.from_dict(verd)
                    reviewer_name = re.search(r'_(.+?)\.xlsx', file_name).group(1)
                    verdict[f'Reviewer_Name_{i+1}'] = reviewer_name
                    file_list.append(verdict)
                else:
                    print(f"File {file_name} does not contain expected columns.")
            except Exception as e:
                print(f"Error reading file {file_name}: {str(e)}")

    return file_list

def merge_reviewer_ratings1(file_list):
    # Combine the DataFrames on the Name column
    merged_df = pd.concat(file_list).groupby('Applicant_Name', as_index=False).first()

    # Create a new DataFrame with just the Name column
    name_df = merged_df[['Applicant_Name', 'Reader 2']].copy()

    # Loop through the columns for each reviewer's ratings
    for i, df in enumerate(file_list):
        # Rename the Rating column to include the reviewer's name
        df.rename(columns={'Rating': f'reviewer {i+1} Rating'}, inplace=True)

        # Merge the ratings for this reviewer onto the name DataFrame
        name_df = pd.merge(name_df, df[['Applicant_Name', f'reviewer {i+1} Rating', f"Reviewer_Name_{i+1}"]], on='Applicant_Name', how='left')

    return name_df

def make_decision_matrix2(decision_df):
    # Initialize decision matrix
    decision_matrix = pd.DataFrame(columns=['Applicant_Name', "Ratings", "Reviewers", 'Decision'])

    # Loop through each applicant in the decision dataframe
    for index, row in decision_df.iterrows():
        name = row['Applicant_Name']
        ratings = []
        reviewers = []
        for i, col in enumerate(decision_df.columns):
            if 'Rating' in col:
                rating = row[col]
                if not pd.isna(rating):
                    reviewers.append(row[i+1])
                    ratings.append(rating)

        # Get reader 2 info 
        reader2 = row['Reader 2']
        
        # Check decision conditions
        if not ratings:
            decision = 'Missing reviewers'
        elif len(ratings) == 1:
            decision = 'Missing 2nd reviewer'
            if re.search("high gpa", reader2, re.IGNORECASE) is not None:
                if ratings[0] >= 4:
                    decision = 'ADMIT'
                elif ratings[0] == 3:
                    decision = 'WAITLIST - HIGH'
                elif ratings[0] == 2:
                    decision = 'LOOK AGAIN'
                elif ratings[0] == 1:
                    decision = 'DISCUSS'
                else:
                    decision = 'DENY'
        elif len(ratings) == 2:
            ratings.sort()
            rating1, rating2 = ratings
            if rating1 == rating2:
                if rating1 == 1:
                    decision = 'DENY'
                elif rating1 == 2:
                    decision = 'LOOK AGAIN'
                elif rating1 == 3:
                    decision = 'WAITLIST - LOW'
                elif rating1 == 4 or rating1 == 5:
                    decision = 'ADMIT'
            elif rating1 == 1 and rating2 == 2:
                decision = 'LOOK AGAIN'
            elif rating1 == 1 and rating2 == 3:
                decision = 'DISCUSS'
            elif rating1 == 1 and rating2 == 4:
                decision = 'DISCUSS'
            elif rating1 == 1 and rating2 == 5:
                decision = 'DISCUSS'
            elif rating1 == 2 and rating2 == 3:
                decision = 'LOOK AGAIN'
            elif rating1 == 2 and rating2 == 4:
                decision = 'LOOK AGAIN'
            elif rating1 == 2 and rating2 == 5:
                decision = 'LOOK AGAIN'
            elif rating1 == 3 and rating2 == 4:
                decision = 'WAITLIST - HIGH'
            elif rating1 == 3 and rating2 == 5:
                decision = 'DISCUSS'
            elif rating1 == 4 and rating2 == 5:
                decision = 'ADMIT'
            else:
                decision = 'DENY'
        ratings = ", ".join(str(rating) for rating in ratings)
        reviewers = ", ".join(reviewers)
        # Add decision to decision matrix
        decision_matrix = pd.concat([decision_matrix, pd.DataFrame({'Applicant_Name': [name], "Ratings": ratings, "Reviewers": reviewers, 'Decision': [decision]})], ignore_index=True)

    return decision_matrix


# In[28]:
f_p = r"./China Example/"
output_path = r"./China Example/final_decision.csv"
if not output_path:
    output = f_p

reviews = read_reviewer_folder1(folder_path=f_p)
merge_data = merge_reviewer_ratings1(reviews)
applicant_dec = make_decision_matrix2(merge_data)
applicant_dec.to_csv(output_path)


