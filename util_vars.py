# -*- coding: utf-8 -*-
"""
Created on Sun Feb 26 18:26:54 2023

@author: Yue
"""


decision_matrix = {
    "high_gpa": {1: "Discuss", 2: "Look Again", 3: "Wait - High", 4: "Admit", 5: "Admit"},
    "2nd_rev_if_needed": {1: "Deny", 2: "Look Again", 3: "Wait - Low", 4: "Send to 2nd Rev", 5: "Send to 2nd Rev"},
    "missing_2nd_rev": {1: "Need 2nd Rev", 2: "Need 2nd Rev", 3: "Need 2nd Rev", 4: "Need 2nd Rev", 5: "Need 2nd Rev"},
    1: {1: "Deny", 2: "Look Again", 3: "Discuss", 4: "Discuss", 5: "Discuss"},
    2: {1: "Look Again", 2: "Look Again", 3: "Look Again", 4: "Wait - High", 5: "Discuss"},
    3: {1: "Discuss", 2: "Look Again", 3: "Wait - Low", 4: "Discuss", 5: "Discuss"},
    4: {1: "Discuss", 2: "Look Again", 3: "Wait - High", 4: "Admit", 5: "Admit"},
    5: {1: "Discuss", 2: "Look Again", 3: "Discuss", 4: "Admit", 5: "Admit"},
}

#%% Samuel Done ###
# refer to the request list at https://docs.google.com/document/d/1AuXrk_64MoJz_aQwfpzxk1XVGMSTGWbgeMT4Rnq2Su4/edit#
# add the columns below that you think we can add. Mark the ones that are not reasonable. 
rename_columns = {
    "On a scale of 1-5, do you think this student will succeed in our curriculum  (see RUBRIC) (1= Deny, 2= additional review needed,  3=waitlist,  4= Accept 5= Strong Accept and increase scholarship)": "Rating",
    "any notes that make this highlight this candidate": "Highlights",
    "Institution 1 GPA (4.0 Scale)": "GPA 1",
    "Institution 1 Level of Study": "Level of Study 1", 
    "Institution 1 Name": "Institution Name 1", 
    "School 1 Country": "School 1 Country",
    "Institution 1 Degree": "Degree 1", 
    "Institution 1 Date Conferred": "Date Conferred 1",       
    "Institution 1 Major": "Major 1", 
    "Citizenship1": "Citizenship1", 
    "Test Date": "Test Date", 
    "Verified":"Verified", 
    "TOEFL Total":"TOEFL Total", 
    "TOEFL Speaking":"TOEFL Speaking", 
    "IELTS Test Date":"IELTS Test Date", 
    "IELTS Verified":"IELTS Verified", 
    "IELTS Total": "IELTS Total", 
    "IELTS Speaking":"IELTS Speaking", 
    "GRE Test Date":"GRE Test Date", 
    "GRE Verified":"GRE Verified", 
    "GRE Quantitative": "GRE Quantitative", 
    "GRE Quantitative Percentile": "Quantitative Percentile", 
    "GRE Verbal":"GRE Verbal", 
    "GRE Verbal Percentile":"Verbal Percentile", 
    "GRE Analytical Writing": "Analytical Writing", 
    "GRE Analytical Writing Percentile":"Analytical Writing Percentile", 
    "Sex":"Sex", 
    "Age":"Age", 
    "Institution 2 Level of Study": "Level of Study 2",
    "School 2 Country": "School 2 Country", 
    "Institution 2 Name": "Institution Name 2", 
    "Institution 2 Major":"Major 2", 
    "Institution 2 Date Conferred": "Date Conferred 2", 
    "Institution 2 GPA (4.0 Scale)": "GPA 2"
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

#%%