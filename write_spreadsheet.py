# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 15:04:37 2023

@author: Yue
"""

import json

import xlsxwriter

from util_vars import selected_cols

selected_cols.remove("Ref")
DATA_PATH = r"./ROUND 1 Reviews/"
PROGRAM_DATA_PATH = DATA_PATH + "program_data/"
write_path = PROGRAM_DATA_PATH

with open(PROGRAM_DATA_PATH + "applicant_records.json", "r") as infile1:
    applicant_records = json.load(infile1)
    
app = applicant_records["439935221"] 

headers1 = {
    "Name": "Applicant Name",
    "Rating": "Ratings",
    "Reviewer Name": "Reviewer Names",
    "Highlights": "Highlights",
    "Recommended Action": "Recommended Action",
}

dates = set(["Date Conferred 1",
             "Test Date",
             "IELTS Test Date",
             "GRE Test Date",
             "Date Conferred 2",
             ])

floats = set(["GPA 1",
              "Math Proficiency",
              "Stats Proficiency",
              "Programming Proficiency",
              "Data Structure Proficiency",
              "Communications Skills",
              "Applied data science skills",
              "References",
              "TOEFL Total",
               ])


workbook = xlsxwriter.Workbook(write_path + 'admission_recommendations.xlsx')
worksheet = workbook.add_worksheet()

format1 = workbook.add_format({'bg_color': 'red'})

row = 0
col = 0
worksheet.write(row, col, "Ref")
col += 1
for header, renamed_header in headers1.items():
    worksheet.write(row, col, renamed_header)
    col += 1
    
col += 4
for header in selected_cols:
    if header in headers1:
        continue
    worksheet.write(row, col, header)
    col += 1    
    
row += 1
for ref, info in applicant_records.items():
    col = 0
    worksheet.write(row, col, ref)
    col += 1
    for header, renamed_header in headers1.items():
        values = info[header]
        if isinstance(values, list):
            values = ", ".join(str(value) for value in values)
        worksheet.write(row, col, values)
        col += 1
    col += 4
    for field in selected_cols:
        if field not in headers1:
            values = info[field]
            if isinstance(values, list):
                values = ", ".join(str(value) for value in values)
            try:    
                worksheet.write(row, col, values)
            except TypeError:
                pass
            col += 1
    
    row += 1

worksheet.conditional_format('F1:F203', {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'Deny',
                                       'format':   format1})
workbook.close()   

    