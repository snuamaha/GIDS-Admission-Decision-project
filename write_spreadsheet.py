# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 15:04:37 2023

@author: Yue
"""

import os
import json
import logging

import xlsxwriter

from util_vars import selected_cols


headers1 = {
    "Name": "Applicant Name",
    "Raw Rating": "Ratings",
    "Reviewer Name": "Reviewer Names",
    "Highlights": "Highlights",
    "Recommended Action": "Recommended Action",
    "Suggested Scholarship": "Suggested Scholarship",
}

dates = set(
    [
        "Date Conferred 1",
        "Test Date",
        "IELTS Test Date",
        "GRE Test Date",
        "Date Conferred 2",
    ]
)

floats = set(
    [
        "GPA 1",
        "Math Proficiency",
        "Stats Proficiency",
        "Programming Proficiency",
        "Data Structure Proficiency",
        "Communications Skills",
        "Applied data science skills",
        "References",
        "TOEFL Total",
    ]
)


def main(data_path=r"./ROUND 1 Reviews/"):
    program_data_path = os.path.join(data_path, "program_data/")
    write_path = program_data_path
    record_file = os.path.join(program_data_path, "applicant_records.json")

    with open(record_file, "r") as infile1:
        applicant_records = json.load(infile1)

    selected_cols.remove("Ref")
    selected_cols.remove("Rating")
    logging.info("Preparing the admission decision spreadsheet...")
    workbook = xlsxwriter.Workbook(
        os.path.join(write_path, "admission_recommendations.xlsx")
    )
    worksheet = workbook.add_worksheet()

    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(3, 3, 20)
    worksheet.set_column(4, 4, 10)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 19)
    worksheet.set_column(7, 7, 10)  # Gap

    worksheet.set_column(8, 8, 12)

    header_format = workbook.add_format({"border": 2, "bold": True})
    deny_format = workbook.add_format({"bg_color": "red"})
    admit_format = workbook.add_format({"bg_color": "green"})
    wait_format = workbook.add_format({"bg_color": "yellow"})
    fraud_format = workbook.add_format({"bg_color": "red"})
    gpa_format = workbook.add_format({"bg_color": "red"})

    row = 0
    col = 0
    worksheet.write(row, col, "Ref", header_format)
    col += 1
    # write headers
    for header, renamed_header in headers1.items():
        worksheet.write(row, col, renamed_header, header_format)
        col += 1
    worksheet.write(row, col, "Gap", header_format)

    col += 1
    for header in selected_cols:
        if header in headers1:
            continue
        worksheet.write(row, col, header, header_format)
        col += 1

    # fill in the values
    row += 1
    for ref, info in applicant_records.items():
        col = 0
        worksheet.write(row, col, ref, header_format)
        col += 1
        for header, renamed_header in headers1.items():
            values = info[header]
            if isinstance(values, list):
                values = "; ".join(str(value) for value in values)
            worksheet.write(row, col, values)
            col += 1
        col += 1  # mind the gap!

        for field in selected_cols:
            if field not in headers1:
                values = info[field]
                if isinstance(values, list):
                    values = "; ".join(str(value) for value in values)
                try:
                    worksheet.write(row, col, values)
                except TypeError:
                    pass
                col += 1
        row += 1

    worksheet.conditional_format(
        f"F1:F{row}",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Deny",
            "format": deny_format,
        },
    )

    worksheet.conditional_format(
        f"F1:F{row}",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Admit",
            "format": admit_format,
        },
    )

    worksheet.conditional_format(
        f"F1:F{row}",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Wait",
            "format": wait_format,
        },
    )

    worksheet.conditional_format(
        f"M1:M{row}",
        {
            "type": "no_blanks",
            "format": fraud_format,
        },
    )

    worksheet.conditional_format(
        f"J1:J{row}",
        {
            "type": "cell",
            "criteria": "greater than or equal to",
            "value": 4,
            "format": gpa_format,
        },
    )

    workbook.close()
    logging.info("Done.")


if __name__ == "__main__":
    main()
