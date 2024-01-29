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
    worksheet.set_column(2, 2, 6)
    worksheet.set_column(3, 3, 15)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 5, 10)
    worksheet.set_column(6, 6, 10)
    worksheet.set_column(7, 7, 10) 
    worksheet.set_column(8, 8, 14)

    header_format = workbook.add_format({"border": 2, "bold": True})
    gpa_format = workbook.add_format({"bg_color": "red"})

    #color coding format
    red_format = workbook.add_format({"bg_color": "red"})
    orange_format = workbook.add_format({"bg_color": "orange"})
    yellow_format = workbook.add_format({"bg_color": "yellow"})
    green_format = workbook.add_format({"bg_color": "green"})
    cyan_format = workbook.add_format({"bg_color": "cyan"})

    row = 0
    col = 0
    worksheet.write(row, col, "Ref", header_format)
    col += 1

    # Write headers
    for header, renamed_header in headers1.items():
        if header == "Reviewer Name":
            worksheet.write(row, col, "Reviewer 1 Name", header_format)
            col += 1
            worksheet.write(row, col, "Reviewer 2 Name", header_format)
            col += 1
            worksheet.write(row, col, "Reviewer 1 Rating", header_format)
            col += 1
            worksheet.write(row, col, "Reviewer 2 Rating", header_format)
            col += 1
            continue
        worksheet.write(row, col, renamed_header, header_format)
        col += 1


    #worksheet.write(row, col, "Gap", header_format)

    #col += 1
    for header in selected_cols:
        if header in headers1:
            continue
        worksheet.write(row, col, header, header_format)
        col += 1

    # Fill in the values
    row += 1
    for ref, info in applicant_records.items():
        col = 0
        worksheet.write(int(row), col, ref, header_format)
        col += 1

        for header, renamed_header in headers1.items():
            if header == "Reviewer Name":
                # Write values for "Reviewer 1 Name" and "Reviewer 2 Name"
                reviewer_names = info.get(header, ["", ""])

                worksheet.write(int(row), col, reviewer_names[0].split("_")[1] if "_" in reviewer_names[0] else reviewer_names[0])
                col += 1
                worksheet.write(int(row), col, reviewer_names[1].split("_")[1] if len(reviewer_names) > 1 and "_" in reviewer_names[1] else reviewer_names[1] if len(reviewer_names) > 1 else "")
                col += 1

                #write values for "Reviewer 1 Rating" and "Reviewer 2 Rating"
                rating_values = info.get("Rating", ["", ""])
                worksheet.write(int(row), col, rating_values[0])

                #add conditional formatting for "Reviewer 1 Rating"
                worksheet.conditional_format(
                    f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": 1,
                        "format": red_format,
                    },
                )
                worksheet.conditional_format(
                    f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": 2,
                        "format": orange_format,
                    },
                )
                worksheet.conditional_format(
                    f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": 3,
                        "format": yellow_format,
                    },
                )
                worksheet.conditional_format(
                    f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": 4,
                        "format": green_format,
                    },
                )
                worksheet.conditional_format(
                    f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": 5,
                        "format": cyan_format,
                    },
                )

                col += 1
                if len(rating_values) > 1:
                    worksheet.write(int(row), col, rating_values[1])

                    #conditional formatting for "Reviewer 2 Rating"
                    worksheet.conditional_format(
                        f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                        {
                            "type": "cell",
                            "criteria": "equal to",
                            "value": 1,
                            "format": red_format,
                        },
                    )
                    worksheet.conditional_format(
                        f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                        {
                            "type": "cell",
                            "criteria": "equal to",
                            "value": 2,
                            "format": orange_format,
                        },
                    )
                    worksheet.conditional_format(
                        f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                        {
                            "type": "cell",
                            "criteria": "equal to",
                            "value": 3,
                            "format": yellow_format,
                        },
                    )
                    worksheet.conditional_format(
                        f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                        {
                            "type": "cell",
                            "criteria": "equal to",
                            "value": 4,
                            "format": green_format,
                        },
                    )
                    worksheet.conditional_format(
                        f"{xlsxwriter.utility.xl_col_to_name(col)}{row}",
                        {
                            "type": "cell",
                            "criteria": "equal to",
                            "value": 5,
                            "format": cyan_format,
                        },
                    )

                col += 1

                continue

            values = info.get(header, "")  # Use get() to avoid KeyError
            if isinstance(values, list):
                values = "; ".join(str(value) for value in values)

            #apply color coding based on recommended action
            if header == "Recommended Action":
                if values in ["Discuss", "Look Again", "Need 2nd Rev"]:
                    worksheet.write(int(row), col, values, orange_format)
                elif values == "Deny":
                    worksheet.write(int(row), col, values, red_format)
                elif values in ["Admit", "Admit - Summer"]:
                    worksheet.write(int(row), col, values, green_format)
                elif values in ["Wait - High", "Wait - Low"]:
                    worksheet.write(int(row), col, values, yellow_format)
                else:
                    worksheet.write(int(row), col, values)
            else:
                worksheet.write(int(row), col, values)

            col += 1

        #col += 1  # Mind the gap!

        for field in selected_cols:
            if field not in headers1:
                values = info[field]
                if isinstance(values, list):
                    values = "; ".join(str(value) for value in values)
                try:
                    worksheet.write(int(row), col, values)  # Convert row to int
                except TypeError:
                    pass
                col += 1
        row += 1

    #color GPA greater than 4 red
    worksheet.conditional_format(
        f"M2:M{row}",
        {
            "type": "cell",
            "criteria": "greater than or equal to",
            "value": 4,
            "format": gpa_format,
        },
    ) 

    #GPA less than 3.0
    worksheet.conditional_format(
        f"M2:M{row}",
        {
            "type": "formula",
            "criteria": 'AND(VALUE(M2) < 3, M2<>"")',
            "format": red_format,
        },
    ) 

    #correcting the bug in color coding in "Reviwer 2 Column"
    worksheet.conditional_format(
        f"G2:G{row}",
        {
            "type": "cell",
            "criteria": "equal to",
            "value": 1,
            "format": red_format,
        },
    ) 

    #correcting the bug in color coding in "Reviwer 2 Column"
    worksheet.conditional_format(
        f"G2:G{row}",
        {
            "type": "cell",
            "criteria": "equal to",
            "value": 2,
            "format": orange_format,
        },
    ) 

    #correcting the bug in color coding in "Reviwer 2 Column"
    worksheet.conditional_format(
        f"G2:G{row}",
        {
            "type": "cell",
            "criteria": "equal to",
            "value": 3,
            "format": yellow_format,
        },
    ) 

    #correcting the bug in color coding in "Reviwer 2 Column"
    worksheet.conditional_format(
        f"G2:G{row}",
        {
            "type": "cell",
            "criteria": "equal to",
            "value": 4,
            "format": green_format,
        },
    ) 

    #correcting the bug in color coding in "Reviwer 2 Column"
    worksheet.conditional_format(
        f"G2:G{row}",
        {
            "type": "cell",
            "criteria": "equal to",
            "value": 5,
            "format": cyan_format,
        },
    ) 

    #highlighing 'no' if a student didn't take data structures (col L)
    worksheet.conditional_format(
        f"L2:L{row}",
        {
            "type": "text",
            "criteria": "containing",
            "value": "No",
            "format": red_format,
        },
    ) 

    #TOEFL greater than 5
    worksheet.conditional_format(
        f"AJ2:AJ{row}",
        {
            "type": "formula",
            "criteria": 'AND(VALUE(AJ2) < 100, AJ2<>"")',
            "format": red_format,
        },
    )

    #IETLS less than 7
    worksheet.conditional_format(
        f"AN2:AN{row}",
        {
            "type": "formula",
            "criteria": 'AND(VALUE(AN2) < 7.0, AN2<>"")',
            "format": red_format,
        },
    )

    #color code cells in column K that are not No Fraud and are not blank
    worksheet.conditional_format(
        f"K2:K{row}",
        {
            "type": "formula",
            "criteria": 'AND(K2<>"", K2<>"No Fraud")',
            "format": red_format,
        },
    )

    #highlighing 'university of rochester'
    worksheet.conditional_format(
        f"AB2:AB{row}",
        {
            "type": "text",
            "criteria": "containing",
            "value": "University of Rochester",
            "format": green_format,
        },
    ) 

    workbook.close()
    logging.info("Done.")


if __name__ == "__main__":
    main()


if __name__ == "__main__":
    main()
