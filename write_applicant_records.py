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

from util_vars import (
    decision_matrix,
    dtype_map,
    rename_columns,
    reviewer_items,
    selected_cols,
    necessary_cols,
)

gpa_pattern = re.compile(r"\d.[\d]+")


def assign_scholarship(gpa):
    scholarship = None
    try:
        gpa = float(gpa)
    except (ValueError, TypeError):
        return scholarship
    if gpa >= 4.0:
        pass
    elif gpa > 3.8:
        scholarship = "45%"
    elif gpa > 3.5:
        scholarship = "40%"
    elif gpa > 3.3:
        scholarship = "30%"
    elif gpa > 3.0:
        scholarship = "25%"
    return scholarship


def make_applicant_record(excel_file) -> dict:
    reviewer_name = excel_file.rstrip(".xlsx").split(" - ")[-1]
    df = pd.read_excel(excel_file, header=1)
    df.columns = [
        header.strip() for header in df.columns
    ]  # remove leading and trailing spaces
    df = df.rename(columns=rename_columns)
    df = df.dropna(subset=["Ref"])
    df = df.astype(dtype_map)
    df = df.replace({np.nan: None})  # replace all nan with None
    df = df.astype("str")  # convert all data to string

    existing_cols = set(df.columns)
    for col in selected_cols:
        if col not in existing_cols:
            if col in necessary_cols:
                logging.getLogger("logger").exception(
                    f"{excel_file} not parsed successfully because the reviewer modified the name Ref or Rating header."
                )
                return {}
            logging.getLogger("logger").warning(
                f"{col} is not present in spreadsheet from {reviewer_name}."
            )
            df[col] = None

    df = df.assign(
        **{
            "Reviewer Name": reviewer_name,
            "GPA 1": df["GPA 1"].apply(
                lambda x: float(gpa_pattern.match(x).group(0))
                if gpa_pattern.match(x)
                else np.nan
            ),
        }
    )
    df = df.set_index("Ref")
    return df.to_dict("index")


def make_admission_recommendation(record: dict, applicant_records: dict) -> None:
    for applicant_id, data in record.items():
        applicant_name = data["Name"]
        reviewer_name = data["Reviewer Name"]
        rating = data["Rating"]

        # to show the raw raing from the reviewer
        data["Raw Rating"] = rating
        # check rating is an int
        try:
            rating = int(rating)
        except (ValueError, TypeError):
            logging.getLogger("logger").warning(
                f"invalid rating, {rating}, for {applicant_name} by {reviewer_name}."
            )
            rating = "other"

        data["Rating"] = rating
        if applicant_id not in applicant_records:
            applicant_records[applicant_id] = data
            for reviewer_item in reviewer_items:
                item = applicant_records[applicant_id][reviewer_item]
                applicant_records[applicant_id][reviewer_item] = [item]

            reviewers_needed = 2
            if (
                re.search("only one", data["Reader 2 Name"], re.IGNORECASE) is not None
                or re.search("only one", data["Reader 1 Name"], re.IGNORECASE)
                is not None
            ):
                reviewers_needed = 1
                review_case = "high_gpa"
                recommeneded_action = decision_matrix[review_case][rating]
            elif (
                re.search("2nd rev.*?needed", data["Reader 2 Name"], re.IGNORECASE)
                is not None
                or re.search("2nd rev.*?needed", data["Reader 1 Name"], re.IGNORECASE)
                is not None
            ):
                reviewers_needed = 1
                review_case = "2nd_rev_if_needed"
                recommeneded_action = decision_matrix[review_case][rating]
            elif (
                re.search("missing", data["Reader 2 Name"], re.IGNORECASE) is not None
                or re.search("missing", data["Reader 1 Name"], re.IGNORECASE)
                is not None
            ):
                reviewers_needed = 2
                review_case = "missing_2nd_rev"
                recommeneded_action = decision_matrix[review_case][rating]
            else:
                reviewers_needed = 2
                recommeneded_action = "Need 2nd Rev"

            applicant_records[applicant_id]["Reviewer Count"] = 1
            applicant_records[applicant_id]["Reviewers Needed"] = reviewers_needed

        # second reviewer's rating
        else:
            # no applicant should be assigned to more than two reviewers
            applicant_records[applicant_id]["Reviewer Count"] += 1
            if (
                applicant_records[applicant_id]["Reviewer Count"]
                > applicant_records[applicant_id]["Reviewers Needed"]
            ):
                logging.getLogger("logger").critical(
                    f"more reviewers than needed are assigned for {applicant_name}."
                )
                continue

            for reviewer_item in reviewer_items:
                applicant_records[applicant_id][reviewer_item].append(
                    data[reviewer_item]
                )

            reviewer1_rating = applicant_records[applicant_id]["Rating"][0]
            recommeneded_action = decision_matrix[reviewer1_rating][rating]

        scholarship = None
        if recommeneded_action == "Admit":
            scholarship = assign_scholarship(applicant_records[applicant_id]["GPA 1"])
            if applicant_records[applicant_id]["Data Structures course"] == "No":
                recommeneded_action = "Admit - Summer"
        applicant_records[applicant_id]["Recommended Action"] = recommeneded_action
        applicant_records[applicant_id]["Suggested Scholarship"] = scholarship


def main(data_path=r"./ROUND 1 Reviews"):
    write_path = os.path.join(data_path, "program_data/")
    log_file_path = os.path.join(write_path, "errors.log")
    Path(write_path).mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("logger")
    logging.getLogger().setLevel(logging.INFO)
    logger.propagate = False
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    file_handler = logging.FileHandler(filename=log_file_path, mode="w")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    applicant_records = {}
    excel_files = []
    n = 0
    for file in os.listdir(data_path):
        if not os.path.isfile(os.path.join(data_path, file)):
            continue
        if file.endswith(".xlsx"):
            excel_files.append(file)
        else:
            logging.getLogger("logger").critical(f"{file} not in xlsx format.\n")
            logging.critical(f"{file} not in xlsx format.\n")
        n += 1
    logging.info(f"{len(excel_files)}/{n} files are in xlsx format.\n")

    n = 0
    for file in excel_files:
        excel_file = os.path.join(data_path, file)
        record = make_applicant_record(excel_file)
        if not record:  # if anything went wrong when parsing the excel file, skip
            continue
        logging.info(f"parsing {file}")
        make_admission_recommendation(record, applicant_records)
        n += 1

    logging.info(
        f"{n}/{len(excel_files)} xlsx files read successfully. See errors.log for any errors.\n"
    )
    with open(os.path.join(write_path, "applicant_records.json"), "w") as outfile1:
        json.dump(applicant_records, outfile1, indent=2)

    handlers = logger.handlers
    for handler in handlers:
        logger.removeHandler(handler)
        handler.close()


if __name__ == "__main__":
    main()
