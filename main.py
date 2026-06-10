import re
import unicodedata
import os
import codecs
import random
from datetime import datetime
import json
import csv
import openpyxl

def assessment_json_generation(sheet_id="assessment_sheet/Urdu_assessment_Worksheet.xlsx", lang="Urdu", tab_number=0):
    assessment_type=get_assessment_bucket_title(sheet_id,tab_number)
    if "words" in assessment_type:
        assessment_type= "sight-words"
    else:
        assessment_type= "letter-sounds"
        
    assessment_content= get_assessment_bucket(sheet_id,tab_number)
    content_version=get_and_update_content_version(sheet_id,tab_number)
    
    json_content = create_json_from_data(assessment_content,lang,assessment_type,content_version)
    return json_content

def get_assessment_bucket_title(sheet_id,tab_number):
    wb = openpyxl.load_workbook(sheet_id, data_only=True)
    worksheet = wb.worksheets[tab_number]
    tab_title = worksheet.title
    tab_name=re.sub(r"[ -]", "", tab_title)
    return tab_name

def get_and_update_content_version(sheet_id, tab_number):
    wb = openpyxl.load_workbook(sheet_id)
    worksheet = wb.worksheets[tab_number]
    
    content_version = worksheet["E1"].value
    if not content_version:
        content_version = "v0.0"
    if isinstance(content_version, str) and content_version.startswith('v'):
        numeric_version = content_version[1:]
        try:
            current_version = float(numeric_version)
            new_version = current_version + 0.1
            new_version_str = f"v{new_version:.1f}"
            worksheet["E1"] = new_version_str
            wb.save(sheet_id)
            return new_version_str
        except ValueError:
            return "Failed to parse the current version. Ensure it is in the format 'vX.Y'."
    else:
        return "Invalid or missing version format in cell E1."

def get_assessment_bucket(sheet_id,tab_number):
    wb = openpyxl.load_workbook(sheet_id, data_only=True)
    worksheet = wb.worksheets[tab_number]
    fetched_content = []
    for row in worksheet.iter_rows(min_row=2, max_row=151, min_col=1, max_col=2, values_only=True):
        if row[0] is not None or row[1] is not None:
            fetched_content.append([str(c) if c is not None else "" for c in row])
    return fetched_content

def create_json_from_data(data,lang,assessment_type,content_version):
    bucket_name=lang.replace(" ","-")+"let-b"
    json_data = {
        "quizName": lang+" "+assessment_type.replace("-"," "),
        "appType": "assessment",
        "assessmentType": assessment_type,
        "feedbackText": "fantastic!",
        "contentVersion": content_version,
        "buckets": []
    }

    buckets = {}
    for row in data:
        if len(row) < 2:
            continue  
        
        bucketID = unicodedata.normalize('NFC', str(row[0]))
        itemName = unicodedata.normalize('NFC', str(row[1]))
        itemText = unicodedata.normalize('NFC', str(row[1]))

        if bucketID not in buckets:
            try:
                b_id_int = int(float(bucketID))
            except ValueError:
                continue
            
            buckets[bucketID] = {
                "bucketID": b_id_int,
                "bucketName": f"{bucket_name}-{b_id_int}",
                "usedItems": [],
                "items": []
            }

        buckets[bucketID]["items"].append({
            "itemName": itemName,
            "itemText": itemText
        })

    json_data["buckets"] = list(buckets.values())
    return json.dumps(json_data, indent=2)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate assessment content JSON from Excel.")
    parser.add_argument("--sheet_id", type=str, default="assessment_sheet/Literacy Assessment Worksheet _ English.xlsx", help="Path to the Excel sheet")
    parser.add_argument("--lang", type=str, default="English", help="Language for the assessment")
    parser.add_argument("--tab", type=int, default=1, help="Tab index")
    
    args = parser.parse_args()

    json_content = assessment_json_generation(
        sheet_id=args.sheet_id,
        lang=args.lang,
        tab_number=args.tab
    )
    
    with open("content.json", "w", encoding="utf-8") as f:
        f.write(json_content)
    print("content.json has been generated successfully.")