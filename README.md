# Assessment JSON Generation

This project contains a simple Python automation script that generates an assessment `content.json` file from an Excel spreadsheet.

## Prerequisites
Make sure you have the required dependencies installed:
```bash
pip install -r requirements.txt
```

## Running the Script

To generate the `content.json` file, you can run the `main.py` script:

```bash
python main.py
```

This processes the default Excel file (`assessment_sheet/Literacy Assessment Worksheet _ English.xlsx`). A `content.json` file will be generated in the same directory.

## Parameters

You can pass arguments to customize the process:
- `--sheet_id`: Path to the Excel file (default: `assessment_sheet/Literacy Assessment Worksheet _ English.xlsx`).
- `--lang`: The language of the assessment (default: `English`).
- `--tab`: The index of the tab in the spreadsheet (default: `1`).

Example:
```bash
python main.py --lang "Spanish" --tab 2
```