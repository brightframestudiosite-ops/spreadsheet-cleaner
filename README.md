# Spreadsheet Cleaner

Upload a messy CSV or Excel file and get back a clean, review-ready Excel report.
Nothing is deleted silently: fixable problems are corrected, doubtful rows are
flagged with a reason, and duplicates are kept on their own sheet.

## What it does
- Standardises names (spacing and capitalisation) and emails (trimmed, lowercase)
- Cleans phone numbers and flags invalid ones (length rules and country code are adjustable)
- Converts mixed date formats to real dates (day-first or month-first)
- Converts numbers such as "€1,200" and flags text where a number is expected
- Finds duplicates using the column you choose
- Exports an Excel file with four sheets: Summary, Clean data, Needs review, Duplicates removed

## Demo
[Watch the 45-second demo](link) · screenshot here

## How to run
pip install -r requirements.txt
streamlit run app.py

## Known limits
- Capitalisation turns "McDonald" into "Mcdonald"
- Phone checks look at length only, not whether the number exists
- Dates like 04/05/2021 are ambiguous, so you choose the order in the sidebar

## Sample data
All files in `sample-data/` are synthetic. No real personal data is included.
