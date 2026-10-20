# Day 3 — Excel, Word and Python File Processing

## Goal
Read a sample registration workbook, clean and organize its contents with Python, and create a Word report programmatically.

## Files
- `process_files.py` — complete Python automation script.
- `club_registration_data.xlsx` — fictional sample registration data containing extra spaces, mixed email capitalization, one duplicate email, and one missing email.
- `organized_club_registrations.xlsx` — cleaned Excel output.
- `club_registration_cleaning_report.docx` — generated Word report.
- `output_screenshot.png` — visual record of the successful script output.
- `requirements.txt` — Python dependencies.

## Run locally
1. Install Python.
2. Open a terminal in this folder.
3. Install libraries:
   ```bash
   pip install -r requirements.txt
   ```
4. Run:
   ```bash
   python process_files.py
   ```

## What the program does
1. Reads the Excel workbook with pandas.
2. Trims unwanted spaces and standardizes column names.
3. Converts email addresses to lowercase and flags missing email addresses.
4. Standardizes department names to CSE, IT, and DS.
5. Removes duplicate registrations based on email.
6. Sorts registrations by department and student name.
7. Saves the organized Excel workbook with headers, filters, and frozen top row.
8. Creates a Word report with counts, cleaning steps, and a preview table.

All registration records are fictional and are used for practice only.
