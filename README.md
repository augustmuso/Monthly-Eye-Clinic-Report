# README for Monthly Eye Clinic Report Script

## Overview
This Python script generates a **Monthly Eye Clinic Report** using data from a Google Sheet. The report includes a summary of patient data, weekly breakdowns, and comments. The output is formatted in LaTeX, generating a report that can be compiled into a PDF.

## Prerequisites
Before running the script, ensure you have the following:
1. **Google API Credentials**: A JSON key file for accessing Google Sheets, named `eye-clinic-report.json`.
2. **Google Sheet**: A Google Sheet titled `Patient Tracker` with the following columns:
   - **Date**: Date in `DD MMM` format (e.g., `01 Sep`).
   - **New Px**: Number of new patients.
   - **Returning Px**: Number of returning patients.
   - **Comments**: Notes or feedback from patients.

## Libraries Required
The script uses the following Python libraries:
1. `gspread`: For interacting with Google Sheets.
2. `google.oauth2.service_account`: For authenticating with Google services.
3. `pandas`: For data manipulation.
4. `datetime`: To handle date and time functions.

### Install Dependencies
To install the required libraries, run the following:
```bash
pip install gspread pandas google-auth
```

## How to Run the Script
1. **Google API Setup**:
   - Obtain a Google API JSON key file (`eye-clinic-report.json`).
   - Share the Google Sheet (`Patient Tracker`) with the service account email from the credentials.
   
2. **Modify the Spreadsheet and Worksheet**:
   - Ensure your Google Sheet matches the expected structure.
   - Modify the `spreadsheet = client.open("Patient Tracker")` line to refer to the correct sheet name, if different.

3. **Run the Script**:
   - Run the script in a Python environment:
   ```bash
   python eye_clinic_report.py
   ```

4. **Output**:
   - The script generates a LaTeX file named `monthly_eye_clinic_report.tex`.
   - This file includes a summary of patient data for the current month, a weekly breakdown, and all comments for the month.

## Key Features
- **Google Sheets Integration**: Fetches patient data directly from Google Sheets.
- **Data Aggregation**: Calculates totals for new, returning, and overall patients on a weekly and monthly basis.
- **LaTeX Report Generation**: Outputs a detailed LaTeX report that can be compiled into PDF.

## File Structure
- `eye-clinic-report.json`: Google API credentials.
- `monthly_eye_clinic_report.tex`: The LaTeX file with the report content.

## Troubleshooting
- Ensure the Google Sheets API is enabled and the service account has access to the sheet.
- Check that the date format in the Google Sheet matches `DD MMM`.

## License
This script is licensed under the MIT License.

## Author
Musoke Ntaate Augustine

