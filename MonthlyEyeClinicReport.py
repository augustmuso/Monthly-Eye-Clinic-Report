import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime


# Set up the credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("eye-clinic-report.json", scopes=scope)
client = gspread.authorize(creds)


# Open the Google Sheet
spreadsheet = client.open("Patient Tracker")
worksheet = spreadsheet.sheet1  # or use .get_worksheet(index) if multiple sheets

# Fetch all data from the sheet
data = worksheet.get_all_records()

# Convert data to a DataFrame for easier manipulation
df = pd.DataFrame(data)

# Get the current year
current_year = datetime.now().year

# Append the current year to the 'Date' strings
df['Date'] = df['Date'] + f" {current_year}"

# Convert the 'Date' column to datetime
df['Date'] = pd.to_datetime(df['Date'], format='%d %b %Y')

# Set 'Date' as the index
df.set_index('Date', inplace=True)

# Filter data for the current month
current_month = datetime.now().month
df = df[df.index.month == current_month]

# Ensure numeric conversion for New Px and Returning Px columns
df['New Px'] = pd.to_numeric(df['New Px'], errors='coerce')
df['Returning Px'] = pd.to_numeric(df['Returning Px'], errors='coerce')

# Calculate the total patients per day
df['Total'] = df['New Px'] + df['Returning Px']

# Select only numeric columns for resampling
numeric_columns = ['New Px', 'Returning Px', 'Total']

# Aggregate data by week, summing the number of tests, using the first day of the week as the index
weekly_summary = df[numeric_columns].resample('W').sum()

# Compute Monthly Summary
total_patients = df['Total'].sum()
total_new_patients = df['New Px'].sum()
total_returning_patients = df['Returning Px'].sum()

# Collect all comments for the current month into a single string
comments = df['Comments'].dropna().tolist()
all_comments = "\n".join(comments)

# Prepare LaTeX content
latex_content = f"""
\\documentclass[a4paper, 12pt]{{article}}
\\usepackage{{geometry}}
\\usepackage{{array}}
\\usepackage{{longtable}}
\\usepackage{{booktabs}} % Improved table aesthetics
\\geometry{{margin=1in}}
\\title{{Monthly Eye Clinic Report}}
\\author{{Musoke Ntaate Augustine}}
\\date{{\\today}}

\\begin{{document}}

\\maketitle

\\section*{{Monthly Summary}}
\\begin{{tabbing}}
\\hspace*{{5.5cm}}\\= \\kill
\\textbf{{Month:}} \\> {datetime.now().strftime('%B, %Y')} \\\\
\\textbf{{Total Patients Tested:}} \\> {total_patients} \\\\
\\textbf{{Total New Patients:}} \\> {total_new_patients} \\\\
\\textbf{{Total Returning Patients:}} \\> {total_returning_patients} \\\\
\\end{{tabbing}}

\\section*{{Weekly Breakdown}}
\\begin{{longtable}}{{| p{{4cm}} | p{{3cm}} | p{{3cm}} | p{{3cm}} |}}
\\toprule
\\textbf{{Week End}} & \\textbf{{Total Px}} & \\textbf{{New Px}} & \\textbf{{Returning Px}} \\\\ 
\\midrule
\\endfirsthead
\\toprule
\\textbf{{Week End}} & \\textbf{{Total Px}} & \\textbf{{New Px}} & \\textbf{{Returning Px}} \\\\ 
\\midrule
\\endhead
"""

# Add weekly summary to LaTeX table
for date, row in weekly_summary.iterrows():
    week_end = date.strftime('%Y-%m-%d')
    latex_content += f"{week_end} & {row['Total']} & {row['New Px']} & {row['Returning Px']} \\\\ \\midrule\n"

# Add final row for total monthly summary
latex_content += f"\\textbf{{Monthly Totals}} & \\textbf{{{total_patients}}} & \\textbf{{{total_new_patients}}} & \\textbf{{{total_returning_patients}}} \\\\ \\bottomrule\n"

# Close the table and add the comments section
latex_content += """
\\end{longtable}

\\section*{Comments/Notes}
\\begin{itemize}
"""

# Add all comments to the LaTeX document for manual summarization
latex_content += f"""
    \\item All Comments:
    \\begin{{verbatim}}
    {all_comments}
    \\end{{verbatim}}
"""

latex_content += """
\\end{itemize}

\\end{document}
"""

# Save the LaTeX content to a file
with open('monthly_eye_clinic_report.tex', 'w') as f:
    f.write(latex_content)

print("LaTeX report with current month's comments has been saved to 'monthly_eye_clinic_report.tex'")
