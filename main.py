import os
import gdown
import pandas as pd
from flask import Flask

app = Flask(__name__)

DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

FILE_IDS = {
    "new_employee": "1Wb5dBl1DL-NsTVzbxc1Pv2E-EZo2dzqh",
    "daily_raewwadee": "1v4YxphVcP8j24cDcxC0OkwzgXmaoy3ek",
    "daily_pattama": "1TxIOKhR2ZDMPAZAERpNlDYI8my63HW2S"
}

def download_google_sheet_as_excel(file_id, filename):
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    output_path = os.path.join(DOWNLOAD_DIR, filename)
    if not os.path.exists(output_path):
        print(f"Downloading {filename} from Google Sheets...")
        gdown.download(url, output_path, quiet=False)
    else:
        print(f"{filename} already downloaded.")
    return output_path

def load_new_employee():
    path = download_google_sheet_as_excel(FILE_IDS["new_employee"], "New_Employee.xlsx")
    return pd.read_excel(path)

def load_daily_reports():
    dfs = []
    for key in ["daily_raewwadee", "daily_pattama"]:
        filename = key + ".xlsx"
        path = download_google_sheet_as_excel(FILE_IDS[key], filename)
        df = pd.read_excel(path)
        df.columns = df.columns.str.strip()
        team_member = "Raewwadee Jaidee" if key == "daily_raewwadee" else "Pattama Sooksan"
        df["Team Member"] = team_member
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def match_employees(new_employees, daily_reports):
    passed = daily_reports[
        (daily_reports["Interview"].str.lower() == "yes") &
        (daily_reports["Status"].str.lower() == "pass")
    ]
    passed = passed.rename(columns={"Candidate Name": "Employee Name"})
    merged = pd.merge(
        new_employees,
        passed[["Employee Name", "Role", "Team Member"]],
        on=["Employee Name", "Role"],
        how="inner"
    )
    return merged[["Employee Name", "Join Date", "Role", "Team Member"]]

def main():
    new_employees = load_new_employee()
    daily_reports = load_daily_reports()

    if new_employees.empty or daily_reports.empty:
        print("Data not found or failed to download.")
        return

    result = match_employees(new_employees, daily_reports)

    if result.empty:
        print("No matching employee records found.")
        return

    print("\nList of new employees according to the interview team:")
    print(result.to_string(index=False))

    output_path = os.path.join(DOWNLOAD_DIR, "Matched_Employees_Report.xlsx")
    try:
        result.to_excel(output_path, index=False)
        print(f"\nResult saved to: {output_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

@app.route('/')
def index():
    return "Skill Test for Programmer Trainee_Jirawat Sunakam"

if __name__ == "__main__":
    app.run(debug=True)
