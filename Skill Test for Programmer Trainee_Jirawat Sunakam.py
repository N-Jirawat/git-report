import pandas as pd
import os
import re

# -------------------------------
# Specify the folder where the Excel file is stored
# -------------------------------

# # Specify the folder address of the .xlsx file.
# กำหนดที่อยู่โฟลเดอร์ของไฟล์ .xlsx

FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))

# -------------------------------
# Read the New Employee file (New Employee)
# -------------------------------
def load_new_employee(folder):
    for filename in os.listdir(folder):
        if filename.startswith("New Employee_") and filename.endswith(".xlsx"):
            file_path = os.path.join(folder, filename)
            df = pd.read_excel(file_path)
            return df
    return pd.DataFrame() 

# -------------------------------
# อ่านไฟล์ Daily Report หลายไฟล์
# -------------------------------
def load_daily_reports(folder):
    all_reports = []
    
    for filename in os.listdir(folder):
        if filename.startswith("Daily report_") and (filename.endswith(".xls") or filename.endswith(".xlsx")):
            match = re.match(r"Daily report_\d+_(.+)_(.+)\.xlsx", filename)
            if match:
                team_member = f"{match.group(1)} {match.group(2)}"
                file_path = os.path.join(folder, filename)
                df = pd.read_excel(file_path)
                
                df.columns = df.columns.str.strip()
                
                # Add a Team Member column.
                # เพิ่มคอลัมน์ Team Member
                df["Team Member"] = team_member
                all_reports.append(df)
    
    if all_reports:
        return pd.concat(all_reports, ignore_index=True)
    else:
        return pd.DataFrame()

# -------------------------------
# Match new employee information with the team
# -------------------------------
def match_employees(new_employees, daily_reports):

    # Filter only people who passed the interview
    # กรองเฉพาะคนที่ผ่านการสัมภาษณ์
    passed = daily_reports[
        (daily_reports["Interview"].str.lower() == "yes") &
        (daily_reports["Status"].str.lower() == "pass")
    ]

    #print("Passed interview records:")
    #print(passed[["Candidate Name", "Role", "Team Member"]])

    # Rename the columns to match.
    # เปลี่ยนชื่อคอลัมน์ให้ตรงกัน
    passed = passed.rename(columns={"Candidate Name": "Employee Name"})

    # Combine matching information from both sides
    # รวมข้อมูลที่ตรงกันจากทั้งสองฝั่ง
    merged = pd.merge(
        new_employees,
        passed[["Employee Name", "Role", "Team Member"]],
        on=["Employee Name", "Role"],
        how="inner"
    )

    # Show only the columns you want to show.
    # แสดงแค่คอลัมน์ที่ต้องการโชว์
    result = merged[["Employee Name", "Join Date", "Role", "Team Member"]]
    return result

# -------------------------------
# Start
# -------------------------------
def main():
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    FOLDER_PATH = SCRIPT_DIR 


    new_employees = load_new_employee(FOLDER_PATH)
    daily_reports = load_daily_reports(FOLDER_PATH)

    if new_employees.empty or daily_reports.empty:
        print("Data not found in specified Excel file.")
        return

    result = match_employees(new_employees, daily_reports)

    if result.empty:
        print("No matching employee records found.")
        return

    print("\nList of new employees according to the interview team:")
    print(result.to_string(index=False))

    os.makedirs(FOLDER_PATH, exist_ok=True)

    output_path = os.path.join(FOLDER_PATH, "Matched_Employees_Report.xlsx")
    try:
        result.to_excel(output_path, index=False)
        relative_path = os.path.relpath(output_path, SCRIPT_DIR)
        print(f"\n Result saved to: {relative_path}")
    except Exception as e:
        print(f" Error saving Excel file: {e}")

# -------------------------------
# Run
# -------------------------------
if __name__ == "__main__":
    main()
