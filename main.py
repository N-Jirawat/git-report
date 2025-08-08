from flask import Flask, jsonify, render_template_string
import pandas as pd
import os
import gdown

app = Flask(__name__)

# ------------------------------
# DOWNLOAD & PROCESS
# ------------------------------

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
        gdown.download(url, output_path, quiet=False)
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

# ------------------------------
# ROUTES
# ------------------------------

@app.route("/")
def index():
    try:
        new_employees = load_new_employee()
        daily_reports = load_daily_reports()
        result = match_employees(new_employees, daily_reports)

        if result.empty:
            return "No matching employee records found."

        # Convert DataFrame to HTML table
        html_table = result.to_html(index=False)

        # Render as simple web page
        html_template = f"""
        <html>
            <head><title>Matched Employees</title></head>
            <body>
                <h1>Matched Employees</h1>
                {html_table}
            </body>
        </html>
        """
        return render_template_string(html_template)

    except Exception as e:
        return f"Error occurred: {str(e)}"

@app.route("/match")
def match_route():
    new_employees = load_new_employee()
    daily_reports = load_daily_reports()
    result = match_employees(new_employees, daily_reports)

    if result.empty:
        return jsonify({"message": "No matches found."}), 200

    return jsonify(result.to_dict(orient="records")), 200

# ------------------------------
# MAIN
# ------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
