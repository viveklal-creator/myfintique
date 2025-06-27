from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import os, json
import pandas as pd
from io import BytesIO
from flask import send_file
import xlsxwriter

app = Flask(__name__)
app.secret_key = "your_secret_key"

DATA_DIR = "data"
TB_UPLOAD_FOLDER = os.path.join(DATA_DIR, "tb_files")
REMARKS_FILE = os.path.join(DATA_DIR, "tb_remarks.json")

os.makedirs(TB_UPLOAD_FOLDER, exist_ok=True)

# Dummy users
USERS = {
    "vivek.lal@batterysmart.in": {
        "password": "Vivek@1234",
        "role": "admin"
    },
    "anubhav.jain@batterysmart.in": {
        "password": "Anubhav@1234",
        "role": "manager"
    }
}

ALLOWED_EXTENSIONS = {'csv', 'xlsx'}
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]
        user = USERS.get(email)
        if user and user["password"] == password:
            session.permanent = True  # Auto-logout after inactivity period
            session["user"] = email
            session["role"] = user["role"]
            return redirect(url_for("dashboard"))
        else:
            return render_template("login.html", error="Invalid credentials")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("login"))

@app.route("/upload_tb", methods=["GET", "POST"])
def upload_tb():
    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        tb_month = request.form.get("tb_month")
        file = request.files.get("file")

        if file and allowed_file(file.filename) and tb_month:
            ext = file.filename.rsplit('.', 1)[1].lower()
            filename = f"{tb_month}.{ext}"
            filepath = os.path.join(TB_UPLOAD_FOLDER, filename)
    
            if ext == 'csv':
                df = pd.read_csv(file)
                df.to_csv(filepath, index=False)
            else:
                df = pd.read_excel(file)
                df.to_excel(filepath, index=False)  # ✅ Save as real Excel
    
            return redirect(url_for("dashboard"))


        return render_template("upload_tb.html", error="Invalid file or TB month selection")

    return render_template("upload_tb.html")

@app.route("/mom_tb_comparison")
def mom_tb_comparison():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("mom_tb_comparison.html")

@app.route("/api/mom_tb_comparison")
def api_mom_tb_comparison():
    if "user" not in session:
        return jsonify([])

    all_data = {}
    months = []

    tb_files = sorted([
        f for f in os.listdir(TB_UPLOAD_FOLDER)
        if f.endswith('.csv')
    ])

    # Fix: Recreate dropdown-style month label e.g., April'25 from filename
    def file_to_month(filename):
        name = filename.replace(".csv", "").replace(".xlsx", "").replace("_", "")
        month_map = {
            "April25": "April'25", "May25": "May'25", "June25": "June'25", "July25": "July'25",
            "August25": "August'25", "September25": "September'25", "October25": "October'25",
            "November25": "November'25", "December25": "December'25", "January26": "January'26",
            "February26": "February'26", "March26": "March'26"
        }
        return month_map.get(name, name)


    for fname in tb_files:
        month_label = file_to_month(fname)
        months.append(month_label)

        df = pd.read_csv(os.path.join(TB_UPLOAD_FOLDER, fname))

        required_cols = {"Expense Type", "GL Name", "Amount"}
        if not required_cols.issubset(df.columns):
            continue

        for _, row in df.iterrows():
            etype = row["Expense Type"].strip()
            gl = row["GL Name"].strip()
            amount_str = str(row["Amount"]).replace(",", "").strip()
            if amount_str.startswith("(") and amount_str.endswith(")"):
                amount_str = "-" + amount_str[1:-1]
            amt = float(amount_str)
            key = (etype, gl)

            if key not in all_data:
                all_data[key] = {}
            all_data[key][month_label] = all_data[key].get(month_label, 0) + amt

    # Sort months based on dropdown order
    month_order = {
        "April'25": 1, "May'25": 2, "June'25": 3, "July'25": 4, "August'25": 5,
        "September'25": 6, "October'25": 7, "November'25": 8, "December'25": 9,
        "January'26": 10, "February'26": 11, "March'26": 12
    }
    months = sorted(set(months), key=lambda m: month_order.get(m, 99))

    response = {
        "months": months,
        "rows": []
    }

    for (etype, gl), month_vals in sorted(all_data.items()):
        row = {
            "type": etype,
            "gl": gl,
            "months": {},
            "variance": "",
            "variance_pct": ""
        }

        for m in months:
            row["months"][m] = f"{month_vals.get(m, 0):,.0f}"

        # Last 2-month variance logic
        if len(months) >= 2:
            m1, m2 = months[-2], months[-1]
            v1 = month_vals.get(m1, 0)
            v2 = month_vals.get(m2, 0)
            variance = v2 - v1
            pct = (variance / v1 * 100) if v1 != 0 else 0
            row["variance"] = f"{variance:,.0f}"
            row["variance_pct"] = f"{pct:.1f}"

        response["rows"].append(row)

    return jsonify(response)

REMARKS_FILE = os.path.join(DATA_DIR, "tb_remarks.json")

@app.route('/api/save_remarks', methods=['POST'])
def save_remarks():
    if "user" not in session:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401

    data = request.get_json()
    month = data.get("month")
    remarks = data.get("remarks")
    user_role = session.get("role")

    if not month or not remarks:
        return jsonify({"status": "error", "message": "Invalid data"}), 400

    # Load existing file
    if os.path.exists(REMARKS_FILE):
        with open(REMARKS_FILE, "r") as f:
            existing = json.load(f)
    else:
        existing = {}

    if month not in existing:
        existing[month] = {}

    for gl, content in remarks.items():
        if gl not in existing[month]:
            existing[month][gl] = {}

        # Admin can only write 'remark'
        if user_role == "admin":
            if "manager_remark" in content:
                return jsonify({"status": "error", "message": "You don't have permission to update manager remarks"}), 403
            if "remark" in content:
                existing[month][gl]["remark"] = content["remark"]

        # Manager can only write 'manager_remark'
        elif user_role == "manager":
            if "remark" in content:
                return jsonify({"status": "error", "message": "You don't have permission to update remarks"}), 403
            if "manager_remark" in content:
                existing[month][gl]["manager_remark"] = content["manager_remark"]

    # Save file
    with open(REMARKS_FILE, "w") as f:
        json.dump(existing, f, indent=2)

    return jsonify({"status": "success"})

@app.route("/api/get_remarks")
def get_remarks():
    if "user" not in session:
        return jsonify({"error": "Unauthorized"}), 401

    if os.path.exists(REMARKS_FILE):
        with open(REMARKS_FILE, "r") as f:
            json_data = json.load(f)
        return jsonify(json_data)  # return full month-based remarks
    else:
        return jsonify({})

@app.route("/gl_details")
def gl_details():
    if "user" not in session:
        return redirect(url_for("login"))

    gl_name = request.args.get("gl")
    if not gl_name:
        return "GL not specified", 400

    # Load all uploaded TB files
    tb_files = sorted([
        f for f in os.listdir(TB_UPLOAD_FOLDER)
        if f.endswith(".csv")
    ])

    def file_to_month(filename):
        name = filename.replace(".csv", "").replace(".xlsx", "").replace("_", "")
        month_map = {
            "April25": "April'25", "May25": "May'25", "June25": "June'25", "July25": "July'25",
            "August25": "August'25", "September25": "September'25", "October25": "October'25",
            "November25": "November'25", "December25": "December'25", "January26": "January'26",
            "February26": "February'26", "March26": "March'26"
        }
        return month_map.get(name, name)

    result = {}
    totals = {}
    months = []

    for file in tb_files:
        month_label = file_to_month(file)
        months.append(month_label)
        df = pd.read_csv(os.path.join(TB_UPLOAD_FOLDER, file))

        if {"GL Name", "Vendor Name", "Amount"}.issubset(df.columns):
            filtered = df[df["GL Name"].astype(str).str.strip() == gl_name.strip()]
            for _, row in filtered.iterrows():
                vendor = str(row["Vendor Name"]).strip()
                amount_raw = row["Amount"]
                try:
                    amount = float(str(amount_raw).replace(",", "").replace("(", "-").replace(")", ""))
                    amount = round(amount)
                except (ValueError, TypeError):
                    amount = 0

                if vendor not in result:
                    result[vendor] = {}
                result[vendor][month_label] = result[vendor].get(month_label, 0) + amount

                totals[month_label] = totals.get(month_label, 0) + amount

    # Format numbers as comma-style
    for vendor in result:
        for m in result[vendor]:
            try:
                result[vendor][m] = f"{int(float(result[vendor][m])):,}"
            except (ValueError, TypeError):
                result[vendor][m] = "0"

    for m in totals:
        try:
            totals[m] = f"{int(float(totals[m])):,}"
        except (ValueError, TypeError):
            totals[m] = "0"

    return render_template("gl_details.html", gl_name=gl_name, rows=result, months=sorted(set(months)), totals=totals)

@app.route("/export_mom_tb_excel")
def export_mom_tb_excel():
    if "user" not in session:
        return redirect(url_for("login"))

    # Reuse the same data logic as the API
    response = api_mom_tb_comparison().get_json()
    months = response["months"]
    rows = response["rows"]

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("MoM TB")

    header = ["Type", "GL"] + months + ["Variance", "Variance %"]
    for col, title in enumerate(header):
        worksheet.write(0, col, title)

    for row_idx, row in enumerate(rows, start=1):
        worksheet.write(row_idx, 0, row["type"])
        worksheet.write(row_idx, 1, row["gl"])
        for col_idx, month in enumerate(months, start=2):
            worksheet.write(row_idx, col_idx, row["months"].get(month, ""))
        worksheet.write(row_idx, 2 + len(months), row["variance"])
        worksheet.write(row_idx, 3 + len(months), row["variance_pct"])

    workbook.close()
    output.seek(0)

    return send_file(
        output,
        download_name="MoM_TB_Comparison.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

print("==> Entered /api/mom_tb_comparison")
print("Files in TB_UPLOAD_FOLDER:", os.listdir(TB_UPLOAD_FOLDER))


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)