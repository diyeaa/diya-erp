from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from datetime import datetime
from flask import flash, get_flashed_messages
import os


app = Flask(__name__)

app.secret_key = os.urandom(24)

FILE = "inventory.xlsx"

# -----------------------------
# Load and save inventory
# -----------------------------
def load_inventory():
    return pd.read_excel(FILE)

def save_inventory(df):
    df.to_excel(FILE, index=False)


def record_sale(product_id, quantity):
    sales_df = pd.read_excel("sales.xlsx")

    new_sale = {
        "Product_ID": product_id,
        "Quantity_Sold": quantity,
        "Date": datetime.now().strftime("%Y-%m-%d")
    }

    sales_df = pd.concat([sales_df, pd.DataFrame([new_sale])], ignore_index=True)
    sales_df.to_excel("sales.xlsx", index=False)

    inv_df = pd.read_excel("inventory.xlsx")
    inv_df.loc[inv_df["Product_ID"] == product_id, "Current_Stock"] -= quantity
    inv_df["Current_Stock"] = inv_df["Current_Stock"].clip(lower=0)
    inv_df.to_excel("inventory.xlsx", index=False)
    

def daily_sales_report(date):
    sales_df = pd.read_excel("sales.xlsx")

    # Filter by date
    daily_data = sales_df[sales_df["Date"] == date]

    # Group by Product_ID and sum quantity
    report = daily_data.groupby("Product_ID")["Quantity_Sold"].sum().reset_index()

    return report


def most_demanded_product():
    sales_df = pd.read_excel("sales.xlsx")
    total_sales = sales_df.groupby("Product_ID")["Quantity_Sold"].sum().reset_index()
    if not total_sales.empty:
        top_product_id = total_sales.loc[total_sales["Quantity_Sold"].idxmax(), "Product_ID"]
        product_name = pd.read_excel(FILE).loc[pd.read_excel(FILE)["Product_ID"]==top_product_id, "Product_Name"].values[0]
        total_qty = total_sales["Quantity_Sold"].max()
        return product_name, total_qty
    return None, 0
    
    
def product_classification(days=30):
    sales_df = pd.read_excel("sales.xlsx")
    total_sales = sales_df.groupby("Product_ID")["Quantity_Sold"].sum().reset_index()
    inventory_df = pd.read_excel(FILE)
    classification = {}
    for _, row in total_sales.iterrows():
        pid = row["Product_ID"]
        avg_daily = row["Quantity_Sold"] / days
        name = inventory_df.loc[inventory_df["Product_ID"]==pid, "Product_Name"].values[0]
        if avg_daily >= 2:  # threshold example
            classification[name] = "Fast-Moving"
        else:
            classification[name] = "Slow-Moving"
    return classification




EMP_FILE = "hrm_data/employees.xlsx"
ATT_FILE = "hrm_data/attendance.xlsx"

def load_employees():
    return pd.read_excel(EMP_FILE)

def load_attendance():
    return pd.read_excel(ATT_FILE)

def save_attendance(df):
    df.to_excel(ATT_FILE, index=False)

@app.route("/hr_management", methods=["GET", "POST"])
def hr_management():
    attendance_df = load_attendance()
    employees_df = load_employees()

    # ---- Handle attendance submission ----
    if request.method == "POST":
        emp_id = int(request.form["emp_id"])
        today = datetime.now().strftime("%Y-%m-%d")

        # Check if attendance already exists for this employee today
        if not ((attendance_df["Emp_ID"] == emp_id) & (attendance_df["Date"] == today)).any():
            new_row = {
                "Emp_ID": emp_id,
                "Date": today,
                "Status": request.form["status"]
            }
            attendance_df = pd.concat(
                [attendance_df, pd.DataFrame([new_row])],
                ignore_index=True
            )
            save_attendance(attendance_df)
        else:
            flash("Attendance for this employee has already been recorded today.", "warning")

        # Redirect after POST
        return redirect(url_for("hr_management"))

    # ---- Payroll calculation ----
    payroll_data = []
    for _, emp in employees_df.iterrows():
        days_present = attendance_df[
            (attendance_df["Emp_ID"] == emp["Emp_ID"]) &
            (attendance_df["Status"] == "Present")
        ].shape[0]

        daily_salary = emp["Basic_Salary"] / 30
        net_salary = daily_salary * days_present

        payroll_data.append({
            "Emp_ID": emp["Emp_ID"],
            "Name": emp["Name"],
            "Department": emp["Department"],
            "Net_Salary": round(net_salary, 2)
        })

    return render_template(
        "hr_management.html",
        attendance=attendance_df.to_dict(orient="records"),
        payroll=payroll_data
    )

CRM_FILE = "hrm_data/crm.xlsx"
# -----------------------------
# CRM Module
# -----------------------------
def load_crm():
    try:
        return pd.read_excel(CRM_FILE)
    except FileNotFoundError:
        return pd.DataFrame(columns=["Customer_Name", "Feedback", "Rating", "Date"])

def save_crm(df):
    df.to_excel(CRM_FILE, index=False)


@app.route("/crm", methods=["GET", "POST"])
def crm():
    crm_df = load_crm()

    if request.method == "POST":
        customer_name = request.form["customer_name"]
        feedback = request.form["feedback"]
        rating = float(request.form["rating"])
        date = datetime.now().strftime("%Y-%m-%d")

        new_row = {
            "Customer_Name": customer_name,
            "Feedback": feedback,
            "Rating": rating,
            "Date": date
        }

        crm_df = pd.concat([crm_df, pd.DataFrame([new_row])], ignore_index=True)
        save_crm(crm_df)
        flash("Feedback submitted successfully!", "success")
        return redirect(url_for("crm"))

    avg_rating = round(crm_df["Rating"].mean(), 2) if not crm_df.empty else 0

    return render_template(
        "crm.html",
        reviews=crm_df.to_dict(orient="records"),
        average_rating=avg_rating
    )



# -----------------------------
# Home page: Inventory & Alerts
# -----------------------------
@app.route("/")
def index():
    df = load_inventory()
    alerts = df[df["Current_Stock"] < df["Threshold"]]
    classification = product_classification()  # calculate classification

    most_demanded, qty = most_demanded_product()  # optional, if you use it

    return render_template(
        "index.html",
        inventory=df.to_dict(orient="records"),
        alerts=alerts.to_dict(orient="records"),
        classification=classification,
        most_demanded=most_demanded,
        max_sales=qty
    )

# -----------------------------
# Sell Product
# -----------------------------
@app.route("/sell", methods=["POST"])
def sell():
    product_id = int(request.form["product_id"])
    quantity = int(request.form["quantity"])

    record_sale(product_id, quantity)   # CALL THE FUNCTION

    return redirect(url_for("index"))

# -----------------------------
# Restock Product
# -----------------------------
@app.route("/restock", methods=["POST"])
def restock():
    product_id = int(request.form["product_id"])
    quantity = int(request.form["quantity"])
    df = load_inventory()
    df.loc[df["Product_ID"] == product_id, "Current_Stock"] += quantity
    save_inventory(df)
    return redirect(url_for("index"))


@app.route("/dashboard")
def dashboard():
    inventory_df = load_inventory()
    alerts = inventory_df[inventory_df["Current_Stock"] < inventory_df["Threshold"]]
    most_demanded, qty = most_demanded_product()
    classification = product_classification()
    
    return render_template(
        "dashboard.html",
        inventory=inventory_df.to_dict(orient="records"),
        alerts=alerts.to_dict(orient="records"),
        most_demanded=most_demanded,
        max_sales=qty,
        classification=classification
    )



@app.route("/export_dss")
def export_dss():
    inventory_df = load_inventory()
    classification = product_classification()
    report = []
    for _, item in inventory_df.iterrows():
        restock = "Yes" if item["Current_Stock"] < item["Threshold"] else "No"
        name = item["Product_Name"]
        report.append({
            "Product_ID": item["Product_ID"],
            "Product_Name": name,
            "Current_Stock": item["Current_Stock"],
            "Threshold": item["Threshold"],
            "Restock": restock,
            "Category": classification.get(name, "N/A")
        })
    df = pd.DataFrame(report)
    df.to_excel("DSS_Report.xlsx", index=False)
    return "DSS report exported successfully!"



@app.route("/add_product", methods=["POST"])
def add_product():
    product_name = request.form["product_name"]
    category = request.form["category"]
    stock = int(request.form["stock"])
    threshold = int(request.form["threshold"])

    df = load_inventory()

    # Generate new Product_ID (e.g., max ID + 1)
    if not df.empty:
        new_id = df["Product_ID"].max() + 1
    else:
        new_id = 1

    new_product = {
        "Product_ID": new_id,
        "Product_Name": product_name,
        "Category": category,
        "Current_Stock": stock,
        "Threshold": threshold
    }

    df = pd.concat([df, pd.DataFrame([new_product])], ignore_index=True)
    save_inventory(df)

    return redirect(url_for("index"))






# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)


