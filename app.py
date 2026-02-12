from flask import Flask, render_template, request, redirect, session
import pandas as pd
from datetime import date
import os

app = Flask(__name__)
app.secret_key = "inventory_secret_key"

INVENTORY_FILE = "inventory.xlsx"
SALES_FILE = "sales.xlsx"
USERS_FILE = "users.xlsx"
EMP_FILE = "employees.xlsx"
ATT_FILE = "attendance.xlsx"
PAY_FILE = "payroll.xlsx"
PERF_FILE = "performance.xlsx"


# ---------------- DSS LOGIC ----------------

def get_inventory_status():
    inventory = pd.read_excel(INVENTORY_FILE)
    alerts = inventory[inventory["Current_Stock"] < inventory["Threshold"]]
    return inventory, alerts


def sales_summary_by_date(selected_date):
    sales = pd.read_excel(SALES_FILE)
    sales["Date"] = pd.to_datetime(sales["Date"]).dt.date

    return (
        sales[sales["Date"] == selected_date]
        .groupby("Product_ID")["Quantity Sold"]
        .sum()
        .reset_index()
    )


def inventory_sales_dss():
    inventory = pd.read_excel(INVENTORY_FILE)
    sales = pd.read_excel(SALES_FILE)

    sales_summary = (
        sales.groupby("Product_ID")["Quantity Sold"]
        .sum()
        .reset_index()
    )

    dss = pd.merge(inventory, sales_summary, on="Product_ID", how="left").fillna(0)

    dss["Decision"] = dss.apply(
        lambda row: "Restock"
        if row["Current_Stock"] < row["Threshold"]
        else "No Action",
        axis=1
    )

    return dss


def classify_products():
    inventory = pd.read_excel(INVENTORY_FILE)
    sales = pd.read_excel(SALES_FILE)

    sales["Date"] = pd.to_datetime(sales["Date"]).dt.date

    summary = (
        sales.groupby("Product_ID")
        .agg(
            total_sold=("Quantity Sold", "sum"),
            days_sold=("Date", "nunique")
        )
        .reset_index()
    )

    summary["Avg Daily Sales"] = summary["total_sold"] / summary["days_sold"]

    summary["Category"] = summary["Avg Daily Sales"].apply(
        lambda avg: "Fast Moving" if avg >= 10 else
                    "Medium Moving" if avg >= 5 else
                    "Slow Moving"
    )

    return pd.merge(
        inventory[["Product_ID", "Product_Name"]],
        summary,
        on="Product_ID",
        how="left"
    ).fillna(0).to_dict(orient="records")


# ---------------- AUTH ROUTES ----------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        users = pd.read_excel(USERS_FILE)
        user = users[
            (users["Username"] == username) &
            (users["Password"] == password)
        ]

        if not user.empty:
            session["user"] = username
            session["role"] = user.iloc[0]["Role"]
            return redirect("/")
        else:
            return "❌ Invalid credentials"

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ---------------- MAIN ROUTES ----------------

@app.route("/hrm")
def hrm():
    if "user" not in session:
        return redirect("/login")

    employees = pd.read_excel(EMP_FILE).to_dict(orient="records")
    attendance = pd.read_excel(ATT_FILE).to_dict(orient="records")
    payroll = pd.read_excel(PAY_FILE).to_dict(orient="records")
    performance = pd.read_excel(PERF_FILE).to_dict(orient="records")

    return render_template(
        "hrm.html",
        employees=employees,
        attendance=attendance,
        payroll=payroll,
        performance=performance
    )


@app.route("/add_employee", methods=["POST"])
def add_employee():
    emp_id = int(request.form["emp_id"])
    name = request.form["name"]
    role = request.form["role"]
    salary = int(request.form["salary"])

    df = pd.read_excel(EMP_FILE)
    df = pd.concat([df, pd.DataFrame([{
        "Employee_ID": emp_id,
        "Name": name,
        "Role": role,
        "Salary": salary
    }])], ignore_index=True)

    df.to_excel(EMP_FILE, index=False)
    return redirect("/hrm")


@app.route("/mark_attendance", methods=["POST"])
def mark_attendance():
    df = pd.read_excel(ATT_FILE)

    new = pd.DataFrame([{
        "Employee_ID": request.form["emp_id"],
        "Name": request.form["name"],
        "Date": date.today(),
        "Status": request.form["status"]
    }])

    df = pd.concat([df, new], ignore_index=True)
    df.to_excel(ATT_FILE, index=False)
    return redirect("/hrm")


@app.route("/process_payroll", methods=["POST"])
def process_payroll():
    df = pd.read_excel(PAY_FILE)

    salary = int(request.form["salary"])
    bonus = int(request.form["bonus"])

    new = pd.DataFrame([{
        "Employee_ID": request.form["emp_id"],
        "Name": request.form["name"],
        "Month": request.form["month"],
        "Salary": salary,
        "Bonus": bonus,
        "Total_Paid": salary + bonus
    }])

    df = pd.concat([df, new], ignore_index=True)
    df.to_excel(PAY_FILE, index=False)
    return redirect("/hrm")


@app.route("/evaluate", methods=["POST"])
def evaluate():
    df = pd.read_excel(PERF_FILE)

    new = pd.DataFrame([{
        "Employee_ID": request.form["emp_id"],
        "Name": request.form["name"],
        "Rating": request.form["rating"],
        "Remarks": request.form["remarks"]
    }])

    df = pd.concat([df, new], ignore_index=True)
    df.to_excel(PERF_FILE, index=False)
    return redirect("/hrm")

@app.route("/", methods=["GET", "POST"])
def index():
    if "user" not in session:
        return redirect("/login")

    inventory, alerts = get_inventory_status()
    dss = inventory_sales_dss()
    sales_summary = None

    if request.method == "POST":
        selected_date = request.form["sale_date"]
        sales_summary = sales_summary_by_date(
            pd.to_datetime(selected_date).date()
        )

    return render_template(
        "index.html",
        inventory=inventory.to_dict(orient="records"),
        alerts=alerts.to_dict(orient="records"),
        sales=sales_summary.to_dict(orient="records") if sales_summary is not None else None,
        dss=dss.to_dict(orient="records")
    )


@app.route("/sell", methods=["POST"])
def sell():
    product_id = int(request.form["product_id"])
    quantity = int(request.form["quantity"])

    inventory = pd.read_excel(INVENTORY_FILE)
    sales = pd.read_excel(SALES_FILE)

    current_stock = inventory.loc[
        inventory["Product_ID"] == product_id, "Current_Stock"
    ].values[0]

    if current_stock < quantity:
        return "❌ Insufficient stock"

    inventory.loc[
        inventory["Product_ID"] == product_id, "Current_Stock"
    ] -= quantity

    new_sale = pd.DataFrame([{
        "Product_ID": product_id,
        "Quantity Sold": quantity,
        "Date": date.today()
    }])

    sales = pd.concat([sales, new_sale], ignore_index=True)

    inventory.to_excel(INVENTORY_FILE, index=False)
    sales.to_excel(SALES_FILE, index=False)

    return redirect("/")


@app.route("/add", methods=["POST"])
def add():
    product_id = int(request.form["product_id"])
    quantity = int(request.form["quantity"])

    inventory = pd.read_excel(INVENTORY_FILE)
    inventory.loc[
        inventory["Product_ID"] == product_id, "Current_Stock"
    ] += quantity

    inventory.to_excel(INVENTORY_FILE, index=False)
    return redirect("/")


@app.route("/export_dss")
def export_dss():
    dss = inventory_sales_dss()
    dss.to_excel("dss_output.xlsx", index=False)
    return "✅ DSS Output Exported to dss_output.xlsx"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

