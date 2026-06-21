from flask import Flask, render_template, request, redirect, url_for, session, send_file
from datetime import datetime, timedelta
import os
import jdatetime
from openpyxl import Workbook
from io import BytesIO
from functools import wraps
from pymongo import MongoClient
import certifi

app = Flask(__name__)
app.secret_key = "beauty_clinic_secret_key_change_me_2024"

PASSWORD = os.environ.get("APP_PASSWORD", "1234")

# اتصال به MongoDB با تنظیمات SSL
MONGO_URI = os.environ.get("MONGO_URI", "")
client = MongoClient(
   client = MongoClient(
    MONGO_URI,
    tls=True,
    tlsAllowInvalidCertificates=True,
    serverSelectionTimeoutMS=30000,
    connectTimeoutMS=30000,
    socketTimeoutMS=30000,
)
db = client.clinic
customers_collection = db.customers


def get_today_shamsi():
    return jdatetime.date.today().strftime("%Y/%m/%d")

def get_tomorrow_shamsi():
    return (jdatetime.date.today() + timedelta(days=1)).strftime("%Y/%m/%d")

def generate_id():
    last = customers_collection.find_one(sort=[("id", -1)])
    if last:
        return last["id"] + 1
    return 1

def get_all_customers():
    return list(customers_collection.find({}, {"_id": 0}))

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("dashboard"))
        else:
            error = "رمز عبور اشتباه است!"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    customers = get_all_customers()
    today = get_today_shamsi()
    tomorrow = get_tomorrow_shamsi()
    
    total = len(customers)
    today_customers = [c for c in customers if c.get("next_appointment") == today]
    tomorrow_customers = [c for c in customers if c.get("next_appointment") == tomorrow]
    future_count = len([c for c in customers 
                       if c.get("next_appointment") and c["next_appointment"] > today])
    
    return render_template("dashboard.html",
                          today=today,
                          tomorrow=tomorrow,
                          total=total,
                          today_count=len(today_customers),
                          tomorrow_count=len(tomorrow_customers),
                          future_count=future_count,
                          today_customers=today_customers)


@app.route("/add", methods=["GET", "POST"])
@login_required
def add_customer():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        phone = request.form.get("phone", "").strip().replace(" ", "").replace("-", "")
        service = request.form.get("service", "").strip()
        date = request.form.get("next_appointment", "").strip()
        notes = request.form.get("notes", "").strip()
        
        if not name or not phone:
            return render_template("add.html", error="نام و تلفن الزامی است!")
        
        next_date = None
        if date:
            try:
                parts = date.replace("-", "/").split("/")
                year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
                if not (1400 <= year <= 1450 and 1 <= month <= 12 and 1 <= day <= 31):
                    raise ValueError
                next_date = f"{year:04d}/{month:02d}/{day:02d}"
            except:
                return render_template("add.html", error="فرمت تاریخ نادرست است! مثال: 1403/12/15")
        
        customer = {
            "id": generate_id(),
            "name": name,
            "phone": phone,
            "service": service,
            "next_appointment": next_date,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "notes": notes
        }
        customers_collection.insert_one(customer)
        
        return redirect(url_for("customers_list"))
    
    return render_template("add.html")


@app.route("/customers")
@login_required
def customers_list():
    customers = get_all_customers()
    today = get_today_shamsi()
    tomorrow = get_tomorrow_shamsi()
    return render_template("customers.html", 
                          customers=customers,
                          title="لیست مشتریان",
                          today=today,
                          tomorrow=tomorrow)


@app.route("/search", methods=["GET", "POST"])
@login_required
def search():
    if request.method == "POST":
        query = request.form.get("query", "").strip().lower()
        customers = get_all_customers()
        results = [
            c for c in customers
            if query in c["name"].lower() or
               query in c["phone"] or
               query in c.get("service", "").lower()
        ]
        today = get_today_shamsi()
        tomorrow = get_tomorrow_shamsi()
        return render_template("customers.html",
                              customers=results,
                              title=f"نتایج جستجو: {query}",
                              today=today,
                              tomorrow=tomorrow)
    return render_template("search.html")


@app.route("/today")
@login_required
def today_appointments():
    today = get_today_shamsi()
    customers = list(customers_collection.find({"next_appointment": today}, {"_id": 0}))
    return render_template("customers.html",
                          customers=customers,
                          title=f"نوبت‌های امروز ({today})",
                          today=today,
                          tomorrow=get_tomorrow_shamsi())


@app.route("/tomorrow")
@login_required
def tomorrow_appointments():
    tomorrow = get_tomorrow_shamsi()
    customers = list(customers_collection.find({"next_appointment": tomorrow}, {"_id": 0}))
    return render_template("customers.html",
                          customers=customers,
                          title=f"نوبت‌های فردا ({tomorrow})",
                          today=get_today_shamsi(),
                          tomorrow=tomorrow)


@app.route("/all_appointments")
@login_required
def all_appointments():
    today = get_today_shamsi()
    customers = list(customers_collection.find(
        {"next_appointment": {"$gte": today}}, 
        {"_id": 0}
    ).sort("next_appointment", 1))
    return render_template("customers.html",
                          customers=customers,
                          title="همه نوبت‌های آینده",
                          today=today,
                          tomorrow=get_tomorrow_shamsi())


@app.route("/edit/<int:customer_id>", methods=["GET", "POST"])
@login_required
def edit_customer(customer_id):
    customer = customers_collection.find_one({"id": customer_id}, {"_id": 0})
    
    if not customer:
        return redirect(url_for("customers_list"))
    
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        phone = request.form.get("phone", "").strip().replace(" ", "").replace("-", "")
        service = request.form.get("service", "").strip()
        date = request.form.get("next_appointment", "").strip()
        notes = request.form.get("notes", "").strip()
        
        if not name or not phone:
            return render_template("edit.html", customer=customer, error="نام و تلفن الزامی است!")
        
        next_date = None
        if date:
            try:
                parts = date.replace("-", "/").split("/")
                year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
                if not (1400 <= year <= 1450 and 1 <= month <= 12 and 1 <= day <= 31):
                    raise ValueError
                next_date = f"{year:04d}/{month:02d}/{day:02d}"
            except:
                return render_template("edit.html", customer=customer, error="فرمت تاریخ نادرست است!")
        
        customers_collection.update_one(
            {"id": customer_id},
            {"$set": {
                "name": name,
                "phone": phone,
                "service": service,
                "next_appointment": next_date,
                "notes": notes
            }}
        )
        
        return redirect(url_for("customers_list"))
    
    return render_template("edit.html", customer=customer)


@app.route("/delete/<int:customer_id>", methods=["POST"])
@login_required
def delete_customer(customer_id):
    customers_collection.delete_one({"id": customer_id})
    return redirect(url_for("customers_list"))


@app.route("/export")
@login_required
def export_excel():
    customers = get_all_customers()
    if not customers:
        return redirect(url_for("dashboard"))
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    ws.sheet_view.rightToLeft = True
    
    headers = ["شناسه", "نام و نام خانوادگی", "شماره تلفن", "نوع خدمات", "نوبت بعدی", "یادداشت", "تاریخ ثبت"]
    ws.append(headers)
    
    for c in customers:
        ws.append([
            c["id"], c["name"], c["phone"], c["service"],
            c.get("next_appointment") or "—",
            c.get("notes", ""), c.get("created_at", "")
        ])
    
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 20
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"clinic_{get_today_shamsi().replace('/', '_')}.xlsx"
    return send_file(output, 
                    download_name=filename,
                    as_attachment=True,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
