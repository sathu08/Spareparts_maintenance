import secrets
import sqlite3
import time

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from datetime import datetime
import pandas as pd
import os
import threading
app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
selected_part_name = ''
value_check = False
Date = datetime.today().strftime('%Y-%m-%d')
Time = datetime.today().strftime("%H:%M:%S")
notification = ''
login_details = ""
UPLOAD_FOLDER = 'uploads'  # Directory to save uploaded files
ALLOWED_EXTENSIONS = {'xlsx'}  # Allowed file extension
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def time_count():
    seconds = 2
    while seconds > 0:
        time.sleep(1)
        seconds -= 1


def start_thead():
    thread = threading.Thread(target=time_count)
    thread.daemon = True
    thread.start()



def excel_to_df(path):
    df = pd.read_excel("uploads/current_excel_file.xlsx", sheet_name='Sheet1')
    conn = sqlite3.connect('spare_part_maintenance.db')
    data_title_name = ['Date', 'part_name', 'part_number', 'machine_name', 'supplier', 'Quantity', 'price_inr',
                       'price_usd', 'VPN', 'Category', 'Bin']
    df.columns = data_title_name
    df.to_sql('new_spare_in', conn, if_exists='replace', index=False)
    conn.commit()
    conn.close()


@app.route('/')
@app.route('/login', methods=['GET', 'POST'])
def login():
    global login_details
    if request.method == 'POST':
        admin_username = request.form['admin_username']
        admin_password = request.form['admin_password']
        username = request.form['user_name']
        password = request.form['user_password']
        if admin_username == 'ad' and admin_password == 'ad':
            login_details ="ad"
            return redirect(url_for('home'))
        if username == 's' and password == 's':
            login_details ="s"
            return redirect(url_for('home'))

    return render_template('login.html')


@app.route('/home', methods=['GET', 'POST'])
def home():
    global notification, login_details
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()

    notification = cursor.execute("select VPN from new_spare_in WHERE Quantity=0", ).fetchall()
    if notification:
        notification = len(notification)
    total_number_spares = cursor.execute("select VPN from new_spare_in ", ).fetchall()
    if total_number_spares:
        total_number_spares = len(total_number_spares)
    unit_cost_inr_non_0 = cursor.execute("select price_inr from new_spare_in WHERE price_inr>0 ", ).fetchall()
    if unit_cost_inr_non_0:
        unit_cost_inr_non_0 = len(unit_cost_inr_non_0)
    unit_cost_inr = cursor.execute("select price_usd from new_spare_in where Category = 'Consumable'", ).fetchall()
    if unit_cost_inr:
        unit_cost_inr = sum([int(value[0]) for value in unit_cost_inr])
    unit_cost_usd_non_0 = cursor.execute("select price_usd from new_spare_in WHERE price_usd>0", ).fetchall()
    if unit_cost_usd_non_0:
        unit_cost_usd_non_0 = len(unit_cost_usd_non_0)
    unit_cost_usd = cursor.execute("select price_usd from new_spare_in ", ).fetchall()
    print(unit_cost_usd)
    if unit_cost_usd:
        unit_cost_usd = sum([int(value[0]) for value in unit_cost_usd])
    cursor.close()
    conn.close()
    return render_template('home.html', notification=notification,
                           unit_cost_inr=unit_cost_inr, unit_cost_usd=unit_cost_usd,
                           total_number_spares=total_number_spares,login_details=login_details,
                           unit_cost_inr_non_0=unit_cost_inr_non_0, unit_cost_usd_non_0=unit_cost_usd_non_0)


@app.route('/alert', methods=['GET', 'POST'])
def alert():
    global login_details
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()
    alert_parts = cursor.execute("select VPN,part_name,part_number,machine_name,supplier,Quantity "
                                 "from new_spare_in WHERE Quantity=0", ).fetchall()
    cursor.close()
    conn.close()
    start_thead()
    return render_template('alert.html', alert_parts=alert_parts, login_details=login_details,
                           notification=notification)


@app.route('/new_spare', methods=['GET', 'POST'])
def new_spare():
    global login_details
    if request.method == "POST":
        Date = request.form['Date'].strip()
        VPN = request.form['VPN'].strip()
        part_name = request.form['part_name'].strip()
        part_number = request.form['part_number'].strip()
        machine_name = request.form['machine_name'].strip()
        supplier = request.form['supplier'].strip()
        Quantity = request.form['Quantity'].strip()
        price_inr = request.form['price_inr']
        price_usd = request.form['price_usd']
        Category = request.form['Category']
        Bin = request.form['bin'].strip()
        conn = sqlite3.connect('spare_part_maintenance.db')
        cur = conn.cursor()
        cur.execute('''INSERT INTO new_spare_in (Date, part_number,part_name, machine_name,Quantity, supplier, price_inr, price_usd, 
                            Category , Bin, VPN ) VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                    (Date, part_number, part_name, machine_name, Quantity, supplier, price_inr, price_usd,
                     Category, Bin, VPN))
        conn.commit()
        conn.close()
        start_thead()
    return render_template('new_spare.html',login_details=login_details,
                           notification=notification)


@app.route('/master_sheet', methods=['GET', 'POST'])
def master_sheet():
    global selected_part_name,login_details
    edit_parts = request.args.get('edit_parts')
    if edit_parts:
        selected_part_name = edit_parts
        return redirect(url_for('edit_parts_request_form'))
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()
    machine_parts = cursor.execute("select  machine_name from new_spare_in").fetchall()
    machine_option = request.form.get("Category")
    set_part = set([val[0] for val in machine_parts])
    if machine_option == "all":
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        master_parts = cursor.execute("select  VPN,part_name,part_number,machine_name,supplier,Quantity,price_inr,"
                                      "price_usd,Category,Bin from new_spare_in ",
                                      ).fetchall()
        updated_list_date = []
        not_updated_list_date = []
        for row in [a[0] for a in master_parts]:
            pop1 = cursor.execute("SELECT date FROM In_request_form WHERE VPN=? ORDER BY date DESC LIMIT 1",
                                  (row,)).fetchone()
            pop2 = cursor.execute("select date from new_spare_in WHERE VPN=?", (row,)).fetchone()
            updated_list_date.append(pop1)
            not_updated_list_date.append(pop2)
        cursor.close()
        conn.close()
        start_thead()
        for i in range(len(master_parts)):
            if updated_list_date[i]:
                master_parts[i] = master_parts[i] + updated_list_date[i]
            else:
                master_parts[i] = master_parts[i] + not_updated_list_date[i]
        return render_template('master_sheet.html', values=set_part,login_details=login_details,
                               result_all=master_parts, notification=notification)
    if machine_option and set_part:
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        master_parts = cursor.execute("select  VPN,part_name,part_number,machine_name,supplier,Quantity,price_inr,"
                                      "price_usd,Category,Bin from new_spare_in WHERE machine_name=?",
                                      (machine_option,)).fetchall()
        updated_list_date = []
        not_updated_list_date = []
        for row in [a[0] for a in master_parts]:
            pop1 = cursor.execute("SELECT date FROM In_request_form WHERE VPN=? ORDER BY date DESC LIMIT 1",
                                  (row,)).fetchone()
            pop2 = cursor.execute("select date from new_spare_in WHERE VPN=?", (row,)).fetchone()
            updated_list_date.append(pop1)
            not_updated_list_date.append(pop2)
        cursor.close()
        conn.close()
        start_thead()
        for i in range(len(master_parts)):
            if updated_list_date[i]:
                master_parts[i] = master_parts[i] + updated_list_date[i]
            else:
                master_parts[i] = master_parts[i] + not_updated_list_date[i]
        return render_template('master_sheet.html', values=set_part,login_details=login_details,
                               result_all=master_parts, notification=notification)
    cursor.close()
    conn.close()
    start_thead()
    return render_template('master_sheet.html', values=set_part,login_details=login_details, notification=notification)


@app.route('/edit_parts_request_form', methods=['GET', 'POST'])
def edit_parts_request_form():
    global login_details
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()
    edit_master_parts = cursor.execute("select  VPN,part_name,part_number,machine_name,supplier,Quantity,price_inr,"
                                       "price_usd,Category,Bin from new_spare_in WHERE VPN=?",
                                       (selected_part_name,)).fetchall()
    if request.method == "POST":
        Part_Number = request.form['Part_Number'].strip()
        Machine = request.form['Machine'].strip()
        Supplier = request.form['Supplier'].strip()
        Cost_in_INR = request.form['Cost_in_INR'].strip()
        USD = request.form['USD'].strip()
        Category = request.form['Category'].strip()
        Bin = request.form['Bin'].strip()
        cursor.execute(
            'UPDATE new_spare_in SET part_number = ?, machine_name = ?, supplier = ?, '
            'price_inr = ?, price_usd = ?, Category = ?, Bin = ? WHERE VPN = ?',
            (Part_Number, Machine, Supplier, Cost_in_INR, USD, Category,Bin,
             selected_part_name))
        conn.commit()
        cursor.close()
        conn.close()
        start_thead()
        return redirect(url_for('master_sheet'))
    cursor.close()
    conn.close()
    start_thead()
    return render_template('edit_parts_request_form.html',login_details=login_details,
                           edit_parts_details=edit_master_parts, notification=notification)


@app.route('/search', methods=['GET', 'POST'])
def search():
    global selected_part_name, login_details
    part = request.args.get('part_name')
    add = request.args.get('add_parts')
    if part:
        selected_part_name = part
        return redirect(url_for('out_request_form'))
    if add:
        selected_part_name = add
        return redirect(url_for('in_request_form'))
    if request.method == "POST":
        type_list = 'Date, VPN, part_name ,part_number ,machine_name ,supplier ,Quantity,price_inr , price_usd,category'
        search_name = request.form["searchInput"].strip()
        machine_option = request.form.get("Category")  # Use request.form.get() to avoid KeyError
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        quantity_check = cursor.execute("SELECT {} FROM new_spare_in WHERE part_name LIKE ?".format(type_list),
                                        ('%' + search_name + '%',)).fetchall()
        machine_list = set([value[4] for value in quantity_check])
        if machine_option and machine_list:  # Check if machine_option exists and machine_list is not empty
            machine_option = machine_option.strip()
            quantity_check = cursor.execute(
                "SELECT {} FROM new_spare_in WHERE part_name LIKE ? and machine_name =?".format(type_list),
                ('%' + search_name + '%', machine_option)).fetchall()
            cursor.close()
            conn.close()
            return render_template('search.html', result_all=quantity_check, machine_list=machine_list,
                                   search_value=search_name, login_details=login_details, notification=notification)
        cursor.close()
        conn.close()
        return render_template('search.html', result_all=quantity_check, machine_list=machine_list,
                               search_value=search_name,login_details=login_details,  notification=notification)
    return render_template('search.html',login_details=login_details,  notification=notification)


@app.route('/out_request_form', methods=['GET', 'POST'])
def out_request_form():
    global Date, Time, login_details
    if request.method == "POST":
        Employee_name = request.form["EmployeeId"].strip()
        Part_number = selected_part_name
        Line = request.form.get("Line")
        print(Line)
        Quantity = request.form["Quantity"].strip()
        conn = sqlite3.connect('spare_part_maintenance.db')
        cur = conn.cursor()
        data = cur.execute(
            "select Quantity,Bin,part_name,part_number,machine_name,supplier from new_spare_in where VPN =?",
            (Part_number,)).fetchone()
        if int(data[0]) >= int(Quantity):
            reduced_data = int(data[0]) - int(Quantity)
            cur.execute('UPDATE new_spare_in SET Quantity = ? WHERE VPN = ?', (reduced_data, Part_number))
            cur.execute('''INSERT INTO out_request_form (
                                                date,employee_name, VPN, line, quantity,part_name,part_number,machine_name,supplier,time)
                                                 VALUES (?,?,?,?,?,?,?,?,?,?)''',
                        (Date, Employee_name, Part_number, Line,
                         Quantity, data[2], data[3], data[4],
                         data[5], Time))
            conn.commit()
            conn.close()
            flash(f"Location : {data[1]}", 'success')
            return render_template('out_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                                   login_details=login_details, notification=notification)
        flash("Enter the correct Quantity", 'error')
        return render_template('out_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                               login_details=login_details, notification=notification)
    return render_template('out_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                           login_details=login_details, notification=notification)


@app.route('/in_request_form', methods=['GET', 'POST'])
def in_request_form():
    global Date, Time,login_details
    if request.method == "POST":
        Employee_name = request.form["Name"].strip()
        Part_number = selected_part_name
        Po_number = request.form["PO_number"].strip()
        Invoice_number = request.form["Invoice_number"].strip()
        Quantity = request.form["received_quantity"].strip()
        conn = sqlite3.connect('spare_part_maintenance.db')
        cur = conn.cursor()
        data = cur.execute(
            "select Quantity,Bin,part_name,part_number,machine_name,supplier,price_inr,price_usd from new_spare_in where VPN =?",
            (Part_number,)).fetchone()
        if int(Quantity):
            add_data = int(data[0]) + int(Quantity)
            cur.execute('UPDATE new_spare_in SET Quantity = ? WHERE VPN = ?', (add_data, Part_number))
            cur.execute('''INSERT INTO in_request_form (
                                date,PO_number, employee_name,part_name, part_number ,supplier,machine_name,VPN,
                                received_quantity, Invoice_number,time,cost_in_INR,cost_in_USD,bin)
                                 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                        (Date, Po_number, Employee_name, data[2], Part_number, data[5], data[4], Part_number,
                         Quantity, Invoice_number, Time, data[6], data[7], data[1]))
            conn.commit()
            conn.close()
            flash(f"Location : {data[1]}", 'success')
            return render_template('in_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                                   login_details=login_details, notification=notification)
        flash("Enter the correct Quantity", 'error')
        return render_template('in_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                               login_details=login_details, notification=notification)
    return render_template('in_request_form.html', date=Date, current_selected_partnum=selected_part_name,
                           login_details=login_details, notification=notification)


@app.route('/spare_out_history', methods=['GET', 'POST'])
def spare_out_history():
    global login_details
    current_date = datetime.now()
    from_date = current_date.replace(day=1)
    from_date = from_date.strftime("%Y-%m-%d")
    to_date = current_date.strftime("%Y-%m-%d")
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()
    data = cursor.execute("select VPN from out_request_form WHERE date BETWEEN ? AND ?",
                          (from_date, to_date)).fetchall()
    total_data_hist = cursor.execute("SELECT * FROM out_request_form WHERE date BETWEEN ? AND ?",
                                     (from_date, to_date)).fetchall()
    current_stack = []
    next_formate_details = []
    for part_data in data:
        remaining_data_hist = cursor.execute("SELECT Quantity FROM new_spare_in WHERE VPN=?",
                                             (part_data[0],)).fetchall()
        for i in remaining_data_hist:
            current_stack.append(i[0])
    for j in total_data_hist:
        next_formate_details.append([j[4], j[9], j[0], j[5], j[1], j[7], j[3], j[2], j[8]])
    if request.method == 'POST':
        from_date = request.form['from_date']
        to_date = request.form['to_date']
        next_formate_details = []
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        data = cursor.execute("select VPN from out_request_form WHERE date BETWEEN ? AND ?",
                              (from_date, to_date)).fetchall()
        total_data_hist = cursor.execute("SELECT * FROM out_request_form WHERE date BETWEEN ? AND ?",
                                         (from_date, to_date)).fetchall()
        for part_data in data:
            remaining_data_hist = cursor.execute("SELECT Quantity "
                                                 "FROM new_spare_in WHERE VPN=?", (part_data[0],)).fetchall()
            for i in remaining_data_hist:
                current_stack.append(i[0])
        for j in total_data_hist:
            next_formate_details.append([j[4], j[9], j[0], j[5], j[1], j[7], j[3], j[2], j[8]])
        return render_template('spare_out_history.html', result_all=next_formate_details,
                               stack=current_stack, from_date=from_date, login_details=login_details, to_date=to_date, notification=notification)
    return render_template('spare_out_history.html', result_all=next_formate_details,
                           stack=current_stack, from_date=from_date,login_details=login_details, to_date=to_date, notification=notification)


@app.route('/spare_in_history', methods=['GET', 'POST'])
def spare_in_history():
    global login_details
    current_date = datetime.now()
    from_date = current_date.replace(day=1)
    from_date = from_date.strftime("%Y-%m-%d")
    to_date = current_date.strftime("%Y-%m-%d")
    conn = sqlite3.connect('spare_part_maintenance.db')
    cursor = conn.cursor()
    data = cursor.execute("select VPN from In_request_form WHERE date BETWEEN ? AND ?",
                          (from_date, to_date)).fetchall()
    total_data_hist = cursor.execute("SELECT * FROM In_request_form WHERE date BETWEEN ? AND ?",
                                     (from_date, to_date)).fetchall()
    current_stack = []
    next_formate_details = []
    for part_data in data:
        remaining_data_hist = cursor.execute("SELECT Quantity "
                                             "FROM new_spare_in WHERE VPN=?", (part_data[0],)).fetchall()
        for i in remaining_data_hist:
            current_stack.append(i[0])
    for j in total_data_hist:
        next_formate_details.append([j[0], j[12], j[9], j[1], j[2], j[3], j[5], j[4], j[6], j[10], j[7], j[8], j[13]])
    if request.method == 'POST':
        from_date = request.form['from_date']
        to_date = request.form['to_date']
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        data = cursor.execute("select VPN from In_request_form WHERE date BETWEEN ? AND ?",
                              (from_date, to_date)).fetchall()
        total_data_hist = cursor.execute("SELECT * FROM In_request_form WHERE date BETWEEN ? AND ?",
                                         (from_date, to_date)).fetchall()
        current_stack = []
        next_formate_details = []
        for part_data in data:
            remaining_data_hist = cursor.execute("SELECT Quantity "
                                                 "FROM new_spare_in WHERE VPN=?", (part_data[0],)).fetchall()
            for i in remaining_data_hist:
                current_stack.append(i[0])
        for j in total_data_hist:
            next_formate_details.append(
                [j[0], j[12], j[9], j[1], j[2], j[3], j[5], j[4], j[6], j[10], j[7], j[8], j[13]])
        return render_template('spare_in_history.html', result_all=next_formate_details,
                               stack=current_stack, from_date=from_date, login_details=login_details, to_date=to_date, notification=notification)
    return render_template('spare_in_history.html', result_all=next_formate_details,
                           stack=current_stack, from_date=from_date, login_details=login_details, to_date=to_date, notification=notification)


@app.route('/consumption_of_spare', methods=['GET', 'POST'])
def consumption_of_spare():
    global login_details
    current_date = datetime.now()
    from_date = current_date.replace(day=1)
    from_date = from_date.strftime("%Y-%m-%d")
    to_date = current_date.strftime("%Y-%m-%d")
    if request.method == 'POST':
        consumption = request.form.get('Category')
        if consumption == 'purchased_quantity':
            print("purchased_quantity")
            from_date = request.form['from_date']
            to_date = request.form['to_date']
            cons_list_data = "part_number, part_name, received_quantity, machine_name"
            conn = sqlite3.connect('spare_part_maintenance.db')
            cursor = conn.cursor()
            pop = cursor.execute(f'SELECT {cons_list_data} FROM In_request_form WHERE date BETWEEN ? AND ?',
                                 (from_date, to_date)).fetchall()
            list_data = []
            for row in [a[0] for a in pop]:
                pop1 = cursor.execute(f'SELECT Quantity FROM new_spare_in where VPN=?', (row,)).fetchone()
                list_data.append(pop1)
            conn.close()
            if list_data and pop:
                for i in range(len(pop)):
                    pop[i] = pop[i] + list_data[i]
                result_dict = {}
                for row in pop:
                    part_number, part_name, received_quantity, machine_name, current_stock = row
                    key = (part_number, part_name, machine_name, current_stock)
                    if key in result_dict:
                        result_dict[key] += int(received_quantity)
                    else:
                        result_dict[key] = int(received_quantity)
                result_list = [(received_quantity, part_number, part_name, machine_name, str(current_stock)) for
                               (part_number, part_name, machine_name, current_stock), received_quantity in
                               result_dict.items()]
                return render_template('consumption_of_spare.html', from_date=from_date,
                                       consumption=consumption, login_details=login_details, to_date=to_date,
                                       notification=notification, result_all=sorted(result_list)[::-1])
        elif consumption == 'consumed_quantity':
            print("consumed_quantity")
            from_date = request.form['from_date']
            to_date = request.form['to_date']
            cons_list_data = "VPN, part_name, quantity, machine_name"
            conn = sqlite3.connect('spare_part_maintenance.db')
            cursor = conn.cursor()
            pop = cursor.execute(f'SELECT {cons_list_data} FROM out_request_form WHERE date BETWEEN ? AND ?',
                                 (from_date, to_date)).fetchall()
            list_data = []
            for row in [a[0] for a in pop]:
                pop1 = cursor.execute(f'SELECT Quantity FROM new_spare_in where VPN=?', (row,)).fetchone()
                list_data.append(pop1)
            conn.close()
            if list_data and pop:
                for i in range(len(pop)):
                    pop[i] = pop[i] + list_data[i]
                result_dict = {}
                for row in pop:
                    part_number, part_name, received_quantity, machine_name, current_stock = row
                    key = (part_number, part_name, machine_name, current_stock)
                    if key in result_dict:
                        result_dict[key] += int(received_quantity)
                    else:
                        result_dict[key] = int(received_quantity)
                result_list = [(received_quantity, part_number, part_name, machine_name, str(current_stock)) for
                               (part_number, part_name, machine_name, current_stock), received_quantity in
                               result_dict.items()]
                return render_template('consumption_of_spare.html', from_date=from_date,
                                       consumption=consumption, login_details=login_details, to_date=to_date,
                                   notification=notification, result_all=sorted(result_list)[::-1])
        return render_template('consumption_of_spare.html', consumption=consumption,
                               from_date=from_date, to_date=to_date, login_details=login_details)
    return render_template('consumption_of_spare.html', notification=notification,
                           from_date=from_date, to_date=to_date, login_details=login_details)


@app.route('/submit', methods=['POST'])
def submit():
    user_data = request.form['user_data']
    if user_data:
        conn = sqlite3.connect('spare_part_maintenance.db')
        cursor = conn.cursor()
        path = cursor.execute('select * from new_spare_in').fetchall()
        in_request_path = cursor.execute('select * from In_request_form').fetchall()
        out_request_path = cursor.execute('select * from out_request_form').fetchall()
        for data in path:
            cursor.execute("DELETE FROM new_spare_in WHERE Date=?", (data[0],))
            conn.commit()
        for data in in_request_path:
            cursor.execute("DELETE FROM In_request_form WHERE Date=?", (data[0],))
            conn.commit()
        for data in out_request_path:
            cursor.execute("DELETE FROM out_request_form WHERE Date=?", (data[4],))
            conn.commit()
        conn.close()
    print("User data:", user_data)
    return "Data received: " + user_data


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file and allowed_file(file.filename):
        filename = "current_excel_file.xlsx"
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_to_df(path)
        return render_template('master_sheet.html')
    else:
        return render_template('master_sheet.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
