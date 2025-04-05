import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import datetime

EXCEL_FILE = 'attendance.xlsx'

# Initialize Excel
def init_attendance_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.save(EXCEL_FILE)
    
    wb = load_workbook(EXCEL_FILE)
    current_date = datetime.datetime.now().strftime('%Y-%m-%d')
    
    if current_date not in wb.sheetnames:
        ws = wb.create_sheet(current_date)
        ws.append(["Employee Name", "Shift Start", "Break Start", "Break End", "Shift End"])
        wb.save(EXCEL_FILE)

# Update employee action
def update_employee_action(full_name, action_type):
    wb = load_workbook(EXCEL_FILE)
    current_date = datetime.datetime.now().strftime('%Y-%m-%d')

    if current_date not in wb.sheetnames:
        ws = wb.create_sheet(current_date)
        ws.append(["Employee Name", "Shift Start", "Break Start", "Break End", "Shift End"])
    else:
        ws = wb[current_date]

    employee_row = None
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=False):
        if row[0].value == full_name:
            employee_row = row
            break

    if not employee_row:
        ws.append([full_name, None, None, None, None])
        row_index = ws.max_row
    else:
        row_index = employee_row[0].row

    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    if action_type == "Clock In":
        ws.cell(row=row_index, column=2).value = timestamp
    elif action_type == "Break Start":
        ws.cell(row=row_index, column=3).value = timestamp
    elif action_type == "Break End":
        ws.cell(row=row_index, column=4).value = timestamp
    elif action_type == "Shift End":
        ws.cell(row=row_index, column=5).value = timestamp

    wb.save(EXCEL_FILE)
    messagebox.showinfo("Success", f"{action_type} recorded for {full_name}")

# Tkinter UI
def handle_action(action_type):
    full_name = name_entry.get()
    if not full_name:
        messagebox.showerror("Error", "Please enter employee name")
        return
    update_employee_action(full_name, action_type)

app = tk.Tk()
app.title("Basic Clock In System")

tk.Label(app, text="Employee Full Name:").pack(pady=5)
name_entry = tk.Entry(app)
name_entry.pack(pady=5)

tk.Button(app, text="Clock In", command=lambda: handle_action("Clock In")).pack(pady=5)
tk.Button(app, text="Break Start", command=lambda: handle_action("Break Start")).pack(pady=5)
tk.Button(app, text="Break End", command=lambda: handle_action("Break End")).pack(pady=5)
tk.Button(app, text="Shift End", command=lambda: handle_action("Shift End")).pack(pady=5)

init_attendance_excel()
app.mainloop()
