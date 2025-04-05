import cv2
import face_recognition
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import datetime
import numpy as np

EXCEL_FILE = 'attendance.xlsx'
EMPLOYEE_DATA_FILE = 'employee_data.xlsx'

# Load registered faces
def load_registered_faces():
    known_face_encodings = []
    known_face_names = []

    wb = load_workbook(EMPLOYEE_DATA_FILE)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        first_name, last_name, _, face_encoding = row
        full_name = f"{first_name} {last_name}"
        if face_encoding:
            encoding = np.fromstring(face_encoding[1:-1], dtype=float, sep=',')
            known_face_encodings.append(encoding)
            known_face_names.append(full_name)

    return known_face_encodings, known_face_names

# Initialize attendance Excel
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

    col_map = {
        "Clock In": 2,
        "Break Start": 3,
        "Break End": 4,
        "Shift End": 5
    }

    col = col_map.get(action_type)
    if col and ws.cell(row=row_index, column=col).value is None:
        ws.cell(row=row_index, column=col).value = timestamp
        wb.save(EXCEL_FILE)
        messagebox.showinfo("Success", f"{action_type} recorded for {full_name}")
    else:
        messagebox.showinfo("Already Recorded", f"{action_type} already recorded for {full_name}")

# Open camera and detect face
def open_camera_for_recognition(action_type):
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        messagebox.showerror("Error", "Cannot access camera")
        return

    known_encodings, known_names = load_registered_faces()
    if not known_encodings:
        messagebox.showerror("Error", "No registered faces found.")
        return

    found = False

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        for face_encoding in face_encodings:
            matches = face_recognition.compare_faces(known_encodings, face_encoding)
            face_distances = face_recognition.face_distance(known_encodings, face_encoding)

            if True in matches:
                best_index = np.argmin(face_distances)
                name = known_names[best_index]
                update_employee_action(name, action_type)
                found = True
                break

        cv2.imshow("Face Recognition", frame)
        if found or cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

    if not found:
        messagebox.showerror("Not Found", "No matching face found.")

# Tkinter UI
app = tk.Tk()
app.title("Face Recognition Attendance (Unoptimized)")

tk.Button(app, text="Clock In", command=lambda: open_camera_for_recognition("Clock In")).pack(pady=5)
tk.Button(app, text="Break Start", command=lambda: open_camera_for_recognition("Break Start")).pack(pady=5)
tk.Button(app, text="Break End", command=lambda: open_camera_for_recognition("Break End")).pack(pady=5)
tk.Button(app, text="Shift End", command=lambda: open_camera_for_recognition("Shift End")).pack(pady=5)

init_attendance_excel()
app.mainloop()
