import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import matplotlib.pyplot as plt


# ==========================================================
# FILE SETTINGS
# ==========================================================

FILE_NAME = "student_distribution.xlsx"

COLUMNS = [
    "Student ID",
    "Name",
    "Course",
    "Semester",
    "Gender",
    "Marks",
    "Attendance",
    "Result"
]


# ==========================================================
# CREATE EXCEL FILE IF NOT EXISTS
# ==========================================================

if not os.path.exists(FILE_NAME):

    df = pd.DataFrame(columns=COLUMNS)
    df.to_excel(FILE_NAME, index=False)


# ==========================================================
# READ DATA
# ==========================================================

def read_data():

    return pd.read_excel(FILE_NAME)


# ==========================================================
# CLEAR INPUT FIELDS
# ==========================================================

def clear_fields():

    id_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    course_entry.delete(0, tk.END)
    semester_entry.delete(0, tk.END)
    gender_combo.set("")
    marks_entry.delete(0, tk.END)
    attendance_entry.delete(0, tk.END)

    id_entry.focus()


# ==========================================================
# ADD STUDENT
# ==========================================================

def add_student():

    student_id = id_entry.get().strip()
    name = name_entry.get().strip()
    course = course_entry.get().strip()
    semester = semester_entry.get().strip()
    gender = gender_combo.get()
    marks = marks_entry.get().strip()
    attendance = attendance_entry.get().strip()

    # Check empty fields
    if (
        student_id == ""
        or name == ""
        or course == ""
        or semester == ""
        or gender == ""
        or marks == ""
        or attendance == ""
    ):

        messagebox.showwarning(
            "Missing Data",
            "Please fill all fields."
        )

        return

    try:

        semester = int(semester)
        marks = float(marks)
        attendance = float(attendance)

    except ValueError:

        messagebox.showerror(
            "Invalid Data",
            "Semester, Marks and Attendance must be numbers."
        )

        return

    # Validate marks
    if marks < 0 or marks > 100:

        messagebox.showerror(
            "Invalid Marks",
            "Marks must be between 0 and 100."
        )

        return

    # Validate attendance
    if attendance < 0 or attendance > 100:

        messagebox.showerror(
            "Invalid Attendance",
            "Attendance must be between 0 and 100."
        )

        return

    df = read_data()

    # Check duplicate ID
    if student_id in df["Student ID"].astype(str).values:

        messagebox.showerror(
            "Duplicate ID",
            "Student ID already exists."
        )

        return

    # Calculate result
    if marks >= 40:
        result = "Pass"
    else:
        result = "Fail"

    new_student = {
        "Student ID": student_id,
        "Name": name,
        "Course": course,
        "Semester": semester,
        "Gender": gender,
        "Marks": marks,
        "Attendance": attendance,
        "Result": result
    }

    df = pd.concat(
        [df, pd.DataFrame([new_student])],
        ignore_index=True
    )

    df.to_excel(
        FILE_NAME,
        index=False
    )

    messagebox.showinfo(
        "Success",
        "Student added successfully."
    )

    clear_fields()
    show_students()


# ==========================================================
# UPDATE STUDENT
# ==========================================================

def update_student():

    selected = student_table.selection()

    if not selected:

        messagebox.showwarning(
            "Select Student",
            "Please select a student from the table."
        )

        return

    student_id = id_entry.get().strip()

    if student_id == "":
        messagebox.showwarning(
            "Missing ID",
            "Enter Student ID."
        )
        return

    try:

        semester = int(semester_entry.get())
        marks = float(marks_entry.get())
        attendance = float(attendance_entry.get())

    except ValueError:

        messagebox.showerror(
            "Invalid Data",
            "Enter valid numbers."
        )

        return

    if marks < 0 or marks > 100:

        messagebox.showerror(
            "Invalid Marks",
            "Marks must be between 0 and 100."
        )

        return

    if attendance < 0 or attendance > 100:

        messagebox.showerror(
            "Invalid Attendance",
            "Attendance must be between 0 and 100."
        )

        return

    df = read_data()

    index = df[
        df["Student ID"].astype(str) == student_id
    ].index

    if len(index) == 0:

        messagebox.showerror(
            "Error",
            "Student not found."
        )

        return

    i = index[0]

    df.loc[i, "Name"] = name_entry.get().strip()
    df.loc[i, "Course"] = course_entry.get().strip()
    df.loc[i, "Semester"] = semester
    df.loc[i, "Gender"] = gender_combo.get()
    df.loc[i, "Marks"] = marks
    df.loc[i, "Attendance"] = attendance

    if marks >= 40:
        df.loc[i, "Result"] = "Pass"
    else:
        df.loc[i, "Result"] = "Fail"

    df.to_excel(
        FILE_NAME,
        index=False
    )

    messagebox.showinfo(
        "Success",
        "Student information updated."
    )

    clear_fields()
    show_students()


# ==========================================================
# DELETE STUDENT
# ==========================================================

def delete_student():

    selected = student_table.selection()

    if not selected:

        messagebox.showwarning(
            "Select Student",
            "Please select a student first."
        )

        return

    item = student_table.item(selected[0])

    student_id = str(item["values"][0])

    answer = messagebox.askyesno(
        "Confirm Delete",
        "Are you sure you want to delete this student?"
    )

    if not answer:
        return

    df = read_data()

    df = df[
        df["Student ID"].astype(str) != student_id
    ]

    df.to_excel(
        FILE_NAME,
        index=False
    )

    messagebox.showinfo(
        "Deleted",
        "Student deleted successfully."
    )

    clear_fields()
    show_students()


# ==========================================================
# SEARCH STUDENT
# ==========================================================

def search_student():

    search_id = search_entry.get().strip()

    if search_id == "":

        messagebox.showwarning(
            "Search",
            "Enter Student ID."
        )

        return

    df = read_data()

    result = df[
        df["Student ID"].astype(str) == search_id
    ]

    for row in student_table.get_children():

        student_table.delete(row)

    if result.empty:

        messagebox.showinfo(
            "Search Result",
            "Student not found."
        )

        return

    for _, row in result.iterrows():

        student_table.insert(
            "",
            tk.END,
            values=list(row)
        )


# ==========================================================
# SHOW ALL STUDENTS
# ==========================================================

def show_students():

    for row in student_table.get_children():

        student_table.delete(row)

    df = read_data()

    for _, row in df.iterrows():

        student_table.insert(
            "",
            tk.END,
            values=list(row)
        )


# ==========================================================
# LOAD SELECTED STUDENT
# ==========================================================

def load_student(event):

    selected = student_table.selection()

    if not selected:
        return

    item = student_table.item(selected[0])

    values = item["values"]

    if not values:
        return

    clear_fields()

    id_entry.insert(0, values[0])
    name_entry.insert(0, values[1])
    course_entry.insert(0, values[2])
    semester_entry.insert(0, values[3])
    gender_combo.set(values[4])
    marks_entry.insert(0, values[5])
    attendance_entry.insert(0, values[6])


# ==========================================================
# DASHBOARD
# ==========================================================

def show_dashboard():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Dashboard",
            "No student data available."
        )

        return

    total = len(df)

    passed = len(
        df[df["Result"] == "Pass"]
    )

    failed = len(
        df[df["Result"] == "Fail"]
    )

    average_marks = df["Marks"].mean()

    average_attendance = df["Attendance"].mean()

    courses = df["Course"].nunique()

    dashboard_text = (
        "STUDENT DISTRIBUTION DASHBOARD\n"
        "\n"
        "------------------------------\n"
        f"Total Students       : {total}\n"
        f"Total Courses        : {courses}\n"
        f"Passed Students      : {passed}\n"
        f"Failed Students      : {failed}\n"
        f"Average Marks        : {average_marks:.2f}\n"
        f"Average Attendance   : {average_attendance:.2f}%\n"
        "------------------------------"
    )

    messagebox.showinfo(
        "Dashboard",
        dashboard_text
    )


# ==========================================================
# COURSE DISTRIBUTION CHART
# ==========================================================

def course_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    course_data = df["Course"].value_counts()

    plt.figure(figsize=(8, 5))

    course_data.plot(
        kind="bar"
    )

    plt.title(
        "Student Distribution by Course"
    )

    plt.xlabel(
        "Course"
    )

    plt.ylabel(
        "Number of Students"
    )

    plt.xticks(
        rotation=0
    )

    plt.tight_layout()

    plt.show()


# ==========================================================
# PASS / FAIL CHART
# ==========================================================

def result_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    result_data = df["Result"].value_counts()

    plt.figure(figsize=(6, 6))

    plt.pie(
        result_data,
        labels=result_data.index,
        autopct="%1.1f%%",
        startangle=90
    )

    plt.title(
        "Student Pass / Fail Distribution"
    )

    plt.show()


# ==========================================================
# SEMESTER DISTRIBUTION CHART
# ==========================================================

def semester_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    semester_data = (
        df["Semester"]
        .value_counts()
        .sort_index()
    )

    plt.figure(figsize=(8, 5))

    semester_data.plot(
        kind="bar"
    )

    plt.title(
        "Student Distribution by Semester"
    )

    plt.xlabel(
        "Semester"
    )

    plt.ylabel(
        "Number of Students"
    )

    plt.xticks(
        rotation=0
    )

    plt.tight_layout()

    plt.show()


# ==========================================================
# GENDER DISTRIBUTION CHART
# ==========================================================

def gender_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    gender_data = df["Gender"].value_counts()

    plt.figure(figsize=(6, 6))

    plt.pie(
        gender_data,
        labels=gender_data.index,
        autopct="%1.1f%%",
        startangle=90
    )

    plt.title(
        "Gender Distribution"
    )

    plt.show()


# ==========================================================
# MARKS ANALYSIS
# ==========================================================

def marks_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    plt.figure(figsize=(8, 5))

    plt.hist(
        df["Marks"],
        bins=10
    )

    plt.title(
        "Student Marks Distribution"
    )

    plt.xlabel(
        "Marks"
    )

    plt.ylabel(
        "Number of Students"
    )

    plt.tight_layout()

    plt.show()


# ==========================================================
# ATTENDANCE ANALYSIS
# ==========================================================

def attendance_chart():

    df = read_data()

    if df.empty:

        messagebox.showwarning(
            "Chart",
            "No student data available."
        )

        return

    plt.figure(figsize=(8, 5))

    plt.hist(
        df["Attendance"],
        bins=10
    )

    plt.title(
        "Student Attendance Distribution"
    )

    plt.xlabel(
        "Attendance (%)"
    )

    plt.ylabel(
        "Number of Students"
    )

    plt.tight_layout()

    plt.show()


# ==========================================================
# MAIN WINDOW
# ==========================================================

root = tk.Tk()

root.title(
    "Student Distribution Analysis System"
)

root.geometry(
    "1250x750"
)

root.minsize(
    1100,
    650
)


# ==========================================================
# TITLE
# ==========================================================

title_label = tk.Label(
    root,
    text="Student Distribution Analysis System",
    font=("Arial", 22, "bold")
)

title_label.pack(
    pady=12
)


# ==========================================================
# INPUT FRAME
# ==========================================================

input_frame = tk.LabelFrame(
    root,
    text="Student Information",
    font=("Arial", 12, "bold"),
    padx=10,
    pady=10
)

input_frame.pack(
    fill="x",
    padx=15,
    pady=5
)


# Student ID

tk.Label(
    input_frame,
    text="Student ID"
).grid(
    row=0,
    column=0,
    padx=8,
    pady=7,
    sticky="w"
)

id_entry = tk.Entry(
    input_frame,
    width=25
)

id_entry.grid(
    row=0,
    column=1,
    padx=8,
    pady=7
)


# Name

tk.Label(
    input_frame,
    text="Name"
).grid(
    row=0,
    column=2,
    padx=8,
    pady=7,
    sticky="w"
)

name_entry = tk.Entry(
    input_frame,
    width=25
)

name_entry.grid(
    row=0,
    column=3,
    padx=8,
    pady=7
)


# Course

tk.Label(
    input_frame,
    text="Course"
).grid(
    row=1,
    column=0,
    padx=8,
    pady=7,
    sticky="w"
)

course_entry = tk.Entry(
    input_frame,
    width=25
)

course_entry.grid(
    row=1,
    column=1,
    padx=8,
    pady=7
)


# Semester

tk.Label(
    input_frame,
    text="Semester"
).grid(
    row=1,
    column=2,
    padx=8,
    pady=7,
    sticky="w"
)

semester_entry = tk.Entry(
    input_frame,
    width=25
)

semester_entry.grid(
    row=1,
    column=3,
    padx=8,
    pady=7
)


# Gender

tk.Label(
    input_frame,
    text="Gender"
).grid(
    row=2,
    column=0,
    padx=8,
    pady=7,
    sticky="w"
)

gender_combo = ttk.Combobox(
    input_frame,
    values=["Male", "Female", "Other"],
    state="readonly",
    width=22
)

gender_combo.grid(
    row=2,
    column=1,
    padx=8,
    pady=7
)


# Marks

tk.Label(
    input_frame,
    text="Marks"
).grid(
    row=2,
    column=2,
    padx=8,
    pady=7,
    sticky="w"
)

marks_entry = tk.Entry(
    input_frame,
    width=25
)

marks_entry.grid(
    row=2,
    column=3,
    padx=8,
    pady=7
)


# Attendance

tk.Label(
    input_frame,
    text="Attendance (%)"
).grid(
    row=3,
    column=0,
    padx=8,
    pady=7,
    sticky="w"
)

attendance_entry = tk.Entry(
    input_frame,
    width=25
)

attendance_entry.grid(
    row=3,
    column=1,
    padx=8,
    pady=7
)


# ==========================================================
# BUTTON FRAME
# ==========================================================

button_frame = tk.Frame(root)

button_frame.pack(
    pady=10
)


tk.Button(
    button_frame,
    text="Add Student",
    width=14,
    command=add_student
).grid(
    row=0,
    column=0,
    padx=4
)


tk.Button(
    button_frame,
    text="Update",
    width=14,
    command=update_student
).grid(
    row=0,
    column=1,
    padx=4
)


tk.Button(
    button_frame,
    text="Delete",
    width=14,
    command=delete_student
).grid(
    row=0,
    column=2,
    padx=4
)


tk.Button(
    button_frame,
    text="Clear",
    width=14,
    command=clear_fields
).grid(
    row=0,
    column=3,
    padx=4
)


tk.Button(
    button_frame,
    text="Dashboard",
    width=14,
    command=show_dashboard
).grid(
    row=0,
    column=4,
    padx=4
)


# ==========================================================
# SEARCH FRAME
# ==========================================================

search_frame = tk.Frame(root)

search_frame.pack(
    pady=5
)


tk.Label(
    search_frame,
    text="Search Student ID:"
).pack(
    side=tk.LEFT,
    padx=5
)


search_entry = tk.Entry(
    search_frame,
    width=25
)

search_entry.pack(
    side=tk.LEFT,
    padx=5
)


tk.Button(
    search_frame,
    text="Search",
    width=12,
    command=search_student
).pack(
    side=tk.LEFT,
    padx=5
)


tk.Button(
    search_frame,
    text="Show All",
    width=12,
    command=show_students
).pack(
    side=tk.LEFT,
    padx=5
)


# ==========================================================
# CHART BUTTONS
# ==========================================================

chart_frame = tk.LabelFrame(
    root,
    text="Analysis Charts",
    font=("Arial", 11, "bold")
)

chart_frame.pack(
    fill="x",
    padx=15,
    pady=5
)


tk.Button(
    chart_frame,
    text="Course Distribution",
    command=course_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


tk.Button(
    chart_frame,
    text="Pass / Fail",
    command=result_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


tk.Button(
    chart_frame,
    text="Semester Distribution",
    command=semester_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


tk.Button(
    chart_frame,
    text="Gender Distribution",
    command=gender_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


tk.Button(
    chart_frame,
    text="Marks Analysis",
    command=marks_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


tk.Button(
    chart_frame,
    text="Attendance Analysis",
    command=attendance_chart,
    width=20
).pack(
    side=tk.LEFT,
    padx=8,
    pady=8
)


# ==========================================================
# TABLE FRAME
# ==========================================================

table_frame = tk.Frame(root)

table_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=10
)


# Vertical scrollbar

vertical_scrollbar = ttk.Scrollbar(
    table_frame,
    orient="vertical"
)

vertical_scrollbar.pack(
    side="right",
    fill="y"
)


# Horizontal scrollbar

horizontal_scrollbar = ttk.Scrollbar(
    table_frame,
    orient="horizontal"
)

horizontal_scrollbar.pack(
    side="bottom",
    fill="x"
)


# ==========================================================
# STUDENT TABLE
# ==========================================================

student_table = ttk.Treeview(
    table_frame,
    columns=COLUMNS,
    show="headings",
    yscrollcommand=vertical_scrollbar.set,
    xscrollcommand=horizontal_scrollbar.set
)


vertical_scrollbar.config(
    command=student_table.yview
)

horizontal_scrollbar.config(
    command=student_table.xview
)


for column in COLUMNS:

    student_table.heading(
        column,
        text=column
    )

    student_table.column(
        column,
        width=120,
        anchor="center"
    )


student_table.pack(
    fill="both",
    expand=True
)


# Click table row to load student

student_table.bind(
    "<ButtonRelease-1>",
    load_student
)


# ==========================================================
# FOOTER
# ==========================================================

footer = tk.Label(
    root,
    text="Python | Tkinter | Pandas | OpenPyXL | Matplotlib | Excel",
    font=("Arial", 9)
)

footer.pack(
    pady=5
)


# ==========================================================
# LOAD DATA WHEN PROGRAM STARTS
# ==========================================================

show_students()

id_entry.focus()


# ==========================================================
# START APPLICATION
# ==========================================================

root.mainloop()
