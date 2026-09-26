import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
import os

# ==========================================
# EXCEL SETTINGS
# ==========================================

FILE_NAME = "student_records.xlsx"

HEADERS = [
    "Student ID",
    "Student Name",
    "Gender",
    "DOB",
    "Course",
    "Year",
    "Phone",
    "Email",
    "Address"
]


# ==========================================
# CREATE EXCEL FILE
# ==========================================

def initialize_excel():

    if not os.path.exists(FILE_NAME):

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Student Records"

        for column, header in enumerate(HEADERS, start=1):
            sheet.cell(row=1, column=column).value = header

        workbook.save(FILE_NAME)


# ==========================================
# LOGIN WINDOW
# ==========================================

def login_window():

    login = tk.Tk()
    login.title("Student Registration System - Login")
    login.geometry("450x400")
    login.resizable(False, False)

    tk.Label(
        login,
        text="STUDENT REGISTRATION SYSTEM",
        font=("Arial", 18, "bold")
    ).pack(pady=35)

    frame = tk.Frame(login)
    frame.pack()

    tk.Label(
        frame,
        text="Username",
        font=("Arial", 12)
    ).grid(row=0, column=0, padx=10, pady=12)

    username_entry = tk.Entry(
        frame,
        width=25,
        font=("Arial", 12)
    )
    username_entry.grid(row=0, column=1)

    tk.Label(
        frame,
        text="Password",
        font=("Arial", 12)
    ).grid(row=1, column=0, padx=10, pady=12)

    password_entry = tk.Entry(
        frame,
        width=25,
        show="*",
        font=("Arial", 12)
    )
    password_entry.grid(row=1, column=1)

    def validate_login():

        username = username_entry.get()
        password = password_entry.get()

        if username == "admin" and password == "admin123":

            login.destroy()
            dashboard()

        else:

            messagebox.showerror(
                "Login Failed",
                "Invalid Username or Password!"
            )

    tk.Button(
        login,
        text="LOGIN",
        width=15,
        bg="green",
        fg="white",
        font=("Arial", 11, "bold"),
        command=validate_login
    ).pack(pady=25)

    tk.Label(
        login,
        text="Username: admin   Password: admin123",
        fg="gray"
    ).pack()

    login.mainloop()


# ==========================================
# DASHBOARD
# ==========================================

def dashboard():

    root = tk.Tk()
    root.title("Student Registration System")
    root.geometry("1250x750")

    # ======================================
    # TITLE
    # ======================================

    tk.Label(
        root,
        text="STUDENT REGISTRATION SYSTEM",
        font=("Arial", 24, "bold")
    ).pack(pady=15)

    # ======================================
    # FORM FRAME
    # ======================================

    form_frame = tk.LabelFrame(
        root,
        text="Student Registration Form",
        font=("Arial", 12, "bold"),
        padx=15,
        pady=10
    )

    form_frame.pack(
        fill="x",
        padx=20,
        pady=5
    )

    # ======================================
    # VARIABLES
    # ======================================

    student_id = tk.StringVar()
    student_name = tk.StringVar()
    gender = tk.StringVar()
    dob = tk.StringVar()
    course = tk.StringVar()
    year = tk.StringVar()
    phone = tk.StringVar()
    email = tk.StringVar()
    address = tk.StringVar()

    # ======================================
    # FORM FIELDS
    # ======================================

    tk.Label(form_frame, text="Student ID").grid(
        row=0, column=0, padx=5, pady=8
    )

    tk.Entry(
        form_frame,
        textvariable=student_id,
        width=22
    ).grid(row=0, column=1)

    tk.Label(form_frame, text="Student Name").grid(
        row=0, column=2, padx=5
    )

    tk.Entry(
        form_frame,
        textvariable=student_name,
        width=22
    ).grid(row=0, column=3)

    tk.Label(form_frame, text="Gender").grid(
        row=0, column=4, padx=5
    )

    gender_box = ttk.Combobox(
        form_frame,
        textvariable=gender,
        values=["Male", "Female", "Other"],
        width=19,
        state="readonly"
    )

    gender_box.grid(row=0, column=5)

    tk.Label(form_frame, text="Date of Birth").grid(
        row=1, column=0, padx=5, pady=8
    )

    tk.Entry(
        form_frame,
        textvariable=dob,
        width=22
    ).grid(row=1, column=1)

    tk.Label(form_frame, text="Course").grid(
        row=1, column=2, padx=5
    )

    tk.Entry(
        form_frame,
        textvariable=course,
        width=22
    ).grid(row=1, column=3)

    tk.Label(form_frame, text="Year").grid(
        row=1, column=4, padx=5
    )

    year_box = ttk.Combobox(
        form_frame,
        textvariable=year,
        values=[
            "First Year",
            "Second Year",
            "Third Year",
            "Fourth Year"
        ],
        width=19,
        state="readonly"
    )

    year_box.grid(row=1, column=5)

    tk.Label(form_frame, text="Phone").grid(
        row=2, column=0, padx=5, pady=8
    )

    tk.Entry(
        form_frame,
        textvariable=phone,
        width=22
    ).grid(row=2, column=1)

    tk.Label(form_frame, text="Email").grid(
        row=2, column=2, padx=5
    )

    tk.Entry(
        form_frame,
        textvariable=email,
        width=22
    ).grid(row=2, column=3)

    tk.Label(form_frame, text="Address").grid(
        row=2, column=4, padx=5
    )

    tk.Entry(
        form_frame,
        textvariable=address,
        width=22
    ).grid(row=2, column=5)

    # ======================================
    # CLEAR FORM
    # ======================================

    def clear_fields():

        student_id.set("")
        student_name.set("")
        gender.set("")
        dob.set("")
        course.set("")
        year.set("")
        phone.set("")
        email.set("")
        address.set("")

        for item in tree.selection():
            tree.selection_remove(item)

    # ======================================
    # VIEW RECORDS
    # ======================================

    def view_records():

        for item in tree.get_children():
            tree.delete(item)

        workbook = load_workbook(FILE_NAME)
        sheet = workbook.active

        for row in sheet.iter_rows(
            min_row=2,
            values_only=True
        ):

            tree.insert(
                "",
                tk.END,
                values=row
            )

        workbook.close()

    # ======================================
    # ADD STUDENT
    # ======================================

    def add_record():

        if student_id.get() == "" or student_name.get() == "":

            messagebox.showwarning(
                "Warning",
                "Student ID and Student Name are required."
            )
            return

        workbook = load_workbook(FILE_NAME)
        sheet = workbook.active

        for row in sheet.iter_rows(
            min_row=2,
            values_only=True
        ):

            if str(row[0]) == student_id.get():

                messagebox.showerror(
                    "Error",
                    "Student ID already exists!"
                )

                workbook.close()
                return

        sheet.append([
            student_id.get(),
            student_name.get(),
            gender.get(),
            dob.get(),
            course.get(),
            year.get(),
            phone.get(),
            email.get(),
            address.get()
        ])

        workbook.save(FILE_NAME)
        workbook.close()

        messagebox.showinfo(
            "Success",
            "Student registered successfully!"
        )

        clear_fields()
        view_records()

    # ======================================
    # SELECT RECORD
    # ======================================

    def select_record(event):

        selected = tree.focus()

        if not selected:
            return

        values = tree.item(
            selected,
            "values"
        )

        if values:

            student_id.set(values[0])
            student_name.set(values[1])
            gender.set(values[2])
            dob.set(values[3])
            course.set(values[4])
            year.set(values[5])
            phone.set(values[6])
            email.set(values[7])
            address.set(values[8])

    # ======================================
    # UPDATE RECORD
    # ======================================

    def update_record():

        if student_id.get() == "":

            messagebox.showwarning(
                "Warning",
                "Select a student record first."
            )
            return

        workbook = load_workbook(FILE_NAME)
        sheet = workbook.active

        found = False

        for row in sheet.iter_rows(min_row=2):

            if str(row[0].value) == student_id.get():

                row[1].value = student_name.get()
                row[2].value = gender.get()
                row[3].value = dob.get()
                row[4].value = course.get()
                row[5].value = year.get()
                row[6].value = phone.get()
                row[7].value = email.get()
                row[8].value = address.get()

                found = True
                break

        workbook.save(FILE_NAME)
        workbook.close()

        if found:

            messagebox.showinfo(
                "Success",
                "Student record updated successfully!"
            )

            clear_fields()
            view_records()

        else:

            messagebox.showerror(
                "Error",
                "Student record not found!"
            )

    # ======================================
    # DELETE RECORD
    # ======================================

    def delete_record():

        if student_id.get() == "":

            messagebox.showwarning(
                "Warning",
                "Select a student record first."
            )
            return

        answer = messagebox.askyesno(
            "Delete",
            "Do you want to delete this student?"
        )

        if not answer:
            return

        workbook = load_workbook(FILE_NAME)
        sheet = workbook.active

        found = False

        for row in range(2, sheet.max_row + 1):

            if str(sheet.cell(row, 1).value) == student_id.get():

                sheet.delete_rows(row, 1)

                found = True
                break

        workbook.save(FILE_NAME)
        workbook.close()

        if found:

            messagebox.showinfo(
                "Success",
                "Student deleted successfully!"
            )

            clear_fields()
            view_records()

        else:

            messagebox.showerror(
                "Error",
                "Student record not found!"
            )

    # ======================================
    # LOGOUT
    # ======================================

    def logout():

        answer = messagebox.askyesno(
            "Logout",
            "Do you want to logout?"
        )

        if answer:

            root.destroy()
            login_window()

    # ======================================
    # BUTTON FRAME
    # ======================================

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    tk.Button(
        button_frame,
        text="ADD STUDENT",
        width=15,
        bg="green",
        fg="white",
        font=("Arial", 10, "bold"),
        command=add_record
    ).grid(row=0, column=0, padx=5)

    tk.Button(
        button_frame,
        text="VIEW RECORDS",
        width=15,
        command=view_records
    ).grid(row=0, column=1, padx=5)

    tk.Button(
        button_frame,
        text="UPDATE",
        width=15,
        bg="orange",
        command=update_record
    ).grid(row=0, column=2, padx=5)

    tk.Button(
        button_frame,
        text="DELETE",
        width=15,
        bg="red",
        fg="white",
        command=delete_record
    ).grid(row=0, column=3, padx=5)

    tk.Button(
        button_frame,
        text="CLEAR",
        width=15,
        command=clear_fields
    ).grid(row=0, column=4, padx=5)

    tk.Button(
        button_frame,
        text="LOGOUT",
        width=15,
        bg="darkblue",
        fg="white",
        command=logout
    ).grid(row=0, column=5, padx=5)

    # ======================================
    # TABLE
    # ======================================

    table_frame = tk.Frame(root)
    table_frame.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=10
    )

    columns = (
        "ID",
        "Name",
        "Gender",
        "DOB",
        "Course",
        "Year",
        "Phone",
        "Email",
        "Address"
    )

    tree = ttk.Treeview(
        table_frame,
        columns=columns,
        show="headings"
    )

    for column in columns:

        tree.heading(
            column,
            text=column
        )

        tree.column(
            column,
            width=120
        )

    # Scrollbars

    vertical_scroll = ttk.Scrollbar(
        table_frame,
        orient="vertical",
        command=tree.yview
    )

    horizontal_scroll = ttk.Scrollbar(
        table_frame,
        orient="horizontal",
        command=tree.xview
    )

    tree.configure(
        yscrollcommand=vertical_scroll.set,
        xscrollcommand=horizontal_scroll.set
    )

    tree.pack(
        side="left",
        fill="both",
        expand=True
    )

    vertical_scroll.pack(
        side="right",
        fill="y"
    )

    horizontal_scroll.pack(
        side="bottom",
        fill="x"
    )

    tree.bind(
        "<ButtonRelease-1>",
        select_record
    )

    # Load existing records
    view_records()

    root.mainloop()


# ==========================================
# START PROGRAM
# ==========================================

initialize_excel()
login_window()
