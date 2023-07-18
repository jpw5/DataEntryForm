import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from list import NATIONALITIES
import openpyxl
import os
import sqlite3


def enter_data():
    accepted = accept_var.get()

    if accepted == "Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # Course info
            registration_status = reg_status_var.get()
            num_courses = numcourses_spinbox.get()
            num_semesters = numsemesters_spinbox.get()

            print("First name: ", firstname)
            print("Last name: ", lastname)
            print("Title: ", title)
            print("Age: ", age)
            print("Nationality: ", nationality)
            print("# Courses: ", num_courses)
            print("# Semesters: ", num_semesters)
            print("Registration status:", registration_status)
            print("------------------------------------------")

            # filepath = r'C:\Users\kn849jw\PycharmProjects\Data Entry Form\data.xlsx'
            #
            # if not os.path.exists(filepath):
            #     workbook = openpyxl.Workbook()
            #     sheet = workbook.active
            #     heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
            #                "# Courses", "# Semesters", "Registration status"]
            #     sheet.append(heading)
            #     workbook.save(filepath)
            # workbook = openpyxl.load_workbook(filepath)
            # sheet = workbook.active
            # sheet.append([firstname, lastname, title, age, nationality, num_courses,
            #               num_semesters, registration_status])
            # workbook.save(filepath)

            # Create Table
            conn = sqlite3.connect('data.db')
            table_create_query = '''CREATE TABLE IF NOT EXISTS Student_Data 
                    (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT, 
                    registration_status TEXT, num_courses INT, num_semesters INT)
            '''
            conn.execute(table_create_query)

            # Insert Data
            data_insert_query = '''INSERT INTO Student_Data (firstname, lastname, title, 
                        age, nationality, registration_status, num_courses, num_semesters) VALUES 
                        (?, ?, ?, ?, ?, ?, ?, ?)'''
            data_insert_tuple = (firstname, lastname, title,
                                 age, nationality, registration_status, num_courses, num_semesters)
            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()
            conn.close()


        else:
            tk.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tk.messagebox.showwarning(title="Error", message="You have not accepted the terms")


root = ctk.CTk()
root.title('Data Entry Form')
# ctk.set_appearance_mode('Dark')
# root.geometry('600x500')

frame = tk.Frame(root)
frame.pack()

# Saving User Info
user__info_frame = tk.LabelFrame(frame, text='User Information')
user__info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = ctk.CTkLabel(user__info_frame, text='First Name')
first_name_label.grid(row=0, column=0)
last_name_label = ctk.CTkLabel(user__info_frame, text='Last Name')
last_name_label.grid(row=0, column=1)

first_name_entry = ctk.CTkEntry(user__info_frame)
last_name_entry = ctk.CTkEntry(user__info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = ctk.CTkLabel(user__info_frame, text='Title')
title_combobox = ctk.CTkComboBox(user__info_frame, values=["", "Mr.", "Mrs", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = ctk.CTkLabel(user__info_frame, text='Age')
age_spinbox = ctk.CTkEntry(user__info_frame)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = ctk.CTkLabel(user__info_frame, text="Nationality")
nationality_combobox = ctk.CTkComboBox(user__info_frame, values=NATIONALITIES)
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user__info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tk.LabelFrame(frame, text='Registration Information')
courses_frame.grid(row=1, column=0, sticky='news', padx=20, pady=10)

registered_label = ctk.CTkLabel(courses_frame, text='Registration Status')

reg_status_var = tk.StringVar(value="Not Registered")
registered_check = ctk.CTkCheckBox(courses_frame, text="Currently Registered", variable=reg_status_var,
                                   onvalue="Registered", offvalue="Not Registered")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = ctk.CTkLabel(courses_frame, text="# Completed Courses")
numcourses_spinbox = ctk.CTkComboBox(courses_frame, values=['0', '1', '2', '3', '4', '5', '6'])
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = ctk.CTkLabel(courses_frame, text="# Semesters")
numsemesters_spinbox = ctk.CTkComboBox(courses_frame, values=['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tk.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tk.StringVar(value="Not Accepted")
terms_check = ctk.CTkCheckBox(terms_frame, text="I accept the terms and conditions.",
                              variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = ctk.CTkButton(frame, text="Enter data", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

root.mainloop()
