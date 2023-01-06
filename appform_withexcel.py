from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from IPython.display import Image
import ast
import os
import openpyxl

def enter_data():

    # User info
    title = title_combobox.get()
    name = en1.get()
    contact = en3.get()
    
    if name and contact:
        parname = en2.get()
        stream = en4.get()
        # Course info
        registration_status = reg_status_var.get()
        numcourses = numcourses_spinbox.get()
        numsemesters = numsemesters_spinbox.get()
    
        print("Title: ", title)
        print("Name: ", name) 
        print("Guardian name: ", parname)
        print("Contact no.: ", contact)
        print("Stream: ", stream)
        print("Courses: ", numcourses, ", Semesters:", numsemesters)
        print("Registration status", registration_status)
        print("------------------------------------------")
        
        filepath = "D:\AppForm\data.xlsx"
        
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["Title", "Name", "Guardian name", "Contact no.",
                       "Stream", "Courses", "Semesters", "Registration status"]
            sheet.append(heading)
            workbook.save(filepath)
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([title, name, parname, contact, stream,
                      numcourses, numsemesters, registration_status])
        workbook.save(filepath)
         
    else:
        messagebox.showwarning(title = "Error", message = "Name and Contact are required.")


window = Tk()
window.title("Application form")
window.geometry("925x500+300+200")
window.configure(bg="#fff")
window.resizable(False, False)

img = PhotoImage(file='D:\AppForm\login.png')
img = img.zoom(25)
img = img.subsample(32)
Label(window, image = img, bg = "white").place(x = 50, y = 50)

frame = Frame(window, width = 300, height = 350, bg = 'white')
frame.place(x = 480, y = 70)

#Saving user info
user_info_frame = LabelFrame(frame, text = "User information",fg = '#57a1f8',border  = 0,bg = 'white', font =('Microsoft YaHei UI Light',23,'bold'))
user_info_frame.grid(row = 0, column = 0)

lb1= Label(user_info_frame, text="Title :", width=13,border  = 0,bg = 'white', font=("Microsoft YaHei UI Light",12))  
lb1.grid(row = 0, column = 0, padx = 5, pady = 5)
title_combobox = ttk.Combobox(user_info_frame, values=["Mr.", "Ms.", "Mrs.", "Prof.", "Dr.", "Col."], state = "readonly", width = 4)
title_combobox.grid(row = 0, column = 1)  

lb2= Label(user_info_frame, text="Name :", width=13,border  = 1,bg = 'white', font=("Microsoft YaHei UI Light",12))  
lb2.grid(row = 1, column = 0, padx = 5, pady = 5)
en1= Entry(user_info_frame, width = 25, border = 1, fg = 'black', bg = 'white', font = ("Microsoft YaHei UI Light",12))  
en1.grid(row = 1, column = 1, columnspan = 2, padx = 5, pady = 5)
  
lb3= Label(user_info_frame, text="Guardian name :", width=13,border  = 0,bg = 'white', font=("Microsoft YaHei UI Light",12))  
lb3.grid(row = 2, column = 0, padx = 5, pady = 5)
en2= Entry(user_info_frame, width = 25, border = 1, fg = 'black', bg = 'white', font = ("Microsoft YaHei UI Light",12))  
en2.grid(row = 2, column = 1, columnspan = 2, padx = 5, pady = 5)

lb4= Label(user_info_frame, text="Contact no. :", width=13,border  = 0,bg = 'white', font=("Microsoft YaHei UI Light",12))  
lb4.grid(row = 3, column = 0, padx = 5, pady = 5)
en3= Entry(user_info_frame, width = 25, border = 1, fg = 'black', bg = 'white', font = ("Microsoft YaHei UI Light",12))  
en3.grid(row = 3, column = 1, columnspan = 2, padx = 5, pady = 5) 

lb5= Label(user_info_frame, text="Stream :", width=13,border  = 0,bg = 'white', font=("Microsoft YaHei UI Light",12))  
lb5.grid(row = 4, column = 0, padx = 5, pady = 5)
en4= Entry(user_info_frame, width = 25, border = 1, fg = 'black', bg = 'white', font = ("Microsoft YaHei UI Light",12))  
en4.grid(row = 4, column = 1, columnspan = 2, padx = 5, pady = 5) 

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx = 3, pady = 3)

# Saving Course Info
courses_frame = LabelFrame(frame, bg = 'white', border = 0)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = Label(courses_frame, text="Registration Status:", bg = 'white', font = ("Microsoft YaHei UI Light",12))

reg_status_var = StringVar(value="Not Registered")
registered_check = Checkbutton(courses_frame, text="Currently Registered",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered", bg = 'white', font = ("Microsoft YaHei UI Light",12))

registered_label.grid(row=0, column=0)
registered_check.grid(row=0, column=1)

numcourses_label = Label(courses_frame, text= "Completed Courses", bg = 'white', font = ("Microsoft YaHei UI Light",12))
numcourses_spinbox = Spinbox(courses_frame, from_=0, to=100, bg = 'white', font = ("Microsoft YaHei UI Light",12))
numcourses_label.grid(row=1, column=0)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = Label(courses_frame, text="Semesters", bg = 'white', font = ("Microsoft YaHei UI Light",12))
numsemesters_spinbox = Spinbox(courses_frame, from_=0, to=8, bg = 'white', font = ("Microsoft YaHei UI Light",12))
numsemesters_label.grid(row=2, column=0)
numsemesters_spinbox.grid(row=2, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx = 3, pady = 3)

button = Button(frame, text="Enter data", command= enter_data,fg="black",border = 0, bg = 'white', font = ("Microsoft YaHei UI Light",12))
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
