from tkinter import *
from tkinter import messagebox
import os

window = Tk()
window.title("Login form")
window.geometry('340x440')
window.configure(bg='#333333')

def login():
    username = "stcet"
    password = "12345"
    if username_entry.get()==username and password_entry.get()==password:
        messagebox.showinfo(title="Login Success", message="You successfully logged in.")
        os.system('appform_withexcel.py')
    else:
        messagebox.showerror(title="Error", message="Invalid login.")

frame = Frame(bg='#333333')

# Creating widgets
login_label = Label(
    frame, text="Login", bg='#333333', fg="#FF3399", font=("Microsoft YaHei UI Light", 20))
username_label = Label(
    frame, text="Username", bg='#333333', fg="#FFFFFF", font=("Microsoft YaHei UI Light", 16))
username_entry = Entry(frame, font=("Microsoft YaHei UI Light", 12))
password_entry = Entry(frame, show="*", font=("Microsoft YaHei UI Light", 12))
password_label = Label(
    frame, text="Password", bg='#333333', fg="#FFFFFF", font=("Microsoft YaHei UI Light", 16))
login_button = Button(
    frame, text="Login", bg="#FF3399", fg="#FFFFFF", font=("Microsoft YaHei UI Light", 16), command=login)

# Placing widgets on the screen
login_label.grid(row=0, column=0, columnspan=2, sticky="news", pady=5)
username_label.grid(row=1, column=0)
username_entry.grid(row=1, column=1, pady=5)
password_label.grid(row=2, column=0)
password_entry.grid(row=2, column=1, pady=5)
login_button.grid(row=3, column=0, columnspan=2, pady=10)

frame.pack()

window.mainloop()