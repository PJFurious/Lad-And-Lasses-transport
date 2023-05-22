import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import BackEnd as pd
import threading
import os
import signal


def Admin_open():
    window.withdraw()
    
    def save_vehicle():
        reg = reg_input_a.get().strip().upper()
        if not reg:
            messagebox.showerror("Error", "Please add vehicle registration")
            reg_input_a.focus_set()
            return
        pd.add_vehicle(reg)

        reg_input_d.config(values=pd.read_vehicles())
        reg_input_a.delete(0, tk.END)
        input2.config(values=pd.read_vehicles())
        messagebox.showinfo("Done", f"The vehicle ({reg}) has been added")


    def delete_vehicle():
        reg = reg_input_d.get().strip().upper()
        response = messagebox.askokcancel("Warning", f"Are you sure you want to delete the this vehicle? ({reg})")

        if response:
            if not reg:
                messagebox.showerror("Error", "Please select a vehicle registration")
                reg_input_d.focus_set()
                return
            
            
            pd.delete_vehicle(reg)
            reg_input_d.config(values=pd.read_vehicles())
            input2.config(values=pd.read_vehicles())
            input2.set('')
            reg_input_d.set('')
    
    def close():
        root.destroy()
        window.deiconify()

    root = tk.Tk()
    root.title("Admin Form")
    root.resizable(width=False, height=False)
    root.geometry("260x250")
    root.attributes("-topmost", True)
    root.protocol("WM_DELETE_WINDOW", close)


    vehicle_heading = ttk.Label(root, text="Add or remove vehicle", font=("Helvetica", 12, "underline"))
    vehicle_heading.grid(row=0, columnspan=2, pady=10, padx=10)

    add_vehicle_L = ttk.Label(root, text="Add vehicle registration:", font=("Helvetica", 9))
    add_vehicle_L.grid(row=1, column=0)
    reg_input_a = tk.Entry(root, relief="solid", width=21)
    reg_input_a.grid(row=2, column=0)
    add_vehicle_B = ttk.Button(root, text="Add", command=save_vehicle)
    add_vehicle_B.grid(row=2, column=1)

    add_vehicle_L = ttk.Label(root, text="Select vehicle to be removed:", font=("Helvetica", 9))
    add_vehicle_L.grid(row=3, column=0)
    reg_input_d = ttk.Combobox(root, values=pd.read_vehicles(), state="readonly", width=18)
    reg_input_d.grid(row=4, column=0, padx= 10)
    add_vehicle_B = ttk.Button(root, text="Delete", command=delete_vehicle)
    add_vehicle_B.grid(row=4, column=1, padx=10)
    
    select_ex = ttk.Label(root, text="Set Excel Spreadsheet", font=("Helvetica", 12, "underline"))
    select_ex.grid(row=5, columnspan=2, pady=10, padx=10)
    select_path = ttk.Button(root, text="Select Excel Spreadsheet", command=pd.set_path)
    select_path.grid(row=6, columnspan=2)

    close_w = ttk.Button(root, text="Close", command=close)
    close_w.grid(row=7, column=1, pady=15)
    
    root.mainloop()


def add_Heading():
    month = input1.get().strip()
    year = input3.get().strip().upper()
    if not month:
        messagebox.showerror("Error", "Please select a month")
        input1.focus_set()
        return
    if not year:
        messagebox.showerror("Error", "Please enter a year")
        input3.focus_set()
        return
    year = int(year)
    if year > current_year or (year == current_year and months.index(input1.get())+2 > current_month):
        messagebox.showerror("Error", "Please enter a valid date")
    else:
        pd.add_Heading(year, month)


def submit():
    month = input1.get().strip()
    reg = input2.get().strip().upper()
    year = input3.get().strip().upper()
    if not month:
        input1.focus_set()
        messagebox.showerror("Error", "Please select a month")
        return
    if not year:
        messagebox.showerror("Error", "Please enter a year")
        input3.focus_set()
        return
    if not reg:
        messagebox.showerror("Error", "Please enter a registration")
        input2.focus_set()
        return
    year = int(year)
    if year > current_year or (year == current_year and months.index(input1.get())+2 > current_month):
        messagebox.showerror("Error", "Please enter a valid date")
    else:
        threading.Thread(target=pd.ScrapeData(year, (months.index(input1.get()) + 1), reg)).start()

def submit_all():
    month = input1.get().strip()
    reg = input2.get().strip().upper()
    year = input3.get().strip().upper()

    if not month:
        input1.focus_set()
        messagebox.showerror("Error", "Please select a month")
        return
    if not year:
        messagebox.showerror("Error", "Please enter a year")
        input3.focus_set()
        return
    threading.Thread(target=pd.submit_all, args=(reg, month, year, (months.index(input1.get()) + 1))).start()

def stop_process():
    if messagebox.askquestion("Warning", f"Are you sure you want to kill all processes?") == "yes":
        os.kill(os.getpid(), signal.SIGINT)



now = datetime.now()
current_year = now.year
current_month = now.month

window = tk.Tk()
window.title("Lad & Lasses")
window.resizable(width=False, height=False)
window.geometry("300x218+50+50")
window.attributes("-topmost", True)
window.config(background="white")

style = ttk.Style()
style.configure('TButton', width = 20, background="white")
style.configure('TLabel', background="white")

months = ("January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December")
label1 = ttk.Label(window, text="Select a month:")
label1.grid(row=1, column=0, pady=5, padx=10)
input1 = ttk.Combobox(window, values=months, state="readonly", width=18)
input1.grid(row=1, column=1)


label3 = ttk.Label(window, text="Year:")
label3.grid(row=2, column=0, pady=5, padx=10)
input3 = tk.Entry(window, relief="solid", width=21)
input3.grid(row=2, column=1)

label2 = ttk.Label(window, text="Registration:")
label2.grid(row=3, column=0, pady=10, padx=12)
input2 = ttk.Combobox(window, values=pd.read_vehicles(), state="readonly", width=18)
input2.grid(row=3, column=1)


open_sheet = ttk.Button(window, text="Open Spreadsheet", command=pd.Open_existing_Sheet)
open_sheet.grid(row=0, column=0, pady=10, padx=10)

admin_gui = ttk.Button(window, text="Admin", command=Admin_open)
admin_gui.grid(row=0, column=1, pady=10, padx=10)

submit_button = ttk.Button(window, text="Submit", command=submit)
submit_button.grid(row=6, column=0, pady=5, padx=10)

add_Heading_button = ttk.Button(window, text="Add heading", command=add_Heading)
add_Heading_button.grid(row=6, column=1, pady=5, padx=10)

submit_all_button = ttk.Button(window, text="Do all existing vehicles", command=submit_all)
submit_all_button.grid(row=7, column=0, pady=5, padx=10)

stop_process_button = ttk.Button(window, text="Kill All", command=stop_process)
stop_process_button.grid(row=7, column=1, pady=5, padx=10)

window.mainloop()