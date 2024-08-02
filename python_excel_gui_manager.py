import tkinter as tk
from tkinter import ttk
from ctypes import windll
import openpyxl

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use('forest-dark')

def load_data():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    # print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name,text=col_name)
    for value_tuple in list_values[1:]:
        treeview.insert('',tk.END, values=value_tuple)
    
def insert_row():
    name = name_entery.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"
    

    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)
    treeview.insert('',tk.END, values=row_values)

    name_entery.delete(0,"end")
    name_entery.insert(0,"Name")
    age_spinbox.delete(0,"end")
    age_spinbox.insert(0,"Age")
    status_combobox.set(combo_list[0])
    checkbtn.state(["!selected"])

window = tk.Tk()

style =ttk.Style(window)
window.tk.call("source","forest-dark.tcl")
window.tk.call("source","forest-light.tcl")
style.theme_use("forest-dark")

combo_list = ["Subscribed","Not Subscribed"]

frame = ttk.Frame(window)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0,padx=10,pady=10)

name_entery = ttk.Entry(widgets_frame)
name_entery.insert(0,"Name")
name_entery.bind("<FocusIn>", lambda e: name_entery.delete('0','end'))
name_entery.grid(row=0,column=0,padx=5, pady=(0,5),sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0,'Age')
age_spinbox.grid(row=1,column=0,padx=5,pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame,values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2,column=0,padx=5,pady=5,sticky="ew")

a = tk.BooleanVar()
checkbtn = ttk.Checkbutton(widgets_frame, text="Employed",variable=a)
checkbtn.grid(row=3,column=0,padx=5,pady=5,sticky="nsew")

btn = ttk.Button(widgets_frame, text="Insert",command=insert_row)
btn.grid(row=4, column=0,padx=5,pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(15,15),pady=10,sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame,text="Mode",style="Switch",command=toggle_mode)
mode_switch.grid(row=6, column=0,padx=5,pady=5, sticky="ew")


treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0,column=1,pady=10)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right",fill="y")

cols = ("Name","Age","Subscription","Employment")

treeview = ttk.Treeview(treeFrame, show="headings",yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column('Name',width=100)
treeview.column('Age',width=50)
treeview.column('Subscription',width=100)
treeview.column('Employment',width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()



windll.shcore.SetProcessDpiAwareness(2)
window.mainloop()