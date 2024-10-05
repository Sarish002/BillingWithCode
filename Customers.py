# Import Statement
import tkinter.messagebox
from ttkbootstrap import *
import openpyxl as opx
import sqlite3

# Window
root = Window(themename = "vapor")
root.geometry("800x500")
root.title("IM Customers")

# Activating the customers excel sheet
Custum = opx.load_workbook("Customers.xlsx")
Sheeter = Custum.active

# Connecting to a database
CustNum = sqlite3.connect("CustNum.db")
Cust = CustNum.cursor()

# Stirng Variable
Bile = StringVar(value="")

# Style
my = Style()
my.configure("light.TButton",
             font=("Trebuchet MS",
                   18,
                   "bold"))

# Entry
E1 = Entry(root,
           font=("Comic Sans MS",
                 15),
           width = 17,
           bootstyle = "light")
E1.place(relx = 0.2345,
         rely = 0.23,
         anchor = "center")
E1.insert(0,
          "Name or phone no.")

# The combobox
C1 = Combobox(root,
              bootstyle = "info",
              font=("Comic Sans MS",
                    15),
              width=16)
C1["values"] = ["Name",
                "Phone no.",
                "Amount",
                "Items Bought"] # Adding a datalist
C1.place(relx = 0.2345,
         rely = 0.5,
         anchor = "center")

# Label
L1 = Label(root, textvariable=Bile,
           font=("Comic Sans MS", 8),
           bootstyle="light",
           justify="center")
L1.place(relx = 0.745,
         rely = 0.5,
         anchor = "center")

# Getting Info
def GetInfo():
    Info = E1.get()
    Need = C1.get()
    try:
        Info = int(Info)
    except:
        ...
    Cust.execute("select * from Things_Bought")
    rec = Cust.fetchall()

    for i in rec:
        for j in i:
            if j == Info:
                if Need == "Name":
                    Bile.set(f"Name: {i[0]}")
                elif Need == "Phone no.":
                    Bile.set(f"Phone no.: {i[1]}")
                elif Need == "Items Bought":
                    Bile.set(f"{i[2]}")
                elif Need == "Amount":
                    Bile.set(f"Amount: {i[3]}")
                return
    tkinter.messagebox.showinfo("sqlite3.NotFoundError",
                                "We cannnot find the person you're"
                                                             " looking for.")

# Button for getting info
B1 = Button(root,
            text="Submit",
            width = 14,
            style="light.TButton",
            command=GetInfo)
B1.place(relx = 0.24,
         rely = 0.76,
         anchor = "center")

# Seperating with a rule
H1 = Separator(root,
               orient=VERTICAL,
               bootstyle="light")
H1.place(relx = 0.45,
         rely = 0.5,
         anchor = "center",
         relheight = 0.9)

# Mainloop
root.mainloop()