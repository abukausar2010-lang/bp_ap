
import json
import tkinter as tk
from tkinter import messagebox
import datetime
import openpyxl

# ЖСН → Аты-жөні дерекқоры
with open("people_db.json", "r", encoding="utf-8") as f:
    PEOPLE_DB = json.load(f)

def save_to_excel(jsn, name, systolic, diastolic):
    filename = "bp_records.xlsx"
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Күні", "ЖСН", "Аты-жөні", "Қан қысымы (Жоғ)", "Қан қысымы (Төм)"])

    ws.append([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), jsn, name, systolic, diastolic])
    wb.save(filename)

def find_name_by_jsn(jsn):
    return PEOPLE_DB.get(jsn, "Белгісіз ЖСН")

def submit():
    jsn = entry_jsn.get().strip()
    systolic = entry_sys.get().strip()
    diastolic = entry_dia.get().strip()

    if not jsn or not systolic or not diastolic:
        messagebox.showerror("Қате", "Барлық өрісті толтырыңыз!")
        return

    name = find_name_by_jsn(jsn)
    save_to_excel(jsn, name, systolic, diastolic)
    messagebox.showinfo("Сақталды", f"{name} деректері Excel-ге сақталды.")
    entry_jsn.delete(0, tk.END)
    entry_sys.delete(0, tk.END)
    entry_dia.delete(0, tk.END)

# GUI
root = tk.Tk()
root.title("Қан қысымын тіркеу")
root.geometry("400x250")

tk.Label(root, text="ЖСН:").pack(pady=5)
entry_jsn = tk.Entry(root, width=30)
entry_jsn.pack()

tk.Label(root, text="Қан қысымы (Жоғарғы):").pack(pady=5)
entry_sys = tk.Entry(root, width=30)
entry_sys.pack()

tk.Label(root, text="Қан қысымы (Төменгі):").pack(pady=5)
entry_dia = tk.Entry(root, width=30)
entry_dia.pack()

tk.Button(root, text="Сақтау", command=submit).pack(pady=15)

root.mainloop()
