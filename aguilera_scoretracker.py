from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk

def validate_inputs():
    studentname = studentname_entry.get()
    score = score_entry.get()

    if not studentname or not score:
        messagebox.showerror("Error", "All fields are required.")
        return False
    
    try:
        score_value =int(score)
        if score_value < 0 or score_value > 100:
            messagebox.showerror("Score reached limit.")
            return False
    except ValueError:
        messagebox.showerror("Score must be a numeric value.")
        return False

    return True

def save_to_excel():
    if not validate_inputs():
        return
    
    studentname = studentname_entry.get()
    score = int(score_entry.get())
    grade = grade_remarks(score)

    try:
        wb = load_workbook("student_scores.xlsx")
        ws = wb["ScoreTracker"]
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "ScoreTracker"
        ws.append(["Student Name", "Score", "Grade"])

    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == studentname:
            ws.cell(row=row, column=2, value=score)
            ws.cell(row=row, column=3, value=grade)
            messagebox.showinfo("Success", f"Updated record for {studentname}.")
            wb.save("student_scores.xlsx")
            studentname_entry.delete(0, tk.END)
            score_entry.delete(0, tk.END)
            return

    ws.append([studentname, score, grade])
    wb.save("student_scores.xlsx")
    messagebox.showinfo("Succes","Data saved to Excel.")
    studentname_entry.delete(0, tk.END)
    score_entry.delete(0, tk.END)

def grade_remarks(score):
    if score <= 74:
        return "Failed"
    else:
        return "Passed"

def show_all_records():
    try:
        wb = load_workbook("student_scores.xlsx")
        ws = wb["ScoreTracker"]
    except FileNotFoundError:
        messagebox.showerror("Excel File not found")
        return

    records = []
    total_score = 0
    count = 0

    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        score = ws.cell(row=row, column=2).value
        grade = ws.cell(row=row, column=3).value
        records.append((name, score, grade))
        total_score += score
        count += 1

    average_score = total_score / count if count > 0 else 0

    record_window = tk.Toplevel()
    record_window.title("Score Tracker")
    record_window.geometry("400x300")

    layout = ttk.Treeview(record_window, columns=("Student Name", "Score", "Grade"), show='headings')
    layout.heading("Student Name", text="Student Name")
    layout.heading("Score", text="Score")
    layout.heading("Grade", text="Grade")

    layout.column("Student Name", width=150)
    layout.column("Score", width=100)
    layout.column("Grade", width=100)

    for record in records:
        layout.insert("", tk.END,values=record)

    scrollbar = ttk.Scrollbar(record_window, orient="vertical", command=layout.yview)
    layout.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    layout.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)
    average_label = tk.Label(record_window, text=f"Average Score: {average_score:.2f}", font=("Arial", 12))
    average_label.pack(pady=10)

window = tk.Tk()
window.title("Score Tracker")
height = 100
width = 250
x = (window.winfo_screenwidth()//2)-(width//2)
y = (window.winfo_screenheight()//2)-(height//2)
window.geometry('{}x{}+{}+{}'.format(width,height,x,y))

tk.Label(window, text="Student Name").grid(row=2, column=0, padx=20,pady=2,sticky="NSEW")
tk.Label(window, text="Score").grid(row=3, column=0,padx=20, pady=2,sticky="NSEW")

studentname_entry = tk.Entry(window)
score_entry = tk.Entry(window)

studentname_entry.grid(row=2, column=1,sticky="NSEW", pady=2)
score_entry.grid(row=3, column=1,sticky="NSEW", pady=2)

tk.Button(window, text="Submit", command=save_to_excel, bg="#1899D6", fg="#FFFFFF").grid(row=9, column=0,pady=6, ipadx=20,ipady=1, sticky="W", padx=10)
tk.Button(window, text="Show All Records", command=show_all_records, bg="#1899D6", fg="#FFFFFF").grid(row=9,column=1, pady=6, sticky="W",padx=10, ipady=2)

window.mainloop()