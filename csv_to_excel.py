import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def browse_csv():
    csv_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV files", "*.csv")])
    if csv_path:
        entry_csv.delete(0, tk.END) 
        entry_csv.insert(0, csv_path) 

def browse_excel():
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        entry_excel.delete(0, tk.END) 
        entry_excel.insert(0, excel_path) 

def convert_csv_to_excel():
    csv_file = entry_csv.get()  
    excel_file = entry_excel.get()  
    
    if not csv_file or not excel_file:
        messagebox.showerror("Error", "Please select both CSV and Excel file locations.")
        return
    
    try:
        df = pd.read_csv(csv_file)
        
        df.to_excel(excel_file, index=False)
        
        messagebox.showinfo("Success", "CSV file has been successfully converted to Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

root = tk.Tk()
root.title("Prem's CSV to Excel Converter")

tk.Label(root, text="CSV File:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_csv = tk.Entry(root, width=40)
entry_csv.grid(row=0, column=1, padx=10, pady=10)
btn_browse_csv = tk.Button(root, text="Browse", command=browse_csv)
btn_browse_csv.grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save as Excel File:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_excel = tk.Entry(root, width=40)
entry_excel.grid(row=1, column=1, padx=10, pady=10)
btn_browse_excel = tk.Button(root, text="Browse", command=browse_excel)
btn_browse_excel.grid(row=1, column=2, padx=10, pady=10)

btn_convert = tk.Button(root, text="Convert", command=convert_csv_to_excel)
btn_convert.grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
