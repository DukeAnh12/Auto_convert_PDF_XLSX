import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
from datetime import datetime

# Get the current time
current_time = datetime.now()

# Format the current time as dd_mm_yy_HH_MM_ss
formatted_time = current_time.strftime("%d_%m_%y_%H_%M_%S")

current_file_path = os.path.abspath(__file__)
current_directory = os.path.dirname(current_file_path)
# "C:/Users/TSB/Desktop/Demo_Auto_Convert/Excel_output/output.xlsx"
def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_file_entry.delete(0, tk.END)
        pdf_file_entry.insert(0, file_path)

def process_pdf():
    pdf_file_path = pdf_file_entry.get()
    auto_convert_file_path = current_directory.replace("\\","/")+"/Auto_convert.py"
    output_file_path = current_directory.replace("\\","/").replace("/py","/Excel_output/")+"output_converted_"+ formatted_time +".xlsx"
    if pdf_file_path:
        subprocess.run(["py", auto_convert_file_path , pdf_file_path, output_file_path])

root = tk.Tk()
root.title("PDF Processing App")

pdf_file_label = tk.Label(root, text="Select PDF file:")
pdf_file_label.grid(row=0, column=0, padx=5, pady=5)

pdf_file_entry = tk.Entry(root, width=50)
pdf_file_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=select_pdf_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

process_button = tk.Button(root, text="Process PDF", command=process_pdf)
process_button.grid(row=2, column=1, padx=5, pady=5)

root.mainloop()
