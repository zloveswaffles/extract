import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import tabula
import pdfplumber
import camelot
import pandas as pd
import os
import logging
import sys
import threading
import re
from datetime import datetime
import time

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Verify Java/tabula setup
def verify_tabula_java():
    try:
        tabula.environment_info()
        logger.info("Java environment verified successfully.")
    except Exception as e:
        messagebox.showerror("Java Error", f"Ensure Java is installed correctly.\n{str(e)}")
        sys.exit()

# File selection
def select_pdf_file():
    file_path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        logger.info("No file selected. Exiting.")
        sys.exit()
    logger.info(f"Selected file: {file_path}")
    return file_path

# Extract tables with tabula
def extract_tables_tabula(pdf_path, pages):
    tables = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=True, lattice=True)
    dfs = [df.dropna(how='all').reset_index(drop=True) for df in tables if not df.empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# Extract text data with pdfplumber
def extract_tables_pdfplumber(pdf_path, pages):
    with pdfplumber.open(pdf_path) as pdf:
        selected_pages = pdf.pages if pages.lower() == 'all' else [pdf.pages[i] for i in parse_page_numbers(pages, len(pdf.pages))]
        all_text = []
        for page in selected_pages:
            text = page.extract_text()
            if text:
                all_text.extend(text.split('\n'))
    return pd.DataFrame({"Extracted Text": all_text})

# Extract tables with camelot
def extract_tables_camelot(pdf_path, pages):
    page_numbers = pages if pages.lower() == 'all' else ','.join(map(str, [p+1 for p in parse_page_numbers(pages, 1000)]))
    tables = camelot.read_pdf(pdf_path, pages=page_numbers, flavor='stream')
    dfs = [table.df for table in tables if not table.df.empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# Parse page numbers
def parse_page_numbers(pages, max_pages):
    page_nums = []
    for part in pages.split(','):
        if '-' in part:
            start, end = map(int, part.split('-'))
            page_nums.extend(range(start - 1, end))
        else:
            page_nums.append(int(part) - 1)
    return [p for p in page_nums if p < max_pages]

# Solve filename conflict
def ensure_unique_filename(filepath):
    base, ext = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(filepath):
        filepath = f"{base}_{counter}{ext}"
        counter += 1
    return filepath

# Clean extracted data
def clean_dataframe(df):
    def clean_cell(cell):
        if isinstance(cell, str):
            cell = cell.replace('$', '').strip()
            if re.match(r'^\(\s*[\d,.]+\s*\)$', cell):
                cell = '-' + cell.strip('() ').replace(',', '')
            return cell
        return cell

    return df.map(clean_cell)

# Run extraction processes and save to one Excel file with separate tabs
def run_extraction(pdf_path, pages, user_inputs):
    camelot_df = clean_dataframe(extract_tables_camelot(pdf_path, pages))
    plumber_df = clean_dataframe(extract_tables_pdfplumber(pdf_path, pages))
    tabula_df = clean_dataframe(extract_tables_tabula(pdf_path, pages))

    year, period, audit_status, client_name = user_inputs
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_dir = os.path.dirname(pdf_path)
    filename = f"{year}_{period}_{audit_status}_{client_name}_{base_name}_Combined_Extracted.xlsx"
    output_file = ensure_unique_filename(os.path.join(output_dir, filename))

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        if not camelot_df.empty:
            camelot_df.to_excel(writer, sheet_name='Camelot', index=False)
        if not plumber_df.empty:
            plumber_df.to_excel(writer, sheet_name='PDFPlumber', index=False)
        if not tabula_df.empty:
            tabula_df.to_excel(writer, sheet_name='Tabula', index=False)

    logger.info(f"Saved combined extracted data to {output_file}")

# Main UI function
def main():
    root = tk.Tk()
    root.title("ZABL")
    window_width, window_height = 500, 500
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = (screen_width/2) - (window_width/2)
    y_coordinate = (screen_height/2) - (window_height/2)
    root.geometry(f"{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}")

    inputs_frame = ttk.Frame(root)
    inputs_frame.pack(pady=10)

    current_year = datetime.now().year

    client_name_var = tk.StringVar()
    year_var = tk.StringVar(value=str(current_year))
    period_var = tk.StringVar()
    audit_status_var = tk.StringVar()
    pages_var = tk.StringVar()

    ttk.Label(inputs_frame, text="Client Name:").pack(pady=5)
    ttk.Entry(inputs_frame, textvariable=client_name_var).pack(pady=5)

    ttk.Label(inputs_frame, text="Year:").pack(pady=5)
    ttk.Combobox(inputs_frame, textvariable=year_var, values=[str(y) for y in range(2010, 2051)]).pack(pady=5)

    ttk.Label(inputs_frame, text="Period:").pack(pady=5)
    ttk.Combobox(inputs_frame, textvariable=period_var, values=["Quarterly", "Annual"]).pack(pady=5)

    ttk.Label(inputs_frame, text="Audit Status:").pack(pady=5)
    ttk.Combobox(inputs_frame, textvariable=audit_status_var, values=["Audited", "Company Prepared"]).pack(pady=5)

    ttk.Label(inputs_frame, text="Pages (e.g., 1,2,5-7 or 'all'):").pack(pady=5)
    ttk.Entry(inputs_frame, textvariable=pages_var).pack(pady=5)

    def start_extraction():
        pdf_path = select_pdf_file()
        user_inputs = (year_var.get(), period_var.get(), audit_status_var.get(), client_name_var.get().upper())
        verify_tabula_java()

        load_win = tk.Toplevel(root)
        load_win.title("Extracting...")
        img = ImageTk.PhotoImage(Image.open(r"C:\Users\zarik\Desktop\hacker.jpg"))
        label_img = ttk.Label(load_win, image=img)
        label_img.image = img
        label_img.pack()

        progress = ttk.Progressbar(load_win, mode='indeterminate')
        progress.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        progress.start(10)

        def extraction_thread():
            run_extraction(pdf_path, pages_var.get(), user_inputs)
            time.sleep(10)
            load_win.destroy()

        threading.Thread(target=extraction_thread).start()

    ttk.Button(root, text="Upload Financial Statement", command=start_extraction).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
