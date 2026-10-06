"""
==========================================================================================
SCRIPT: Historical Trade Data Analyzer & Price Normalizer (GUI Application)

DESCRIPTION:
This Python desktop application (built with Tkinter and Pandas) is designed to process and 
analyze multi-year historical ERP/SAP trade data. It evaluates thousands of contracts, 
cross-references historical execution dates with daily FX/Pricing calculation files, 
and normalizes all freight and commodity prices into standard EUR/USD metrics.

PROCESS FLOW:
  1. GUI Initialization: Opens a user-friendly dialog to select the raw ERP Excel dump.
  2. Rate Aggregation: Scans the network drive for historical daily rate files matching 
     the contract's 'Doc-date' and extracts specific historical FX and pricing data.
  3. Temporal Matching: Matches the contract execution ('Ship start' + 1 month) with the 
     corresponding future forward rates dynamically.
  4. Financial Normalization: Converts multi-currency flat prices, truck freight, and 
     loading costs into normalized EUR/USD values.
  5. Validation Engine: Filters out invalid internal transfers, missing freights (CPT/DDP), 
     and unapproved users.
  6. Aggregation & Export: Calculates weighted average prices per commodity and exports 
     both a detailed line-by-line dataset and a grouped summary to a final Excel report.

TECHNICAL HIGHLIGHTS:
 - Utilizes 'pandas' for high-performance vectorized data manipulation.
 - Implements asynchronous threading ('threading') to keep the UI responsive during heavy I/O.
 - Custom dynamic regex file-matching for localized date formats in filenames.
==========================================================================================
"""

import os
import time
import glob
import shutil
import threading
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import warnings
import re

# Suppress openpyxl/pandas warnings for cleaner console output
warnings.simplefilter(action='ignore', category=UserWarning)

# ==============================================================================
#  USER CONFIGURATION & PATHS
# ==============================================================================

# Output folder for processed reports
PATH_SAVE_FOLDER = os.path.expanduser(r"~\Documents\Export_Reports")

# Root directory containing historical daily calculation files
PATH_DAILY_RATES_ROOT = r"C:\CompanyShares\TradeDept\Pricing\Daily_Calculations"

# Authorized user IDs for validation
LIST_VALID_USERS = [
    "USER_A", "USER_B", "USER_C", "USER_D", "USER_E", "USER_F"
]

# Internal entities to be excluded from average market calculations
LIST_IGNORE_PARTNERS = [
    "INTERNAL_COMPANY_SA",
    "INTERNAL_SUBSIDIARY_LTD"
]

# ==============================================================================
#  GUI & LOGGER CLASS
# ==============================================================================

class ProgressApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Trade Price Calculator - Local Processor")
        self.root.geometry("600x450")
        self.root.configure(bg="#f0f0f0")
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("green.Horizontal.TProgressbar", foreground='green', background='#4CAF50')
        
        self.lbl_title = tk.Label(root, text="Local Excel Data Processor", font=("Segoe UI", 16, "bold"), bg="#f0f0f0", fg="#333")
        self.lbl_title.pack(pady=(15, 5))
        
        self.lbl_status = tk.Label(root, text="Starting processing...", font=("Segoe UI", 10), bg="#f0f0f0", fg="#666")
        self.lbl_status.pack(pady=5)

        self.progress = ttk.Progressbar(root, style="green.Horizontal.TProgressbar", orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=10)

        self.log_text = tk.Text(root, height=15, width=70, font=("Consolas", 9), state='disabled', bg="white", relief="flat")
        self.log_text.pack(pady=10, padx=10)

        self.btn_close = tk.Button(root, text="Close", command=root.destroy, state='disabled', bg="#d9534f", fg="white", font=("Segoe UI", 9, "bold"))
        self.btn_close.pack(pady=5)

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        print(full_msg) 
        self.root.after(0, self._update_log, full_msg)

    def _update_log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.lbl_status.config(text=msg.split("] ")[1])

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def finish(self):
        self.root.after(0, lambda: self.btn_close.config(state='normal', bg="#4CAF50", text="DONE - Close"))
        self.root.after(0, lambda: messagebox.showinfo("Success", "Processing completed successfully! Check the log for details."))

gui_app = None

# ==============================================================================
#  PANDAS PROCESSING ENGINE
# ==============================================================================

def load_historical_rates(doc_dates):
    gui_app.log("Searching for historical daily rates...")
    rate_db = {}
    unique_dates = pd.to_datetime(doc_dates, errors='coerce').dropna().unique()
    
    for d in unique_dates:
        d_py = pd.to_datetime(d)
        year_str = str(d_py.year) 
        cache_key = d_py.strftime("%Y-%m-%d")
        
        # Possible date formats in filenames
        pat1 = d_py.strftime("%Y %m %d")                   
        pat2 = f"{d_py.year} {d_py.month:02d} {d_py.day}"  
        pat3 = f"{d_py.year} {d_py.month} {d_py.day}"      
        
        found_files = []
        search_pattern = os.path.join(PATH_DAILY_RATES_ROOT, f"*{year_str}*", "DAILY_RATES-*.xls*")
        
        for f in glob.glob(search_pattern):
            fname = os.path.basename(f)
            if pat1 in fname or re.search(rf"{pat2}(?!\d)", fname) or re.search(rf"{pat3}(?!\d)", fname):
                found_files.append(f)
            
        found_files = list(set(found_files))
        
        if not found_files:
            continue
        
        # Get the most recently modified file if multiple match
        file_path = max(found_files, key=os.path.getctime)
        temp_file_name = f"temp_read_{int(time.time())}_{os.path.basename(file_path)}"
        temp_file_path = os.path.join(PATH_SAVE_FOLDER, temp_file_name)
        
        try:
            # Create a temporary copy to avoid locking issues with shared network drives
            shutil.copy2(file_path, temp_file_path)
            
            df_temp = pd.read_excel(
                temp_file_path, 
                sheet_name="FCA_INDICATION", 
                usecols="K:P", 
                header=None, 
                skiprows=1, 
                nrows=18
            )
            df_temp.columns = ["Date", "Col1", "Col2", "USD", "Col4", "EUR"]
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce')
            
            # Filter out invalid dates (e.g., 1970 epoch errors from blank cells)
            df_temp = df_temp[df_temp["Date"] >= pd.Timestamp("2000-01-01")]
            
            df_temp = df_temp.dropna(subset=["Date"]).sort_values("Date")
            df_temp.attrs['source_file'] = os.path.basename(file_path)
            
            rate_db[cache_key] = df_temp
            
        except Exception as e:
            gui_app.log(f"Error reading file: {os.path.basename(file_path)} ({e})")
        finally:
            if os.path.exists(temp_file_path):
                try: os.remove(temp_file_path)
                except: pass
            
    return rate_db

def get_rate(row, rate_db, currency):
    try:
        doc_date = row['Doc-date']
        ship_start = row['Ship start']
        if pd.isnull(doc_date) or pd.isnull(ship_start): return 0.0
        
        key = doc_date.strftime("%Y-%m-%d")
        if key not in rate_db: return 0.0
        
        df_rates = rate_db[key]
        
        # Target matching period is Shipment Start + 1 month
        target_date = ship_start + pd.DateOffset(months=1)
        target_period = pd.Period(target_date, freq='M')
        
        periods = df_rates["Date"].dt.to_period('M')
        matches = df_rates[periods >= target_period]
        
        if matches.empty: return 0.0
        return float(matches.iloc[0][currency])
    except: return 0.0

def process_with_pandas(raw_excel_path, final_save_path):
    gui_app.log("Loading dataset...")
    gui_app.update_progress(20)
    
    try: 
        df = pd.read_excel(raw_excel_path)
    except Exception as e: 
        gui_app.log(f"ERROR reading the specified Excel file: {e}")
        return

    df.columns = df.columns.str.strip()
    
    for col in ['Doc-date', 'Ship start']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('.', '-', regex=False)
            df[col] = pd.to_datetime(df[col], errors='coerce')

    if 'Doc-date' in df.columns:
        rate_db = load_historical_rates(df['Doc-date'].unique())
    else:
        gui_app.log("ERROR: The selected file is missing the 'Doc-date' column!")
        return

    gui_app.update_progress(50)
    
    # =========================================================================
    # DEBUG BLOCK: VALIDATE FIRST 5 ROWS
    # =========================================================================
    gui_app.log("\n--- DEBUG: CHECKING FIRST 5 ROWS ---")
    try:
        for idx, row in df.head(5).iterrows():
            doc_date = row.get('Doc-date')
            ship_start = row.get('Ship start')
            flat_price = row.get('Flat Price', 0)
            flat_cur = str(row.get('Flat Price CUR', '')).strip().upper()
            
            gui_app.log(f"\n[ERP Row {idx+2}] Contract: {flat_price} {flat_cur} | Doc: {doc_date.strftime('%Y-%m-%d') if pd.notnull(doc_date) else 'N/A'}, Ship: {ship_start.strftime('%Y-%m-%d') if pd.notnull(ship_start) else 'N/A'}")
            
            if pd.isnull(doc_date) or pd.isnull(ship_start):
                gui_app.log("  -> Missing date, rate cannot be calculated.")
                continue
                
            key = doc_date.strftime("%Y-%m-%d")
            if key not in rate_db:
                gui_app.log(f"  -> Historical rate file not loaded for date: {key}!")
                continue
                
            df_rates = rate_db[key]
            
            target_date = ship_start + pd.DateOffset(months=1)
            target_period = pd.Period(target_date, freq='M')
            
            source_file = df_rates.attrs.get('source_file', 'Unknown file')
            gui_app.log(f"  -> Source file: {source_file}")
            gui_app.log(f"  -> Target lookup month: {target_period.strftime('%Y-%m')} (Ship start + 1 month)")
            
            periods = df_rates["Date"].dt.to_period('M')
            matches = df_rates[periods >= target_period]
            
            if matches.empty:
                gui_app.log("  -> No matching (or subsequent) month found in rate file column 'K'!")
            else:
                best_match = matches.iloc[0]
                excel_row = best_match.name + 3
                gui_app.log(f"  -> MATCH: Excel row {excel_row}")
                gui_app.log(f"  -> Date (Col K): {best_match['Date'].strftime('%Y-%m-%d')}")
                gui_app.log(f"  -> USD (Col N): {best_match['USD']}")
                gui_app.log(f"  -> EUR (Col P): {best_match['EUR']}")
                
    except Exception as e:
        gui_app.log(f"Debug runtime error: {e}")
    gui_app.log("--- DEBUG END ---\n")
    # =========================================================================

    if rate_db:
        df['EUR_Rate'] = df.apply(lambda row: get_rate(row, rate_db, 'EUR'), axis=1)
        df['USD_Rate'] = df.apply(lambda row: get_rate(row, rate_db, 'USD'), axis=1)
    else:
        df['EUR_Rate'] = 0.0; df['USD_Rate'] = 0.0

    df['EUR/USD Cross Rate'] = np.where(df['USD_Rate'] > 0, df['EUR_Rate'] / df['USD_Rate'], 0.0)

    numeric_cols = ['Flat Price', 'R-Truck freight', 'R-RW Freight', 'R-Loading (R/W)', 'Ctr/nom. Qty']
    for c in numeric_cols:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        else: df[c] = 0.0

    def calc_eur(row, val_col, cur_col):
        val = row.get(val_col, 0)
        cur = str(row.get(cur_col, "")).strip().upper()
        eur = row['EUR_Rate']; usd = row['USD_Rate']
        if val == 0: return 0.0
        if cur == "EUR": return val
        if cur == "HUF" and eur > 0: return val / eur
        if cur == "USD" and eur > 0 and usd > 0: return (val * usd) / eur
        return 0.0

    def calc_usd(row, val_col, cur_col):
        val = row.get(val_col, 0)
        cur = str(row.get(cur_col, "")).strip().upper()
        eur = row['EUR_Rate']; usd = row['USD_Rate']
        if val == 0: return 0.0
        if cur == "USD": return val
        if cur == "HUF" and usd > 0: return val / usd
        if cur == "EUR" and eur > 0 and usd > 0: return (val * eur) / usd 
        return 0.0

    df['Flat EUR'] = df.apply(lambda r: calc_eur(r, 'Flat Price', 'Flat Price CUR'), axis=1)
    df['Flat USD'] = df.apply(lambda r: calc_usd(r, 'Flat Price', 'Flat Price CUR'), axis=1)
    df['freight EUR'] = df.apply(lambda r: calc_eur(r, 'R-Truck freight', 'R-Truck freight CUR'), axis=1)
    df['RW Freight EUR'] = df.apply(lambda r: calc_eur(r, 'R-RW Freight', 'R-RW Freight CUR'), axis=1)
    df['Loading EUR'] = df.apply(lambda r: calc_eur(r, 'R-Loading (R/W)', 'R-Loading (R/W) CUR'), axis=1)
    
    # Calculate Net Flat Price
    df['Net Flat EUR'] = df['Flat EUR'] - df['freight EUR'] - df['RW Freight EUR'] - df['Loading EUR']

    gui_app.log("Generating validation status...")
    gui_app.update_progress(65)
    
    def check_status(row):
        reasons = []
        if str(row.get('Person responsible', '')).strip().upper() not in LIST_VALID_USERS: pass
        if str(row.get('Flat Price CUR', '')).strip().upper() == "HUF" and 0 < row.get('Flat Price', 0) < 50000: reasons.append("LOW PRICE")
        
        p = str(row.get('Partner name', '')).strip().upper()
        for b in LIST_IGNORE_PARTNERS:
            if b in p: reasons.append("INTERNAL PARTNER"); break
            
        inco = str(row.get('Inco1', '')).strip().upper()
        if inco in ["DDP", "CPT"] and row.get('R-Truck freight', 0) == 0 and row.get('R-RW Freight', 0) == 0: reasons.append("MISSING FREIGHT")
        if str(row.get('Type', '')).strip().upper() == "P": reasons.append("TYPE P")
        
        return "SKIPPED: " + ", ".join(reasons) if reasons else "OK"

    status_results = df.apply(check_status, axis=1)
    if "Status" in df.columns: df["Status"] = status_results
    else: df.insert(0, "Status", status_results)

    df.rename(columns={'EUR_Rate': 'EUR/LOCAL_Rate', 'USD_Rate': 'USD/LOCAL_Rate'}, inplace=True)

    gui_app.update_progress(80)
    # Filter only valid rows for the summary
    df_clean = df[(df['Status'] == "OK") & (df['Person responsible'].astype(str).str.strip().str.upper().isin(LIST_VALID_USERS))].copy()
    
    def normalize_comm(row):
        raw = str(row.get('Commodity desc', '')).strip()
        plant = str(row.get('Base loc name', '')).strip()
        if "ANY Rapeseed meal feed" in raw:
            return "HU Rape seed meal feed" if "Plant_A" in plant else "AT Rape seed meal feed"
        elif "ANY Sunflower seed meal" in raw:
            return "HU Sunflower seed meal feed" if "Plant_A" in plant else "AT Sunflower seed meal feed, dehulled"
        return raw

    df_clean['Final_Comm'] = df_clean.apply(normalize_comm, axis=1)
    df_clean['Total_Weighted_Val'] = df_clean['Ctr/nom. Qty'] * df_clean['Net Flat EUR']
    
    # Generate Summary Pivot
    summary = df_clean.groupby('Final_Comm').agg(
        Total_Qty=('Ctr/nom. Qty', 'sum'),
        Total_Val_Sum=('Total_Weighted_Val', 'sum')
    ).reset_index()
    
    summary['Weighted Avg Price (EUR)'] = summary['Total_Val_Sum'] / summary['Total_Qty']
    summary = summary.drop(columns=['Total_Val_Sum']).rename(columns={'Final_Comm': 'Commodity', 'Total_Qty': 'Total Quantity'})

    gui_app.log("Saving Excel report...")
    gui_app.update_progress(90)
    try:
        with pd.ExcelWriter(final_save_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Details", index=False)
            summary.to_excel(writer, sheet_name="Summary", index=False)
        gui_app.log(f"DONE: {os.path.basename(final_save_path)}")
        gui_app.update_progress(100)
        gui_app.finish()
    except Exception as e:
        gui_app.log(f"Save error: {e}")

# ==============================================================================
#  BACKGROUND WORKER THREAD
# ==============================================================================

def worker_thread(setup_data):
    input_file = setup_data["input_file"]
    
    gui_app.log(f"Selected file: {os.path.basename(input_file)}")
    
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_filename = f"{base_name}_Processed.xlsx"
    final_save_path = os.path.join(PATH_SAVE_FOLDER, output_filename)
    
    process_with_pandas(input_file, final_save_path)

# ==============================================================================
#  APPLICATION ENTRY POINT
# ==============================================================================

def get_setup_dialog():
    dialog = tk.Tk()
    dialog.title("Select File")
    
    w, h = 450, 150
    sw, sh = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
    dialog.geometry(f'{w}x{h}+{int(sw/2-w/2)}+{int(sh/2-h/2)}')
    
    file_var = tk.StringVar()
    
    tk.Label(dialog, text="Please select the downloaded ERP Excel file:", font=("Segoe UI", 10)).pack(pady=(20, 5))
    
    file_frame = tk.Frame(dialog)
    file_frame.pack(fill='x', padx=20)
    tk.Entry(file_frame, textvariable=file_var, state='readonly').pack(side='left', fill='x', expand=True, padx=(0, 5))
    
    def browse_file():
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if f: file_var.set(os.path.normpath(f))
            
    tk.Button(file_frame, text="Browse...", command=browse_file).pack(side='right')
    
    res = {}
    def ok():
        selected_file = file_var.get().strip()
        
        if not selected_file or not os.path.exists(selected_file):
            messagebox.showerror("Error", "Please select a valid Excel file!")
            return
            
        res["input_file"] = selected_file
        dialog.destroy()
            
    tk.Button(dialog, text="Start Processing", command=ok, width=20, bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold")).pack(pady=20)
    dialog.mainloop()
    return res

if __name__ == "__main__":
    setup_data = get_setup_dialog()
    if setup_data and "input_file" in setup_data:
        root = tk.Tk()
        gui_app = ProgressApp(root)
        
        # Run processing in a separate thread to prevent GUI freezing
        t = threading.Thread(target=worker_thread, args=(setup_data,))
        t.start()
        
        root.mainloop()
