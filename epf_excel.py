import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# this version works perfectly!
# created by darwish 21/08/2024

def extract_ip(row):
    # first non-N/A IP address from the Asset list
    for ip in [row.get('New IP'), row.get('IP ADD. 1'), row.get('IP ADD. 2')]:
        if pd.notna(ip):
            return ip
    return None

def find_column(columns, possible_names):
    # find correct column name from the list of possible names
    for name in possible_names:
        if name in columns:
            return name
    return None

def process_files(nessus_report_path, asset_list_path, output_path):
    try:
        # load the excel files
        nessus_report = pd.read_excel(nessus_report_path)
        asset_workbook = pd.ExcelFile(asset_list_path)
        merged_data = []
        
        # possible column names for each needed collumns from the asset list
        serial_number_columns = ['SERIAL NUMBER / VMWARE UUID', 'SERIAL NUMBER']
        remark_columns = ['REMARK / STATUS', 'Remark', 'Remark (inventory number)']
        sys_admin_columns = ['SERVER/SYSTEM ADMINISTRATORS (DID)', 'SERVER/SYSTEM ADMINISTRATORS']
        app_admin_columns = ['APPLICATION ADMINISTRATORS']
        
        # iterate through each sheet in the Asset List
        for sheet_name in asset_workbook.sheet_names:
            if sheet_name == "Cover Page":  # Ignore the "Cover Page" sheet
                continue
            
            asset_list = pd.read_excel(asset_workbook, sheet_name=sheet_name)
            
            # find relevant columns
            serial_number_col = find_column(asset_list.columns, serial_number_columns)
            remark_col = find_column(asset_list.columns, remark_columns)
            sys_admin_col = find_column(asset_list.columns, sys_admin_columns)
            app_admin_col = find_column(asset_list.columns, app_admin_columns)
            
            # extract IP addresses from Asset list
            asset_list['Extracted IP'] = asset_list.apply(extract_ip, axis=1)
            
            # DataFrame with only the relevant columns, if they exist
            relevant_columns = ['Extracted IP', serial_number_col, remark_col, sys_admin_col, app_admin_col]
            relevant_columns = [col for col in relevant_columns if col is not None]  # Filter out None values
            
            if not relevant_columns:
                continue  # skip sheets with no relevant columns
            
            asset_list_relevant = asset_list[relevant_columns]
            
            # merge and keep all rows from Asset list, filling missing data with N/A
            merged_report = pd.merge(
                asset_list_relevant,
                nessus_report,
                left_on='Extracted IP', right_on='IP Address',
                how='left'
            )
            
            # column to indicate whether there was a match
            merged_report['Match'] = merged_report['IP Address'].notna()
            merged_report['Match'] = merged_report['Match'].replace({True: 'Match', False: 'No Match'})
            
            # sheet name to the merged report
            merged_report['Sheet Name'] = sheet_name
            
            # append the merged report to the list
            merged_data.append(merged_report)
        
        # combine all merged reports from different sheets
        final_report = pd.concat(merged_data, ignore_index=True)
        
        # write to separate sheets based on severity and separate "No Match" as different sheet
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for severity in ['Low', 'Medium', 'High', 'Critical']:
                severity_df = final_report[final_report['Severity'] == severity]
                if not severity_df.empty:
                    severity_df.to_excel(writer, sheet_name=severity, index=False)
            
            # include a sheet for unmatched rows
            no_match_df = final_report[final_report['Match'] == 'No Match']
            if not no_match_df.empty:
                no_match_df.to_excel(writer, sheet_name='No Match', index=False)
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def save_file(entry):
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# main application window
root = tk.Tk()
root.title("Nessus Report Filter")

# Uploads
tk.Label(root, text="Nessus Report:").grid(row=0, column=0, padx=10, pady=10)
nessus_entry = tk.Entry(root, width=50)
nessus_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(nessus_entry)).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Asset List:").grid(row=1, column=0, padx=10, pady=10)
asset_entry = tk.Entry(root, width=50)
asset_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(asset_entry)).grid(row=1, column=2, padx=10, pady=10)

# Output File Save Location
tk.Label(root, text="Save Filtered Report As:").grid(row=2, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="Save As", command=lambda: save_file(output_entry)).grid(row=2, column=2, padx=10, pady=10)

# Process Button
tk.Button(root, text="Process", command=lambda: process_files(nessus_entry.get(), asset_entry.get(), output_entry.get())).grid(row=3, column=1, pady=20)

root.mainloop()
# note: pyinstaller --onefile --windowed epf_excel.py