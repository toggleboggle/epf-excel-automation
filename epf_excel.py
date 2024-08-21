import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def extract_ip(row):
    # first non-N/A IP address from the Asset list
    for ip in [row['New IP'], row['IP ADD. 1'], row['IP ADD. 2']]:
        if pd.notna(ip):
            return ip
    return None

def process_files(nessus_report_path, asset_list_path, output_path):
    try:
        # load the Nessus report
        nessus_report = pd.read_excel(nessus_report_path)
        
        # load the asset list
        asset_workbook = pd.ExcelFile(asset_list_path)
        
        # create list to store the merged data
        merged_data = []
        
        # iterate through each sheet in the Asset List
        for sheet_name in asset_workbook.sheet_names:
            asset_list = pd.read_excel(asset_workbook, sheet_name=sheet_name)
            
            # Keep only the first 11 columns from the asset list
            asset_list = asset_list.iloc[:, :11]
            
            # Ensure correct columns are used
            asset_list.columns = [
                'No', 'VM Folder', 'SYSTEM NAME', 'HOSTNAME', 'New IP',
                'IP ADD. 1', 'IP ADD. 2', 'SERIAL NUMBER / VMWARE UUID',
                'REMARK / STATUS', 'SERVER/SYSTEM ADMINISTRATORS (DID)', 
                'APPLICATION ADMINISTRATORS'
            ]
            
            # Extract and clean IP addresses from Asset list
            asset_list['Extracted IP'] = asset_list.apply(extract_ip, axis=1)
            
            # Merge and keep all rows from Asset list, filling missing data with NaN
            merged_report = pd.merge(
                asset_list,
                nessus_report,
                left_on='Extracted IP', right_on='IP Address',
                how='left'
            )
            
            # Add a column to indicate whether there was a match
            merged_report['Match'] = merged_report['IP Address'].notna()
            merged_report['Match'] = merged_report['Match'].replace({True: 'Match', False: 'No Match'})
            
            # Add the sheet name to the merged report
            merged_report['Sheet Name'] = sheet_name
            
            # Append the merged report to the list
            merged_data.append(merged_report)
        
        # Combine all merged reports from different sheets
        final_report = pd.concat(merged_data, ignore_index=True)
        
        # Write to separate sheets based on severity and include "No Match" as a separate sheet
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for severity in ['Low', 'Medium', 'High', 'Critical']:
                severity_df = final_report[final_report['Severity'] == severity]
                if not severity_df.empty:
                    severity_df.to_excel(writer, sheet_name=severity, index=False)
            
            # Include a sheet for unmatched rows
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


# Create the main application window
root = tk.Tk()
root.title("Nessus Report Filter")

# Nessus Report Upload
tk.Label(root, text="Nessus Report:").grid(row=0, column=0, padx=10, pady=10)
nessus_entry = tk.Entry(root, width=50)
nessus_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(nessus_entry)).grid(row=0, column=2, padx=10, pady=10)

# Asset List Upload
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

# Run the application
root.mainloop()
