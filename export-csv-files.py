# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 14:16:41 2024

@author: Olivia Ashmoore
"""

import os
import pandas as pd

# Function to export tabs from an Excel file to CSV
def export_tabs_to_csv(excel_file_path, tabs_to_export, output_dir):
    try:
        # Open the Excel file
        xls = pd.ExcelFile(excel_file_path)
        
        for tab_name in tabs_to_export:
            try:
                # Read the specific tab (worksheet) into a DataFrame
                df = pd.read_excel(excel_file_path, sheet_name=tab_name)
                
                # Round all values in the DataFrame to the 10th significant figure
                df = df.round(6)
                
                # Define the CSV file path where you want to save the data
                csv_file_path = os.path.join(output_dir, f'{tab_name}.csv')
                
                # Export the DataFrame to a CSV file
                df.to_csv(csv_file_path, index=False)
                
                print(f'{tab_name} from {excel_file_path} exported to {csv_file_path}')
            except Exception as e:
                print(f'Error exporting {tab_name} from {excel_file_path} to CSV: {str(e)}')
    except Exception as e:
        print(f'Error opening Excel file {excel_file_path}: {str(e)}')

# Directory where you have your subfolders with Excel files
root_directory = r'C:\Users\Olivia Ashmoore\Documents\EPS_Models by Region\RMI\RMI-coding\state-elec-downscaling-oliv\Outputs'

# Recursively search for Excel files in subdirectories
for root, dirs, files in os.walk(root_directory):
    for file in files:
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_file_path = os.path.join(root, file)
            
            # Find existing CSV files in the same directory
            csv_files = [f for f in os.listdir(root) if f.endswith('.csv')]
            
            # Extract tab names from existing CSV file names
            tabs_to_export = [os.path.splitext(csv_file)[0] for csv_file in csv_files]
            
            # Call the export function, passing the output directory as the current directory
            export_tabs_to_csv(excel_file_path, tabs_to_export, root)


