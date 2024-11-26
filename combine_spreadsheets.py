import glob
import pandas as pd

# Specify the folder path containing the Excel files
path = "/home/wsl_ubuntu/Capstone-RIT-Fall2024/Information/H1B 2019 - 2024 Filtered/"

# Get all Excel files in the specified path
file_list = glob.glob(path + "/*.xlsx")

# List to hold DataFrames
excl_list = []

# Load and append each file to the list
for file in file_list:
    try:
        print(f"Processing file: {file}")
        excl_list.append(pd.read_excel(file))
    except Exception as e:
        print(f"Error reading {file}: {e}")

# Combine all DataFrames in the list into a single DataFrame
if excl_list:
    excl_merged = pd.concat(excl_list, ignore_index=True)
    
    # Save the combined DataFrame to a new Excel file
    output_file = 'H1B-combined-2019-24_V2.xlsx'
    try:
        excl_merged.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Combined Excel file saved to {output_file}")
    except Exception as e:
        print(f"Error saving combined file: {e}")
else:
    print("No files were processed. Please check the folder path or file contents.")
