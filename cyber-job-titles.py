import os
import glob
import pandas as pd
from tqdm import tqdm

def filter_cybersecurity_jobs_in_directory(input_dir, output_dir):
    """
    Processes all .xlsx files in a directory to filter cybersecurity-related jobs
    and save selected columns into new files.

    Parameters:
    - input_dir: Directory containing the .xlsx files.
    - output_dir: Directory where the output files will be saved.

    Each output file will be saved with the same base name as the input file in the specified output directory.
    """
    
    # Define keywords for filtering cybersecurity job titles
    cybersecurity_keywords = [
        "cybersecurity", "security", "information security", "infosec", "cyber",
        "network security", "vulnerability", "penetration", "red team", "blue team",
        "purple team", "threat", "incident", "firewall", "IDS", "IPS", "appsec",
        "devsecops", "DFIR", "malware", "reverse engineer", "SOC", "GRC",
        "identity and access management", "IAM", "access control", "zero trust",
        "cryptography", "blockchain security", "PKI", "crypto", "encryption",
        "iot security", "SCADA", "ICS", "ethical hacking", "SIEM", "splunk", "QRadar"
    ]

    # Get list of all .xlsx files in the input directory
    file_paths = glob.glob(os.path.join(input_dir, "*.xlsx"))

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Process each file in the directory
    for file_path in file_paths:
        try:
            print(f"Processing file: {file_path}")
            data = pd.read_excel(file_path, sheet_name=0)

            # Filter rows with cybersecurity-related keywords in job titles
            tqdm.pandas(desc="Filtering Job Titles")
            cybersecurity_jobs = data[data['JOB_TITLE'].progress_apply(
                lambda x: any(keyword in str(x).lower() for keyword in cybersecurity_keywords))]

            # Select relevant columns
            selected_columns = cybersecurity_jobs[
                [
                    'EMPLOYER_NAME', 'JOB_TITLE', 'SOC_TITLE', 'WAGE_RATE_OF_PAY_FROM',
                    'WAGE_RATE_OF_PAY_TO', 'WAGE_UNIT_OF_PAY', 'PREVAILING_WAGE',
                    'PW_WAGE_LEVEL', 'EMPLOYER_ADDRESS1', 'EMPLOYER_CITY',
                    'EMPLOYER_STATE', 'EMPLOYER_POSTAL_CODE', 'WORKSITE_CITY',
                    'WORKSITE_COUNTY', 'WORKSITE_STATE', 'WORKSITE_POSTAL_CODE',
                    'RECEIVED_DATE', 'DECISION_DATE', 'VISA_CLASS',
                    'BEGIN_DATE', 'END_DATE', 'CASE_STATUS'
                ]
            ]

            # Rename columns
            selected_columns.rename(columns={
                'EMPLOYER_NAME': 'Employer Name',
                'JOB_TITLE': 'Job Title',
                'SOC_TITLE': 'SOC Title',
                'WAGE_RATE_OF_PAY_FROM': 'Wage Rate of Pay From',
                'WAGE_RATE_OF_PAY_TO': 'Wage Rate of Pay To',
                'WAGE_UNIT_OF_PAY': 'Wage Unit of Pay',
                'PREVAILING_WAGE': 'Prevailing Wage',
                'PW_WAGE_LEVEL': 'PW Wage Level',
                'EMPLOYER_ADDRESS1': 'Employer Address 1',
                'EMPLOYER_CITY': 'Employer City',
                'EMPLOYER_STATE': 'Employer State',
                'EMPLOYER_POSTAL_CODE': 'Employer Postal Code',
                'WORKSITE_CITY': 'Worksite City',
                'WORKSITE_COUNTY': 'Worksite County',
                'WORKSITE_STATE': 'Worksite State',
                'WORKSITE_POSTAL_CODE': 'Worksite Postal Code',
                'RECEIVED_DATE': 'Received Date',
                'DECISION_DATE': 'Decision Date',
                'VISA_CLASS': 'Visa Class',
                'BEGIN_DATE': 'Employment Start Date',
                'END_DATE': 'Employment End Date',
                'CASE_STATUS': 'Case Status'
            }, inplace=True)

            # Define output path
            output_file_name = os.path.join(output_dir, os.path.basename(file_path).replace('.xlsx', '_filtered.xlsx'))

            # Save to Excel
            print(f"Saving to {output_file_name}")
            with tqdm(total=1, desc="Saving") as pbar:
                selected_columns.to_excel(output_file_name, index=False, engine='openpyxl')
                pbar.update(1)

            print(f"File saved: {output_file_name}")

        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

# Example usage:
input_dir = '/home/wsl_ubuntu/Capstone-RIT-Fall2024/LCA Disclosure Data 2008 - 2024/LCA-FY-2024/'
output_dir = '/home/wsl_ubuntu/Capstone-RIT-Fall2024/Information/'
filter_cybersecurity_jobs_in_directory(input_dir, output_dir)
