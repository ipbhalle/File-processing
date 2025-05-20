import pandas as pd
import os

def parse_csv_and_extract_protein_info(csv_file, excel_file):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(csv_file)
    
    # Extract the columns Protein_ID and Localizations
    protein_data = df[['Protein_ID', 'Localizations']]

    # Write the extracted data to an Excel file
    protein_data.to_excel(excel_file, index=False, engine="openpyxl")
    print(f"Excel file created: {excel_file}")

def process_directory():
    # Automatically process all .csv files in the current directory (PWD)
    current_directory = os.getcwd()
    for filename in os.listdir(current_directory):
        if filename.endswith(".csv"):
            csv_file = os.path.join(current_directory, filename)
            excel_file = os.path.join(current_directory, f"{os.path.splitext(filename)[0]}_output.xlsx")
            parse_csv_and_extract_protein_info(csv_file, excel_file)

if __name__ == "__main__":
    # Process CSV files in the current working directory
    process_directory()
