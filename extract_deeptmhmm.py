import pandas as pd
import re
import os

def parse_gff3_and_extract_gene_info(gff3_file, excel_file):
    # Initialize a list to store the gene data
    gene_data = []
    gene_id = None
    predicted_tmrs = None

    with open(gff3_file, 'r') as f:
        for line in f:
            line = line.strip()
            
            # Skip comment lines that contain gene info
            if line.startswith("#"):
                # Check for gene info lines (Length and Predicted TMRs)
                if "Length" in line:
                    gene_id = re.match(r"^([^\t]+)", line).group(1)  # Extract gene ID
                elif "Number of predicted TMRs" in line:
                    predicted_tmrs = int(re.search(r"Number of predicted TMRs: (\d+)", line).group(1))
                    # Once both gene_id and predicted_tmrs are found, store them
                    if gene_id is not None and predicted_tmrs is not None:
                        gene_data.append({
                            "Gene ID": gene_id,
                            "Predicted TMRs": predicted_tmrs
                        })
                    # Reset for next gene
                    gene_id = None
                    predicted_tmrs = None

    # Create a DataFrame from the gene data
    df = pd.DataFrame(gene_data)

    # Write to Excel
    df.to_excel(excel_file, index=False, engine="openpyxl")
    print(f"Excel file created: {excel_file}")

def process_directory():
    # Automatically process all .gff3 files in the current directory (PWD)
    current_directory = os.getcwd()
    for filename in os.listdir(current_directory):
        if filename.endswith(".gff3"):
            gff3_file = os.path.join(current_directory, filename)
            excel_file = os.path.join(current_directory, f"{os.path.splitext(filename)[0]}_output.xlsx")
            parse_gff3_and_extract_gene_info(gff3_file, excel_file)

if __name__ == "__main__":
    # Process GFF3 files in the current working directory
    process_directory()
