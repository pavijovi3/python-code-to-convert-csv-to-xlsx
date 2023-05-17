import os
import glob
import pandas as pd

# Set the folder path containing CSV files
folder_path = r"H:\FTIR_data\DOCUMENTS\background-manupulation-pyprogram\New folder\excel_raw"

# Get a list of all CSV files in the folder
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

# Convert each CSV file to Excel and add it as a sheet
for csv_file in csv_files:
    # Extract the filename without extension
    filename = os.path.splitext(os.path.basename(csv_file))[0]
    
    # Create a new Excel writer object
    output_filename = os.path.join(folder_path, f"{filename}_conv.xlsx")
    writer = pd.ExcelWriter(output_filename, engine="xlsxwriter")

    # Read the CSV file
    df = pd.read_csv(csv_file)

    # Write the DataFrame to Excel
    df.to_excel(writer, sheet_name="Sheet1", index=False)

    # Save the Excel file
    writer._save()

    print(f"Conversion complete for {filename}. Output file saved as: {output_filename}")
