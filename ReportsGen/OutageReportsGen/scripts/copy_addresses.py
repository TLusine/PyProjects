import openpyxl
import pandas as pd

def copy_addresses(source_filename='../data/ReportsDB.xlsx'):
    # Load the workbook and the specified sheet
    workbook = openpyxl.load_workbook(source_filename, data_only=True)
    sheet = workbook['Addresses']

    # Extract the data from the sheet, including header
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    # Convert the data to a DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])  # Skip the header
    # Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # Select only the required columns
    columns_to_keep = ['Site Code', 'Site_Code_Address_rus', 'Site_Code_Address_eng']

    # Filter the data to keep only the specified columns
    filtered_data = df[columns_to_keep]

    # Use ExcelWriter to overwrite the data in the 'copy_addresses' sheet
    with pd.ExcelWriter(source_filename, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        filtered_data.to_excel(writer, sheet_name='copy_addresses', index=False)
        print(f"Copied values data to '{source_filename}' successfully in the 'copy_addresses' sheet.")

if __name__ == "__main__":
    copy_addresses()
