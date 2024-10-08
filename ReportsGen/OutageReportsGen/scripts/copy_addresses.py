import pandas as pd


def copy_addresses(source_filename='../data/ReportsDB.xlsx'):
    # Load the data from the existing Excel file
    report_data = pd.read_excel(source_filename, sheet_name='Addresses', header=0)

    # Check if report_data is loaded correctly
    if report_data.empty:
        print("The DataFrame is empty. Please check your Excel file.")
        return

    # Log the number of rows loaded
    print(f"Loaded {len(report_data)} rows from '{source_filename}'.")

    # Select only the required columns
    columns_to_keep = ['Site Code', 'Site_Code_Address_rus', 'Site_Code_Address_eng']
    filtered_data = report_data[columns_to_keep]

    # Use ExcelWriter to append the filtered data to the existing Excel file
    with pd.ExcelWriter(source_filename, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        filtered_data.to_excel(writer, sheet_name='copy_addresses', index=False)
        print(f"Copied filtered data to '{source_filename}' successfully in the 'copy_addresses' sheet.")


if __name__ == "__main__":
    copy_addresses()
