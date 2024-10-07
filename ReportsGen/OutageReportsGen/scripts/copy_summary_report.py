import pandas as pd


def copy_summary_report(source_filename='ReportsDB.xlsx'):
    # Load the data from the existing Excel file
    report_data = pd.read_excel(source_filename, sheet_name='summary_report', header=0)

    # Check if report_data is loaded correctly
    if report_data.empty:
        print("The DataFrame is empty. Please check your Excel file.")
        return

    # Log the number of rows loaded
    print(f"Loaded {len(report_data)} rows from '{source_filename}'.")

    # Use ExcelWriter to append the data to the existing Excel file
    with pd.ExcelWriter(source_filename, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        report_data.to_excel(writer, sheet_name='copy_summary_report', index=False)
        print(f"Copied data to '{source_filename}' successfully in the 'copy_summary_report' sheet.")


if __name__ == "__main__":
    copy_summary_report()
