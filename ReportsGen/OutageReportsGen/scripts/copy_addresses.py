import pandas as pd

# Load the Excel file
file_path = '../data/ReportsDB.xlsx'
excel_file = pd.ExcelFile(file_path)

# Read the Addresses sheet
addresses_df = excel_file.parse('Addresses')

# Create a new DataFrame for the copy_addresses sheet
copy_addresses_df = pd.DataFrame()

# Populate the Site Code column
copy_addresses_df['Site Code'] = addresses_df['Site Code']

# Populate the Site_Code_Address_rus column
copy_addresses_df['Site_Code_Address_rus'] = (
    'BTS - ' + addresses_df['Site Code'] + ', ' +
    addresses_df['rus_1'] + addresses_df['rus_2']
)

# Populate the Site_Code_Address_eng column
copy_addresses_df['Site_Code_Address_eng'] = (
    'BTS - ' + addresses_df['Site Code'] + ', ' +
    addresses_df['eng_1'] + ', ' +
    addresses_df['eng_2']
)

# Save the results to the copy_addresses sheet in the same Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    copy_addresses_df.to_excel(writer, sheet_name='copy_addresses', index=False)

print("Data copied successfully to 'copy_addresses' sheet.")
