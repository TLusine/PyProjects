import pandas as pd

# Load the existing Excel file and read the 'Addresses' sheet into a DataFrame
input_file_path = '../data/ReportsDB.xlsx'
addresses_df = pd.read_excel(input_file_path, sheet_name='copy_addresses')

# Strip whitespace from the column names (optional)
addresses_df.columns = addresses_df.columns.str.strip()

# Create a new DataFrame to hold the extended addresses
extended_data = []

# Iterate through each row in the original DataFrame
for _, row in addresses_df.iterrows():
    site_code = row['Site Code']
    site_code_address_rus = row['Site_Code_Address_rus']
    site_code_address_eng = row['Site_Code_Address_eng']

    # Original entry
    extended_data.append([site_code, site_code_address_rus, site_code_address_eng])

    # Variants with suffixes 'U' and 'L'
    for suffix in ['U', 'L']:
        extended_data.append([
            f"{site_code}{suffix}",
            site_code_address_rus.replace(site_code, f"{site_code}{suffix}"),
            site_code_address_eng.replace(site_code, f"{site_code}{suffix}")
        ])

# Create a new DataFrame with the extended data
extended_df = pd.DataFrame(extended_data, columns=['Site Code', 'Site_Code_Address_rus', 'Site_Code_Address_eng'])

# Write the new DataFrame to a new sheet in the same Excel file
with pd.ExcelWriter(input_file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    extended_df.to_excel(writer, sheet_name='Add_extended', index=False)

print("Extended addresses have been successfully added to the 'Add_extended' sheet.")
