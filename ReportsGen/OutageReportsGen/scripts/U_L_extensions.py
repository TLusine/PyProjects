import pandas as pd


class DurationCalculator:
    @staticmethod
    def calculate_duration(start_date, start_time, end_date, end_time):
        start = pd.to_datetime(f"{start_date} {start_time}")
        end = pd.to_datetime(f"{end_date} {end_time}")
        duration = end - start
        total_minutes = int(duration.total_seconds() // 60)
        hours, minutes = divmod(total_minutes, 60)
        return f"{hours:02}:{minutes:02}"


class DateTimeFormatter:
    @staticmethod
    def format_dates_times(original_row):
        formatted_start_time = pd.to_datetime(
            f"{original_row['Start Data for Luso']} {original_row['Start Time']}").strftime('%H:%M')
        formatted_end_time = pd.to_datetime(
            f"{original_row['End Date for Luso']} {original_row['End Time']}").strftime('%H:%M')
        formatted_start_date = pd.to_datetime(original_row['Start Data for Luso']).strftime('%Y-%m-%d')
        formatted_end_date = pd.to_datetime(original_row['End Date for Luso']).strftime('%Y-%m-%d')
        return formatted_start_time, formatted_end_time, formatted_start_date, formatted_end_date


class SiteCodeExtender:
    @staticmethod
    def extend_site_code(original_row, extended_rows):
        site_code = original_row['Site Code']
        for suffix in ['U', 'L']:
            extended_row = original_row.copy()
            extended_row['Site Code'] = f"{site_code}{suffix}"
            SiteCodeExtender.clear_redundant_columns(extended_row)
            extended_rows.append(extended_row)

    @staticmethod
    def clear_redundant_columns(row):
        row['Affected Sites'] = None
        row['Number Of Affected Sites'] = None
        row['Solved By:'] = None
        row['Additional Information'] = None


class DataProcessor:
    def __init__(self, input_df):
        self.input_df = input_df
        self.processed_rows = []

    def process_data(self):
        for _, original_row in self.input_df.iterrows():
            self.process_row(original_row)

        return self.processed_rows

    def process_row(self, original_row):
        duration = DurationCalculator.calculate_duration(
            original_row['Start Data for Luso'], original_row['Start Time'],
            original_row['End Date for Luso'], original_row['End Time']
        )
        start_time, end_time, start_date, end_date = DateTimeFormatter.format_dates_times(original_row)

        final_row = original_row.copy()
        final_row['Duration'] = duration
        final_row['Start Time'] = start_time
        final_row['End Time'] = end_time
        final_row['Start Data for Luso'] = start_date
        final_row['End Date for Luso'] = end_date
        self.processed_rows.append(final_row)

        SiteCodeExtender.extend_site_code(final_row, self.processed_rows)

        self.process_affected_sites(original_row, duration, start_time, end_time, start_date, end_date)

    def process_affected_sites(self, original_row, duration, start_time, end_time, start_date, end_date):
        if pd.notna(original_row['Affected Sites']):
            affected_sites = original_row['Affected Sites'].split(', ')
            for affected_site in affected_sites:
                affected_site_row = original_row.copy()
                affected_site_row['Site Code'] = affected_site
                SiteCodeExtender.clear_redundant_columns(affected_site_row)
                affected_site_row['Duration'] = duration
                affected_site_row['Start Time'] = start_time
                affected_site_row['End Time'] = end_time
                affected_site_row['Start Data for Luso'] = start_date
                affected_site_row['End Date for Luso'] = end_date
                self.processed_rows.append(affected_site_row)
                SiteCodeExtender.extend_site_code(affected_site_row, self.processed_rows)


def main():
    # Read the input Excel file from the specified sheet "FromDB"
    df = pd.read_excel('ReportsDB.xlsx', sheet_name='DB_ProblemsDescriptions')

    # Process data and generate final rows
    data_processor = DataProcessor(df)
    final_data = data_processor.process_data()

    # Convert final rows into a DataFrame
    result_df = pd.DataFrame(final_data)

    # Write the result to a new sheet named 'DB_U-L_extended' in the same Excel file
    with pd.ExcelWriter('../data/ReportsDB.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        result_df.to_excel(writer, sheet_name='DB_U-L_extended', index=False)

    print("Extended Addresses with U(3G) and L(4G) have been added in DB_U-L_extended sheet")


if __name__ == "__main__":
    main()
