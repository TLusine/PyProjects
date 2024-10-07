import pandas as pd


class ProblemDescriptionGenerator:
    def __init__(self, file_path):
        # Load the Excel file and read relevant sheets into DataFrames
        self.from_db_df = pd.read_excel(file_path, sheet_name='FromDB')
        addresses_df = pd.read_excel(file_path, sheet_name='Addresses')
        self.address_mapping = addresses_df.set_index('Site Code')['Site_Code_Address_eng'].to_dict()

    def generate_problem_description_eng(self, site_code, affected_sites, common_problem,
                                         problem_title, disturbed_link, solved_by):
        # English problem description logic
        if pd.isna(affected_sites) or affected_sites.strip() == "":
            if common_problem in ["ENA Power Off", "Cabinet/Shelter Problem"]:
                return f"Connection loss in {site_code} due to discharge of battery."
            elif common_problem == "Not Detected":
                return f"Problem was not detected. {site_code} recovered automatically."
            elif common_problem == "Transmission Problem":
                return self.transmission_problem_eng(site_code, problem_title, disturbed_link, solved_by)
            elif common_problem == "RAN Problem":
                return self.ran_problem_eng(site_code, problem_title, disturbed_link, solved_by)
        else:
            return self.non_empty_affected_sites_eng(site_code, common_problem,
                                                     problem_title, disturbed_link, solved_by)
        return ""

    def generate_problem_description_rus(self, site_code, affected_sites, common_problem,
                                         problem_title, disturbed_link, solved_by):
        # Russian problem description logic
        if pd.isna(affected_sites) or affected_sites.strip() == "":
            if common_problem in ["ENA Power Off", "Cabinet/Shelter Problem"]:
                return f"Потеря соединения на {site_code}."
            elif common_problem == "Not Detected":
                return f"Причина неизвестна. {site_code} восстановилась автоматически."
            elif common_problem == "Transmission Problem":
                return self.transmission_problem_rus(site_code, problem_title, disturbed_link, solved_by)
            elif common_problem == "RAN Problem":
                return self.ran_problem_rus(site_code, problem_title, disturbed_link, solved_by)
        else:
            return self.non_empty_affected_sites_rus(site_code, common_problem,
                                                     problem_title, disturbed_link, solved_by)
        return ""

    def transmission_problem_eng(self, site_code, problem_title, disturbed_link, solved_by):
        if problem_title == "FO Cut":
            return (f"Failure of {disturbed_link} link due to FO cut. "
                    "The connection was recovered after the replacement of damaged FO part.")
        else:
            base_message = (f"Failure of {disturbed_link} link due to problems with {problem_title} "
                            f"on {site_code}.")
            return self.solve_message_eng(base_message, problem_title, solved_by)

    def transmission_problem_rus(self, site_code, problem_title, disturbed_link, solved_by):
        if problem_title == "FO Cut":
            return (f"Сбой работы пролета {disturbed_link} из-за обрыва оптико-волоконного кабеля. "
                    "Связь восстановилась после ремонта поврежденного участка.")
        else:
            base_message = f"Сбой работы пролета {disturbed_link} из-за проблем с {problem_title} на {site_code}."
            return self.solve_message_rus(base_message, problem_title, solved_by)

    def ran_problem_eng(self, site_code, problem_title, disturbed_link, solved_by):
        base_message = f"Failure of {disturbed_link} due to {problem_title} on {site_code}."
        return self.solve_message_eng(base_message, problem_title, solved_by)

    def ran_problem_rus(self, site_code, problem_title, disturbed_link, solved_by):
        base_message = f"Сбой работы пролета {disturbed_link} из-за проблем с {problem_title} на {site_code}."
        return self.solve_message_rus(base_message, problem_title, solved_by)

    @staticmethod
    def solve_message_eng(base_message, problem_title, solved_by):
        if pd.notna(solved_by) and isinstance(solved_by, str):
            if "Hard Restart" in solved_by:
                return f"{base_message} Recovered after hard restart of {problem_title}."
            elif "Software Restart" in solved_by:
                return f"{base_message} Recovered after software restart of {problem_title}."
            elif "Change" in solved_by:
                return f"{base_message} Recovered after replacement of {problem_title}."
            elif "Reconfig" in solved_by:
                return f"{base_message} Recovered after reconfiguration of {problem_title}."
            elif "Autorecovery" in solved_by:
                return f"{base_message} Recovered automatically."
        return base_message

    @staticmethod
    def solve_message_rus(base_message, problem_title, solved_by):
        if pd.notna(solved_by) and isinstance(solved_by, str):
            if "Hard Restart" in solved_by:
                return f"{base_message} Соединение восстановлено после ручной перезагрузки {problem_title}."
            elif "Software Restart" in solved_by:
                return f"{base_message} Соединение восстановлено после перезагрузки ПО {problem_title}."
            elif "Change" in solved_by:
                return f"{base_message} Соединение восстановлено после замены {problem_title}."
            elif "Reconfig" in solved_by:
                return f"{base_message} Соединение восстановлено после реконфигурации {problem_title}."
            elif "Autorecovery" in solved_by:
                return f"{base_message} Соединение восстановлено автоматически."
        return base_message

    def non_empty_affected_sites_eng(self, site_code, common_problem, problem_title, disturbed_link, solved_by):
        if common_problem in ["ENA Power Off", "Cabinet/Shelter Problem"]:
            return f"Failure of {disturbed_link} due to discharge of battery on {site_code}."
        return self.transmission_problem_eng(site_code, problem_title, disturbed_link, solved_by)

    def non_empty_affected_sites_rus(self, site_code, common_problem, problem_title, disturbed_link, solved_by):
        if common_problem in ["ENA Power Off", "Cabinet/Shelter Problem"]:
            return f"Сбой работы пролета {disturbed_link} из-за разрядки батареи на {site_code}."
        return self.transmission_problem_rus(site_code, problem_title, disturbed_link, solved_by)

    def generate_final_data(self):
        results = []
        for _, row in self.from_db_df.iterrows():
            unique_id = row['ID']
            site_code = row['Site Code']
            affected_sites_value = row['Affected Sites']
            common_problem_value = row['Common Problem']
            problem_title_value = row['Problem Title']
            disturbed_link_value = row['Disturbed Link']
            solved_by_value = row['Solved By:']

            if common_problem_value in ["ENA Power Off", "Transmission Problem", "RAN Problem", "Not Detected",
                                        "Cabinet/Shelter Problem"] and site_code in self.address_mapping:
                problem_description_eng = self.generate_problem_description_eng(
                    site_code, affected_sites_value, common_problem_value, problem_title_value, disturbed_link_value,
                    solved_by_value)
                problem_description_rus = self.generate_problem_description_rus(
                    site_code, affected_sites_value, common_problem_value, problem_title_value, disturbed_link_value,
                    solved_by_value)

                results.append([
                    unique_id, site_code, affected_sites_value, row['Number Of Affected Sites'], row['Problem Type'],
                    common_problem_value, problem_title_value, disturbed_link_value, row['Problem Location'],
                    solved_by_value, row['Additional Information'], row['Start Data for Luso'], row['Start Time'],
                    row['End Date for Luso'], row['End Time'], row['Duration'],
                    problem_description_eng, problem_description_rus
                ])
        return results  # Return the collected results


def main():
    # Main code execution
    input_file_path = '../data/ReportsDB.xlsx'
    generator = ProblemDescriptionGenerator(input_file_path)
    final_data = generator.generate_final_data()

    # Create a new DataFrame and save to Excel
    columns = [
        'ID', 'Site Code', 'Affected Sites', 'Number Of Affected Sites', 'Problem Type', 'Common Problem',
        'Problem Title', 'Disturbed Link', 'Problem Location', 'Solved By:', 'Additional Information',
        'Start Data for Luso', 'Start Time', 'End Date for Luso', 'End Time', 'Duration',
        'ProblemDescription_eng', 'ProblemDescription_rus'
    ]
    problems_description_df = pd.DataFrame(final_data, columns=columns)

    with pd.ExcelWriter(input_file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        problems_description_df.to_excel(writer, sheet_name='DB_ProblemsDescriptions', index=False)

    print("Descriptions have been successfully updated in the 'ProblemsDescriptions' sheet.")


if __name__ == "__main__":
    main()
