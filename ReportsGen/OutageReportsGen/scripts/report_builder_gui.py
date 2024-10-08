import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry  # Import DateEntry for the calendar feature
import pandas as pd
from generate_outage_report import get_filtered_data, generate_report  # Import the required functions


# Function to generate the report based on input dates
def report_builder_gui(start_date_str, end_date_str):
    # Adjust the date format to match the new pattern
    start_dt = pd.to_datetime(start_date_str, format='%d-%m-%Y')
    end_dt = pd.to_datetime(end_date_str, format='%d-%m-%Y')
    filtered_data = get_filtered_data(start_dt, end_dt)

    if not filtered_data.empty:
        generate_report(start_dt, end_dt, filtered_data)
        messagebox.showinfo("Success", "Report generated successfully.")
    else:
        messagebox.showwarning("No Data", "No data found for the specified date range.")


def generate_report_action(start_date_entry, end_date_entry):
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()
    report_builder_gui(start_date, end_date)


# Create the main window
def create_gui():
    # Create a new window with a custom size
    root = tk.Tk()
    root.title("Outage Report")
    root.geometry("400x300")  # Set a larger window size for better UI spacing

    # Create and place the labels and input fields
    label1 = tk.Label(root, text="Type in Start Date (dd-mm-yyyy) or select from calendar:", font=("Arial", 10))
    label1.pack(pady=10)

    # Use DateEntry with a new date format
    start_date_entry = DateEntry(root, width=20, background='darkblue',
                                 foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
    start_date_entry.pack(pady=5)

    label2 = tk.Label(root, text="Type in End Date (dd-mm-yyyy) or select from calendar:", font=("Arial", 10))
    label2.pack(pady=10)

    end_date_entry = DateEntry(root, width=20, background='darkblue',
                               foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
    end_date_entry.pack(pady=5)

    # Create and place the Generate button
    done_button = tk.Button(root, text='Generate', font=("Arial", 12, "bold"),
                            bg='lightblue', activebackground='blue', activeforeground='white',
                            command=lambda: generate_report_action(start_date_entry, end_date_entry))
    done_button.pack(pady=20)

    # Hover effect on the Generate button
    def on_enter(_=None):
        done_button['background'] = 'blue'
        done_button['foreground'] = 'white'

    def on_leave(_=None):
        done_button['background'] = 'lightblue'
        done_button['foreground'] = 'black'

    # Automatically move focus to the end date entry after entering the start date
    def on_start_date_enter(_=None):
        end_date_entry.focus_set()

    # Bind the hover effects and focus shift
    done_button.bind("<Enter>", on_enter)
    done_button.bind("<Leave>", on_leave)
    start_date_entry.bind('<Return>', on_start_date_enter)

    # Start the GUI event loop
    root.mainloop()


def main():
    create_gui()  # Call the function to create and run the GUI


if __name__ == "__main__":
    main()  # Call the main function when running the script
