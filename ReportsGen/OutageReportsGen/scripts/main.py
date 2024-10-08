import subprocess
import logging
import report_builder_gui

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Run 1st script
logging.info("\nRunning copy_addresses.py...")
subprocess.run(['python', 'copy_addresses.py'])

# Run 2nd script
logging.info("\nRunning address_extensions.py...")
subprocess.run(['python', 'address_extensions.py'])


# Run 3rd script
logging.info("Running problem_descriptions.py...")
subprocess.run(['python', 'problem_descriptions.py'])

# Run 4th script
logging.info("\nRunning U_L_extensions.py...")
subprocess.run(['python', 'U_L_extensions.py'])

# Run 5th script
logging.info("\nRunning summary_report.py...")
subprocess.run(['python', 'summary_report.py'])

# Run 6th script, full copy/paste value DB
logging.info("\nRunning copy_summary_report.py...")
subprocess.run(['python', 'copy_summary_report.py'])

# Launch GUI for report builder
logging.info("\nLaunching GUI (report_builder_gui.py)...")
report_builder_gui.main()  # Call the main function of the GUI directly

logging.info("Main script complete.")
