import subprocess
import logging
import report_builder_gui
import os
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
python_executable = sys.executable  # Get the path to the Python interpreter


# Function to run scripts and log their output
def run_script(script_name):
    try:
        logging.info(f"\nRunning {script_name}...")
        script_path = os.path.join(script_dir, script_name)  # Full path to the script
        result = subprocess.run([python_executable, script_path],
                                capture_output=True, text=True)

        # Log stdout and stderr
        logging.info(result.stdout)
        logging.error(result.stderr)

    except Exception as e:
        logging.error(f"Error occurred while running {script_name}: {e}")


# Run all scripts
scripts = [
    'copy_addresses.py',
    'address_extensions.py',
    'problem_descriptions.py',
    'U_L_extensions.py',
    'summary_report.py',
    'copy_summary_report.py'
]

for script in scripts:
    run_script(script)

# Launch GUI for report builder
logging.info("\nLaunching GUI (report_builder_gui.py)...")
report_builder_gui.main()  # Call the main function of the GUI directly

logging.info("Main script complete.")
