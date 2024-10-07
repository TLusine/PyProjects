import subprocess

# Run 1st script
subprocess.run(['python', 'problem_descriptions.py'])

# Run 2nd script
subprocess.run(['python', 'U_L_extensions.py'])

# Run 3rd script
subprocess.run(['python', 'address_extensions.py'])

# Run 4th script
subprocess.run(['python', 'summary_report.py'])

# Run 5th script, full copy/paste value DB
subprocess.run(['python', 'copy_summary_report.py'])

# Run 6th, full copy/paste value DB
subprocess.run(['python', 'generate_outage_report.py'])
