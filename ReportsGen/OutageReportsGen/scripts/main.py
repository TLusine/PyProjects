import subprocess

# Run the first script
subprocess.run(['python', 'problem_descriptions.py'])

# Run the second script
subprocess.run(['python', 'U_L_extensions.py'])

# Run the second script
subprocess.run(['python', 'address_extensions.py'])

# Run the second script
subprocess.run(['python', 'summary_report.py'])

# Run the second script, full copy/paste value DB
subprocess.run(['python', 'copy_summary_report.py'])

# Run the second script, full copy/paste value DB
subprocess.run(['python', 'generate_outage_report.py'])
