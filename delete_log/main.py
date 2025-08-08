import logging
import os, sys
import datetime

# set up logging
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)

import common
common.import_log("Delete-Log-Files")

log_fold = os.path.join(parent_dir, "log")
if not os.path.exists(log_fold):
    print("Log folder does not exist, so skip the operation")
    os._exit(0)
    
# list up all log files in each subdirectory
log_files = []
current_date = datetime.datetime.now().strftime("%Y%m%d")

# target date before today, change this as needed
target_date = (datetime.datetime.now() - datetime.timedelta(days=5)).strftime("%Y%m%d")

for root, dirs, files in os.walk(log_fold):
    print(f"Processing directory: {root}")
    print(f"Files found: {files}")
    
    for file in files:
        if file.endswith(".log"):
            # retrieve date from file name
            log_files.append(os.path.join(root, file))
            date_string = file.title()[:8]
            
            # convert date string to datetime object for comparison
            try:
                date_string = datetime.datetime.strptime(date_string, "%Y%m%d").strftime("%Y%m%d")
            except ValueError:
                logging.warning(f"Invalid date format in log file name: {file}")
                continue

            # compare date strings
            if date_string < target_date:
                logging.info(f"Found log file over 5 days before today: {file}")

                os.remove(os.path.join(root, file))
                logging.info(f"Deleted log file: {file}")
                
logging.info(f"Log files processed. Total files deleted: {len(log_files)}")
logging.info("Log cleanup completed successfully.")