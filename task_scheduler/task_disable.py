import os
import time
import win32com.client
import logging

# disable or enable task scheduler
disable_chk = True  

# target task name. change this in cases
target_task_name = "CreateFolderTask"

# current directory
current_dir = os.path.dirname(os.path.abspath(__file__))
current_time = time.strftime("%Y%m%d_%H%M%S")

scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')
tasks = root_folder.GetTasks(0)
tasks_names = [task.Name for task in tasks]

logging.basicConfig(filename=f"{current_dir}\\log\\{current_time}\\task.log", 
                    level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def disable_task_scheduler(tasks_name):

    task = root_folder.GetTask(tasks_name)
    task.Enabled = False
    root_folder.RegisterTaskDefinition(tasks_name, task.Definition, 6, None, None, 3, None)
    logging.info(f"Disabled Task: '{tasks_name}'")
    
def enable_task_scheduler(tasks_name):
    
    task = root_folder.GetTask(tasks_name)
    task.Enabled = True
    root_folder.RegisterTaskDefinition(tasks_name, task.Definition, 6, None, None, 3, None)
    logging.info(f"Enabled Task: '{tasks_name}'")
    
    
if __name__ == "__main__":
    
    if target_task_name not in tasks_names:
        logging.error(f"Task '{target_task_name}' not found.")
    else:
        for tasks_name in tasks_names:
            if tasks_name == target_task_name:
                if disable_chk == True:
                    disable_task_scheduler(tasks_name)
                else:
                    enable_task_scheduler(tasks_name)
                break
        logging.info("Task scheduler switch operation completed.")
        logging.shutdown()