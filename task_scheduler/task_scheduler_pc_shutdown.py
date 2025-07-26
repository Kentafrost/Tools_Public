import os
import win32com.client

# python automation script to create a task scheduler
task_create_names = ["PCShutdownTaskWeekdays", "PCShutdownTaskWeekends"]
action_path = r"C:\\Windows\\System32\\shutdown.exe -s -t 3600"

scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')

tasks = root_folder.GetTasks(0)
tasks_names = [task.Name for task in tasks]

# Check if the tasks already exist in the task schedulers
for task_create_name in task_create_names:
    if task_create_name in tasks_names:

        if not task_create_name == "PCShutdownTaskWeekdays":
            
            print(f"Task '{task_create_name}' doesn't exist.")
            task_def = scheduler.NewTask(0)
            task_def.RegistrationInfo.Description = "PC Shutdown Task."

            task_def.Actions.Create(0).Path = action_path
            task_def.Settings.Enabled = True
            task_def.Settings.StopIfGoingOnBatteries = False

            # Task def
            trigger = task_def.Triggers.Create(2)  # 2: daily trigger
            trigger.StartBoundary = "2025-07-21T23:00:00"
            trigger.DaysOfWeek = 62  # 62: Monday to Friday (binary 111110)

            root_folder.RegisterTaskDefinition(
                task_create_name,
                task_def,
                6,  # create or update
                "", "",  # run as current user
                3  # logon type: interactive
            )
        elif task_create_name == "PCShutdownTaskWeekends":
            print(f"Task '{task_create_name}' already exists.")
            task_def = scheduler.NewTask(0)
            task_def.RegistrationInfo.Description = "PC Shutdown Task on weekends"

            task_def.Actions.Create(0).Path = action_path
            task_def.Settings.Enabled = True
            task_def.Settings.StopIfGoingOnBatteries = False

            # Task def
            trigger = task_def.Triggers.Create(2)
        else:
            print(f"This task schedule already exists. Script name: {task_create_name}")


print("Task Schedules creation completed")
os.system("taskschd.msc")
