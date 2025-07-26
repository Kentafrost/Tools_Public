import os
import win32com.client

# python automation script to create a task scheduler
task_create_name = "CreateFolderTask"

scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')

tasks = root_folder.GetTasks(0)
tasks_names = [task.Name for task in tasks]

for tasks_name in tasks_names:
    if tasks_name == task_create_name:
        exist_chk = True
        break

if not exist_chk == True:
    task_def = scheduler.NewTask(0)
    task_def.RegistrationInfo.Description = "CreateFolder Task"

    task_def.Actions.Create(0).Path = r"G:\\My Drive\\IT_Learning\\Git\\tools\\Create_Folder\\main.py"
    task_def.Settings.Enabled = True
    task_def.Settings.StopIfGoingOnBatteries = False

    # Task def
    trigger = task_def.Triggers.Create(2)  # 2: daily trigger
    trigger.StartBoundary = "2025-07-21T23:00:00"
    trigger.DaysInterval = 1  # 毎日

    root_folder.RegisterTaskDefinition(
        task_create_name,
        task_def,
        6,  # create or update
        "", "",  # run as current user
        3  # logon type: interactive
    )

    print("Task Schedules creation completed")
    os.system("taskschd.msc")
else:
    print(f"This task schedule already exists. Script name: {task_create_name}")