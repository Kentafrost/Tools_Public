import pandas as pd
import time
import os, sys
import boto3
import logging
import upload_local
import common_tool

# parent directory
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)

import common

def read_config(config_path):
    
    config = {}
    if os.path.exists(config_path):
        with open(config_path, 'r') as file:
            for line in file:
                key, value = line.strip().split('=')
                config[key.strip()] = value.strip()
    return config


def main(sheet_name, workbook, remote_chk):
    
    sheet = workbook.worksheet(sheet_name)
    data = sheet.get_all_values()
    df = pd.DataFrame(data)
    extension = ".mp4"

    for index, row in df.iterrows():
        if remote_chk == "y":
            base_path = base_path.replace("D:", "Z:")
            
        if not row[1] == "BasePath":  

            chara_name = row[0].replace(" ","")
            base_path = row[1]
     
            # case 1: local files
            destination_folder = upload_local.folder_path_create(sheet_name, chara_name, base_path, workbook) # define destination folder path
            print(f"Creating local folder: {destination_folder}")

            upload_local.create_folder(destination_folder)
            result = upload_local.move_to_folder(destination_folder, chara_name, extension)
            # delete_path(f'{destination_folder}', chara_name) # to delete unnecessary folder

            # case 2: google drive
            # base_gdrive_path = base_path.replace("D:", "G:\My Drive\Entertainment")
            
            # folder_name = "Entertainment"
            # query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            # results = service.files().list(q=query, fields="files(id, name)").execute()
                            
            # Print folder IDs
            # base_folders_id = results.get('files', [])

            # chara_name = upload_gdrive.folder_name_arrange(base_gdrive_path, chara_name, sheet_name, service)
            # full_path = f'{base_gdrive_path}\{chara_name}'

            # print(f"Creating google drive folder: " + full_path)
            # path_parts = full_path.split("\\")
            
            # Create the folder structure in Google Drive
            # for path in path_parts:
            #     folder_id1 = upload_gdrive.create_folder_gdrive(service, path_parts[3], base_folders_id) 
            #     folder_id2 = upload_gdrive.create_folder_gdrive(service, path_parts[4], folder_id1)  
            #     folder_id3 = upload_gdrive.create_folder_gdrive(service, path_parts[5], folder_id2)  
            #     folder_id4 = upload_gdrive.create_folder_gdrive(service, path_parts[6], folder_id3)  
            #     folder_id5 = upload_gdrive.create_folder_gdrive(service, path_parts[7], folder_id4)

            #folder_metadata = service.files().get(fileId=folder_id, fields='name').execute()
            #folder_name = folder_metadata.get('name')

            #upload_gdrive.move_to_folder_google_drive(folder_id5, chara_name, extension, service)
    if result == "Success":
        msg = f"{sheet_name}: 処理完了"
        logging.info(f"Sheet name({sheet_name}): 処理完了")
    else:
        msg = f"{sheet_name}' 処理失敗"
        logging.info(f"Sheet name({sheet_name}): 処理失敗")
    return msg, base_path


if __name__ == "__main__":

    script_name = "Create_Folder"    
    common.import_log(script_name)
    
    current_time = time.strftime("%Y%m%d_%H%M%S")
    logging.info(f"{current_time}: Script({script_name}) started.")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config = read_config(f'{current_dir}\\config.txt')

    if config.get("flg_filepath"):
        flg_filepath = config.get("flg_filepath")
    else:
        flg_filepath = input("Do you want to list file path? (y/n): ")
    
    if config.get("remote_chk"):
        remote_chk = config.get("remote_chk")
    else:
        remote_chk = input("Are you accessing remotely? (y/n): ")
    
    gc = common.authorize_gsheet() # google Authorizations
    workbook = gc.open("chara_name_list")
    
    sheet_name_list = []
    sheets = workbook.worksheets()
    
    # make list with all sheet names
    for sheet in sheets:
        sheet_name_list.append(sheet.title)
    print(sheet_name_list)
    
    msg_list = []
    folder_list = []

    for sheet in sheet_name_list:
        
        list = main(sheet, workbook, remote_chk)
        
        logging.info(f"Processing {sheet} is complete.")
        print(list[0])
        
        time.sleep(3)
        
        msg_list.append(list[0])
        # make the base paths list in google spreadsheet
        folder_list.append(list[1])
            
    message = str(msg_list).encode('utf-8').decode('utf-8')
    print(message)
    
    # save files path in google spreadsheet
    file_workbook = gc.open("file_list")
    i = 0
    
    if flg_filepath == "y":
        for sheet_name in sheet_name_list:
            print(f"sheet_name: {sheet_name}")

            try:
                # check if the sheet exists
                common_tool.check_sheet_exists(sheet_name, file_workbook)
                sheet = file_workbook.worksheet(sheet_name)
            except Exception as e:
                print(f"Error: {e}")
                logging.error(f"Error: {e}")
            
            print(folder_list[i])
            common_tool.listup_all_files(folder_list[i], sheet)
            i = i + 1
    
    try:
        ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
        common_tool.send_mail(ssm_client, msg_list) # if required
    except Exception as e:
        print(f"Error sending email: {e}")
        logging.error(f"Error sending email: {e}")

    current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print(f"{current_time}: All processes are complete.")
    logging.info(f"{current_time}: All processes are complete.")

    os.system("shutdown /s /t 1800")
    current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    logging.info(f"{current_time}: Shutdown command executed and successfully completed.")