import logging
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os, re, shutil
import common_tool
import boto3
import time

# Function to create a folder path based on character name and base folder
def folder_path_create(sheet_name, chara_name, base_folder, workbook):
    
    # to retrive data from SSM, use worksheets index
    ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
    worksheets = workbook.worksheets()
    sheet_index = None
    
    for idx, sheet in enumerate(worksheets):
        if sheet.title == sheet_name:
            sheet_index = idx
            break
    
    if sheet_index is not None:
        sheet_index = sheet_index + 1  # Adjust index to match SSM parameter naming convention
        
        try:
            title = ssm_client.get_parameter(Name=f'Title{sheet_index}', WithDecryption=True)['Parameter']['Value']    
        except Exception as e:
            print(f"Error retrieving title from SSM: {e}")
            print("Stop the process.")
            os._exit(1)
    
    title5 = ssm_client.get_parameter(Name='Title5', WithDecryption=True)['Parameter']['Value']
    title8 = ssm_client.get_parameter(Name='Title8', WithDecryption=True)['Parameter']['Value']

    if "【" in chara_name and sheet_name == title5:
        pattern = r"【(.*)】"
        extract_txt = common_tool.get_chara_name_between(chara_name, pattern)
        chara_name = chara_name.replace(f"【{extract_txt}】", "")
        
    elif "(" in chara_name:
        pattern = r"((.*))"
        extract_txt = re.search(r"\((.*?)\)", chara_name)
        extract_txt = extract_txt.group(1)
        chara_name = chara_name.replace(f"({extract_txt})", "")
        
    elif "【" in chara_name and not sheet_name == title5:
        pattern = r"【(.*)】"
        chara_name = common_tool.get_chara_name_between(chara_name, pattern)
        
    elif sheet_name == title8:
        match = re.search(r"】(.+)", chara_name)  # 「】」の後の文字を取得
        chara_name = match.group(1) if match else ""
        chara_name = re.sub(r"\(.*?\)", "", chara_name) # ()内の文字を削除
        
    destination_folder = f'{base_folder}\{chara_name}'
    destination_folder = destination_folder.replace(" ", "")

    time.sleep(0.5)
    return destination_folder

    # if "【" in chara_name and sheet_name == title:
    #     pattern = r"【(.*)】"
    #     extract_txt = common_tool.get_chara_name_between(chara_name, pattern)
    #     chara_name = chara_name.replace(f"【{extract_txt}】", "")
        
    # elif "(" in chara_name:
    #     pattern = r"((.*))"
    #     extract_txt = re.search(r"\((.*?)\)", chara_name)
    #     extract_txt = extract_txt.group(1)
    #     chara_name = chara_name.replace(f"({extract_txt})", "")
        
    # elif "【" in chara_name and not sheet_name == title and not title in sheet_name:
    #     pattern = r"【(.*)】"
    #     chara_name = common_tool.get_chara_name_between(chara_name, pattern)
        
    # elif title == sheet_name:
    #     match = re.search(r"】(.+)", chara_name)
    #     chara_name = match.group(1) if match else ""
    #     chara_name = re.sub(r"\(.*?\)", "", chara_name)


    # destination_folder = f'{base_folder}\{chara_name}'
    # destination_folder = destination_folder.replace(" ", "")



# create folder, if it exists already, then nothing happens. Files in a folder untouched
def create_folder(folder_path):
    print('--------------------------------')
    print(f'Folder name: {folder_path}')
    print('--------------------------------')
    
    time.sleep(0.5)

    try:
        os.makedirs(rf"{folder_path}", exist_ok=True)
    except Exception as e:
        print(f"Error creating folder: {e}")


# move the file to the destination folder and rename it if necessary
def move_and_rename_file(src_file, chara_name, dest_folder):
    
    time.sleep(0.5)
    # Get the original file name and extension
    file_name, file_extension = os.path.splitext(os.path.basename(src_file))
    logging.info(f"Original file name: {file_name}, Extension: {file_extension}")
    logging.info(f"Destination folder: {dest_folder}")
 
    dest_path = os.path.join(dest_folder, file_name + file_extension)
    logging.info(f"Destination path: {dest_path}")

    # Check if a file with the same name already exists
    count = 1
    while os.path.exists(dest_path):
        # Append a number to the file name
        new_file_name = f"{file_name}_{count}{file_extension}"
        dest_path = os.path.join(dest_folder, new_file_name)
        count += 1

    # Move the file to the destination folder
    #os.rename(src_file, dest_path)
    try:
        print(src_file)
        print(dest_path)
        shutil.move(src_file, dest_path)
        logging.info(f"Moved file from {src_file} to {dest_path}")
    except Exception as e:
        logging.error(f"Error moving file: {e}")
        logging.info(f"Failed to move file from {src_file} to {dest_path}")


# extract, adjust words to make folder path
# Move To folder if mp4 video already exists
def move_to_folder(dest_directory, charaname, extension):
    dest_directory = f"{dest_directory}\\"
    
    if not os.path.exists(dest_directory): # Ensure the destination folder exists
        os.makedirs(dest_directory)

    ssm = boto3.client('ssm', region_name='ap-southeast-2')
    source_fold = ssm.get_parameter(Name='DownloadPath', WithDecryption=True)['Parameter']['Value']
    
    # Search for files in the source directory to move to
    try:
        for filename in os.listdir(source_fold):
            if filename.endswith(extension):
                chk_name = filename.replace(extension, "") # with no extension
                print(f"charaname: {chk_name} in {charaname}")
                
                # check if files name ends with mp4 includes characters name in csv name
                if charaname in chk_name: 
                    print(f"Found file: {filename}")
                    source_file = f'{source_fold}\{filename}' # files path where you want to move from
                    
                    # delete some words in the filename such as " - Made with Clipchamp"
                    filename = common_tool.name_converter(source_fold, filename)
                    print(f"File name after conversion: {filename}")

                    try:
                        source_file = rf'{source_fold}{filename}' # files path where you want to move from
                        destination_file = f'{dest_directory}\\' # To directory with character name
                    except Exception as e:
                        print(f"Error: {e}")
                        print(f"Source file: {source_file} To")
                        print(f"Destination folder: {destination_file}")
                        
                    # Check if the file already exists in the destination folder
                    # If it does, rename the file
                    move_and_rename_file(source_file, charaname, destination_file)
            time.sleep(0.5)
        msg = "Success"
    except Exception as e:
        print(f"Error: {e}")
        msg = "Error"
    return msg