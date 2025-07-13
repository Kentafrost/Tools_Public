import re, os
from googleapiclient.http import MediaFileUpload
import common_tool

# Function to find folder ID by path
def get_folder_id_gdrive(service, folder_path):
    folder_names = folder_path.strip('/').split('/')  # Split the path into folder names
    parent_id = 'root'  # Start from the root directory

    for folder_name in folder_names:
        query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        folders = results.get('files', [])
        if not folders:
            raise Exception(f"Folder '{folder_name}' not found under parent ID '{parent_id}'.")
        # Assume the first folder found is the correct one (if duplicates exist)
        parent_id = folders[0]['id']
    return parent_id

# Function to create a folder
def create_folder(service, name, parent_id=None):
    file_metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id:
        file_metadata['parents'] = [parent_id]
    folder = service.files().create(body=file_metadata, fields='id').execute()
    print(f"Created folder: {name} with ID: {folder['id']}")
    return folder['id']

def folder_name_arrange(base_folder, chara_name, sheet_name, service):
    
    if "【" in chara_name and sheet_name == "KFantasy":
        pattern = r"【(.*)】"
        extract_txt = common_tool.get_chara_name_between(chara_name, pattern)
        chara_name = chara_name.replace(f"【{extract_txt}】", "")
        
    elif "(" in chara_name:
        pattern = r"((.*))"
        extract_txt = re.search(r"\((.*?)\)", chara_name)
        extract_txt = extract_txt.group(1)
        chara_name = chara_name.replace(f"({extract_txt})", "")
        
    elif "【" in chara_name and not sheet_name == "KFantasy":
        pattern = r"【(.*)】"
        chara_name = common_tool.get_chara_name_between(chara_name, pattern)


def move_to_folder_google_drive(dest_folder_id, charaname, extension, service):
    extension = extension.replace(".", "")

    # from download in local environment
    source_fold = rf"C:\Users\user\Downloads"

    # Search for files in the local source directory
    for filename in os.listdir(source_fold):
        if charaname in filename and filename.endswith(extension):

            # File to upload into google drives specific folder
            file_metadata = {
                'name': f'{filename}',
                'parents': [f'{dest_folder_id}']
            }
            media = MediaFileUpload(f'{filename}', mimetype=f'video/{extension}')
            print(f'Uploading {filename} to {dest_folder_id}')

            # Upload the file
            service.files().create(body=file_metadata, media_body=media, fields='id').execute()