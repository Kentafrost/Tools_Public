import gspread
import os, re, logging
import smtplib

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import boto3
import pandas as pd
import time
import logging

def get_chara_name_between(chara_name, pattern):
    match = re.search(pattern, chara_name)
    if match:
        return match.group(1)
    return None


# Check if the folder exists, then delete it(if nesessary, comment in)
def delete_path(folder_path, word):
    print(f'{folder_path}\{word}')
    
    # if os.path.exists(f'{folder_path}\{word}') and os.path.isdir(f'{folder_path}\{word}'):
    #     try:
    #         shutil.rmtree(f'{folder_path}\{word}')  # Deletes the folder and its contents
    #         print(f"Folder '{folder_path}\{word}' has been deleted.")
    #     except Exception as e:
    #         print(f"Failed to delete '{folder_path}\{word}': {e}")
    # else:
    #     print(f"Folder '{folder_path}\{word}' does not exist at the specified path.")

def send_mail(client, msg_list):
    
    logging.info('メール送信処理を開始します。')
    try:
        subject = "Report"
        bodyText = "Here's the report:\n" + "\n" + "\n".join(map(str, msg_list))
        
        # メールの内容(SSMから取得)
        from_address = client.get_parameter(Name='my_main_gmail_address', WithDecryption=True)['Parameter']['Value']
        from_pw = client.get_parameter(Name='my_main_gmail_password', WithDecryption=True)['Parameter']['Value']  
        to_address = from_address

    except Exception as e:
        logging.error('メールの内容をSSMから取得できませんでした。{}'.format(e))
    
    
    try:
        message = f"Subject: {subject}\nTo: {to_address}\nFrom: {from_address}\n\n{bodyText}".encode('utf-8') # メールの内容をUTF-8でエンコード

        if "gmail.com" in to_address:
            port = 465
            with smtplib.SMTP_SSL('smtp.gmail.com', port) as smtp_server:
                smtp_server.login(from_address, from_pw)
                smtp_server.sendmail(from_address, to_address, message)
            print("gmail送信処理完了。")
            logging.info('正常にgmail送信完了')
        elif "outlook.com" in to_address:
            port = 587
            with smtplib.SMTP('smtp.office365.com', port) as smtp_server:
                smtp_server.starttls()  # Enable security FIRST
                smtp_server.login(from_address, from_pw)  # Then log in
                smtp_server.sendmail(from_address, to_address, message)
            print("Outlookメール送信処理完了。")
            logging.info('Outlookメール送信完了')
            
    except Exception as e:
        logging.error('メール送信処理でエラーが発生しました。{}'.format(e))
    else:
        logging.info('メール送信処理を中止します。')


def name_converter(source_fold, filename):
    del_word = " - Made with Clipchamp"
    if del_word in filename:
        # remove the word from the filename(just variable)
        new_file_name = filename.replace(del_word, "") 
        
        # change the name of the file in the folder
        old_path = os.path.join(source_fold, filename)
        new_path = os.path.join(source_fold, new_file_name)
        os.rename(old_path, new_path)
        
        print(f"Clipchamp detected in filename {filename}")
        print(f'Renamed: {old_path} -> {new_path}')
        return new_file_name # name after changing
    else:
        print("Clipchamp not detected in filename")
        return filename # return original name


# check sheet with certain name exists in a google spreadsheet
def check_sheet_exists(sheet_name, workbook):
    sheets = workbook.worksheets()
    
    chk = False
    for sheet in sheets:
        if sheet.title == sheet_name:
            chk = True
            
        if len(sheets) > 1:
            if "Sheet" in sheet.title:
                workbook.del_worksheet(sheet)
                chk = False
    if chk != True:
        workbook.add_worksheet(title=sheet_name, rows=1000, cols=10)
    return sheet_name


# all file in each game videos folders and write to gsheet
def listup_all_files(folder_path, sheet):
    print(sheet.title)
    sheet.clear()
    
    # list up all files in the folder
    print(f"親フォルダパス: {folder_path}")
    folder_names = os.listdir(folder_path)
    data_list = []
    
    for folder_name in folder_names: # each path in folder_path
        time.sleep(2)
        
        file_path = os.path.join(folder_path, folder_name)
        
        # check if file exists in each folder
        try:
            if not len(os.listdir(file_path)) == 0:
                files = os.listdir(file_path)
                
                if files != None:
                    for file in files:
                        print(f"{file_path}\\{file}")
                        data_list.append([file_path, file])
            else:
                print(f"No files in the folder: {file_path}")
        except Exception as e:
            print(f"Error: {e}")
            print(f"It's not a folder, but a file {file_path}")
    # write down into a google spreadsheet
    df = pd.DataFrame(data_list, columns=["folder", "file_name"])
    
    logging.info(f"Write down into a sheet in Google spreadsheet: {df}")
    sheet.update([df.columns.values.tolist()] + df.values.tolist())
    logging.info(f"List up all files in {folder_path} and write to {sheet.title} in google spreadsheet.")
