from posixpath import dirname
import gspread
import pickle
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import os, logging, time

def authorize_gmail():
    
    gmail_cred = "G:\\My Drive\\IT_Learning\\Git\\tools\\Retrieve_gmail\\OAuth_credentials.json"
    
    try:
        scope = [
            "https://mail.google.com/",
            "https://www.googleapis.com/auth/gmail.modify",
            "https://www.googleapis.com/auth/gmail.readonly",
        ]
       
        creds = None
        token_file = "G:\\My Drive\\Tool\\Retrieve_gmail\\token.pickle"

        # Load existing credentials if available
        if os.path.exists(token_file):
            with open(token_file, "rb") as token:
                creds = pickle.load(token)

        # If no valid credentials, authenticate and save them
        if not creds or not creds.valid:
            # The file token.json stores the user's access and refresh tokens, and is        
            flow = InstalledAppFlow.from_client_secrets_file(gmail_cred, scope)
            creds = flow.run_local_server(port=0)
            
            # Save credentials for future use
            with open(token_file, "wb") as token:
                pickle.dump(creds, token)
        
        gmail_service = build('gmail', 'v1', credentials=creds)
        logging.info("Gmail service authorized successfully.")
        
        return gmail_service
    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        print(f"Error: {e}")
        os._exit(1)


def authorize_gsheet():
    
    gsheet_cred = "G:\\My Drive\\IT_Learning\\Git\\tools\\credentials.json"
    
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_file(
            gsheet_cred, scopes=scope
        )
        
        # Authorize the gspread client
        gspread_client = gspread.authorize(credentials)
        logging.info("Google Sheets service authorized successfully.")

        return gspread_client

    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        print(f"Error: {e}")
        os._exit(1)
    
def import_log(script_name):

    current_time = time.strftime("%Y%m%d_%H%M%S")
    this_dir = os.path.dirname(os.path.abspath(__file__))

    log_dir = f"{this_dir}\\log\\{script_name}\\"
    os.makedirs(log_dir, exist_ok=True)
    log_file = f"{log_dir}\\{current_time}_task.log"

    # Remove all handlers associated with the root logger object.
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='shift_jis'
    )
    