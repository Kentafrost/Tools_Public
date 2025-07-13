import os
import logging
import gspread
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

gmail_cred = "G:\\My Drive\\Tool\\Retrieve_gmail\\OAuth_credentials.json"
gsheet_cred = "G:\\My Drive\\Tool\\credentials.json"

def authorize_gmail():
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
        return gmail_service
    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        return None    


def authorize_gsheet():
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
        return gspread_client
    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        return None


def send_mail(client, attachment_path):
    logging.info('メール送信処理を開始します。')
    try:
        from_address = client.get_parameter(Name='my_main_gmail_address', 
                                            WithDecryption=True)['Parameter']['Value']
        from_pw = client.get_parameter(Name='my_main_gmail_password', 
                                       WithDecryption=True)['Parameter']['Value']  
        to_address = from_address

        # Create MIME message
        msg = MIMEMultipart()
        msg['Subject'] = "Rakuten cards fee summary"
        msg['From'] = from_address
        msg['To'] = to_address

        body = "This is the report of cost used in Rakuten card"
        msg.attach(MIMEText(body, 'plain'))

        # Add attachment
        with open(attachment_path, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-Disposition', 'attachment', 
                           filename=os.path.basename(attachment_path))
            msg.attach(img)

    except Exception as e:
        logging.error('メールの内容をSSMから取得できませんでした。{}'.format(e))
    
    try:
        # message = f"Subject: {subject}\nTo: {to_address}\nFrom: {from_address}\n\n{bodyText}".encode('utf-8') # メールの内容をUTF-8でエンコード

        if "gmail.com" in to_address:
            port = 465
            with smtplib.SMTP_SSL('smtp.gmail.com', port) as smtp_server:
                smtp_server.login(from_address, from_pw)
                smtp_server.send_message(msg)
            print("gmail送信処理完了。")
            logging.info('正常にgmail送信完了')

        elif "outlook.com" in to_address:
            port = 587
            with smtplib.SMTP('smtp.office365.com', port) as smtp_server:
                smtp_server.starttls()  # Enable security FIRST
                smtp_server.login(from_address, from_pw)
                smtp_server.sendmail(from_address, to_address, msg)
            print("Outlookメール送信処理完了。")
            logging.info('Outlookメール送信完了')
            
    except Exception as e:
        logging.error('メール送信処理でエラーが発生しました。{}'.format(e))
    else:
        logging.info('メール送信処理を中止します。')
        return None
    