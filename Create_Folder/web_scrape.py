from tracemalloc import Frame
import requests
from bs4 import BeautifulSoup
from common_tool import google_authorize
import pandas as pd
import boto3

def web_scrape(url):

    try:
        response = requests.get(url)
    except requests.exceptions.RequestException as e:
        print(f"Error: Unable to access the website. {e}")

    if response.status_code != 200:
        print(f"Error: Unable to access the website. Status code: {response.status_code}")
    else:
        soup = BeautifulSoup(response.text, 'html.parser')

        table = soup.find('table', {'id': 'sortabletable1'})  
        name = table.find_all('a')
        data = []
        
        # retrieve only text between <a> tags from table
        for ele in name:
            text = ele.text
            if text != "編集" and text != "":
                data.append(text)
                
    return data

gc = google_authorize() 
ssm_client = boto3.client('ssm', region_name='ap-southeast-2')

chara_wb = gc.open("chara_name_list")
url_wb = gc.open("url_tbl_id_list")
url_sheet = url_wb.worksheet("URL_list")

title7 = ssm_client.get_parameter(Name='Title7')["Parameter"]["Value"]
title7_folder_path = ssm_client.get_parameter(Name='Title7_folder_path')["Parameter"]["Value"]

sheet_list = [title7]
urls = url_sheet.get_all_values()

# 取得したいウェブサイトのURL
url_df = pd.DataFrame(urls)
print(url_df)

url_list = url_df[1:]
url_list = url_list[1]
print(url_list)

data_list = []

for url in url_list:
    data = web_scrape(url)
    data_list.append(data)

# Remove unwanted items from all lists in data_list
unwanted_items = ["追加", "イベント報酬"]
for data in data_list:
    for unwanted_item in unwanted_items:
        while unwanted_item in data:
            data.remove(unwanted_item)
            print(f"Removed unwanted item: {unwanted_item}")


for required_sht_name in sheet_list:
    worksheet = chara_wb.worksheet(required_sht_name)
    worksheet.clear()  # Clear the worksheet before writing new data
    
    count = 2
    worksheet.update("A1", [["Name"]])
    worksheet.update("B1", [["Folder_Path"]])

    # several lists in a list
    formatted_data = [[item] for sublist in data_list for item in sublist]  
    existing_values = worksheet.col_values(1)  # Get all values from Column 
    
    # Filter out duplicates
    if (formatted_data) not in existing_values:
        worksheet.append_rows(formatted_data)
        print(f"New data added successfully! {formatted_data}")
    else:
        print("No new data to add (all values already exist).")

    list_fold_path = [[title7_folder_path]] * len(formatted_data)
    worksheet.update(f"B2:B{len(formatted_data) + 1}", list_fold_path)

print("All webscraping operations and written to Google Spread Sheets successfully!")