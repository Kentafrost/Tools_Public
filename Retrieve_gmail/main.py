import re
import pandas as pd
import matplotlib.pyplot as plt
import boto3
import common
import logging


def main():
    gmail_service = common.authorize_gmail()
    messages = []

    page_token = None
    while len(messages) < 200:  # Set your limit
        response = gmail_service.users().messages().list(
            userId='me',
            maxResults=200,  # Adjust batch size
            pageToken=page_token
        ).execute()

        if 'messages' in response:
            messages.extend(response.get('messages', []))

        page_token = response.get('nextPageToken')
        if not page_token:
            break  # No more pages

    rakuten_cost_list = []
    rakuten_date_list = []
    coursera_list = []
    coursera_enrolled_list = []

    for message in messages:
        msg = gmail_service.users().messages().get(
            userId='me', 
            id=message['id']
        ).execute()
        
        # msg = gmail_service.users().messages().list(userId='me', id=message['id']).execute()
        if "楽天カード" and "カードのご利用があったことを迅速にお知らせ" in msg['snippet']:
            
            # cost
            pattern = r'本人 *([\d,]+) 円'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                cost = int(matches[0].replace(',', ''))
                rakuten_cost_list.append(cost)
            
            # 利用金額～本人までの数字を抽出
            # 例: "ご利用金額 2025/05/04 本人"から、2025/05/04を抽出
            
            # date
            pattern = r'ご利用金額 (\d{4}/\d{2}/\d{2}) 本人'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                # to date type
                # date = datetime.strptime(matches[0], '%Y/%m/%d').date()
                rakuten_date_list.append(matches[0])
            
        # Coursera certificate
        elif "Coursera" and "Congratulations! Your Certificate is Ready"in msg['snippet']:
            # print(msg['snippet'])
            coursera_list.append(msg['snippet'])
            
        elif "Coursera" and "Welcome to the Specialization!" or "We’re thrilled you enrolled" in msg['snippet']:
            # your enrolled in ～ on Coursera
            # print(msg['snippet'])
            pattern = r'you enrolled in (.+?) on Coursera'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                coursera_enrolled_list.append(matches[0])
    
    msg = (
        f"Total cost I spent in Rakuten Card: {sum(rakuten_cost_list)}円" +
        f"\nNumbers of course completed in Coursera: {len(coursera_list)}" +
        f"\nNumbers of course enrolled in Coursera: {len(coursera_enrolled_list)}"
    )
    logging.info(msg)
    
    with open("G:\\My Drive\\Tool\\Retrieve_gmail\\rakuten_cost.txt", "w") as f:
        f.write(msg)
    
    rakuten_cost_list = sorted(rakuten_cost_list)
    rakuten_date_list = sorted(rakuten_date_list)
    
    gc = common.authorize_gsheet()
    workbook = gc.open("gmail_summary")
    sheet_name = "Summary"
    sheet = workbook.worksheet(sheet_name)
    sheet.clear()

    if rakuten_cost_list:
        data = [['Rakuten Card Date', 'Rakuten Card Cost']]  # Header row
        data.extend(zip(rakuten_date_list, rakuten_cost_list))  # Add rows of data
        
        sheet.update(values=data, range_name='A1')

        # Print the data for confirmation
        for date, cost in zip(rakuten_date_list, rakuten_cost_list):
            print(date, cost)
        
        df = pd.DataFrame({
            'Rakuten Card Date': rakuten_date_list,
            'Rakuten Card Cost': rakuten_cost_list
        })
        
        # 2025/01/01 ⇒ 2025/01(all sum up)
        df['Month'] = pd.to_datetime(df['Rakuten Card Date']).dt.strftime('%Y/%m')
        
        pivot_table = pd.pivot_table(
            df, 
            index=['Month'],
            values=['Rakuten Card Cost'],
            aggfunc='sum'
            )
        print(pivot_table)
        
        # make a bar plot
        pivot_table.plot(kind='bar', title='Rakuten Card Cost by Month')
        plt.xlabel('Month')
        plt.ylabel('Cost')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()
        plt.savefig('G:\\My Drive\\Tool\\Retrieve_gmail\\rakuten_cost_by_month.png', 
                    dpi=300, bbox_inches='tight')
        pivot_png_path = "G:\\My Drive\\Tool\\Retrieve_gmail\\rakuten_cost_by_month.png"

        sheet_name = "Pivot_tbl"    
       
        # insert pivot table to google sheet
        sheet = workbook.worksheet(sheet_name)
        sheet.clear()
        
        pivot_reset = pivot_table.reset_index()
        pivot_data = [pivot_reset.columns.tolist()] + pivot_reset.values.tolist()

        sheet.update(values=pivot_data, range_name='A1')

        ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
        common.send_mail(ssm_client, pivot_png_path)
        
    # get most recent data from rakuten_date_list
    if rakuten_date_list:
        most_recent_date = rakuten_date_list[1]
        min_recent_date = min(rakuten_date_list)
        print(f"Most recent date of Rakuten Card usage: {most_recent_date}")
        print(f"Most old date of Rakuten Card usage: {min_recent_date}")
    else:
        print("No Rakuten Card usage data found.")
        
if __name__ == '__main__':
    main()