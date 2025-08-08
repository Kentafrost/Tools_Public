import boto3
import requests
import os, sys
from bs4 import BeautifulSoup
import time
import csv
import logging

# parent directory
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)

import common

# page count variables
page_num_count = 1
maximum_page_num = 10

# URL and search word, change these as needed
web_url = ""
exception_words = []

website_data_list = []
common.import_log("Websites-List")

# skip titles with exception words
if exception_words == []:
    print(f"No exception words written, skipping check to except")

def list_websites(search_word):
    for page_num in range(1, maximum_page_num + 1):
        try:
            print(f"Processing page {page_num}...")
            
            # Add headers to avoid being blocked
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(f"{web_url}?word={search_word}&c=&page={page_num}", headers=headers)
            
            if response.status_code == 200:
                # Parse HTML content
                soup = BeautifulSoup(response.content, 'html.parser')
                
                class_name = "itemTitle"
                item_titles = soup.find_all(class_=class_name)
                print(f"Found {len(item_titles)} titles on page {page_num}")

                if not item_titles:
                    print(f"No links with title found on page {page_num}")
                    break  # No more content

                for item_title in item_titles:
                    href = None
                    title = None
                    page_num_str = f"Page num: {page_num}"

                    link_tag = item_title.find('a')
                    if link_tag:
                        href = link_tag.get('href')
                        title = link_tag.get_text(strip=True) or item_title.get_text(strip=True)
                        print(title)

                    else:
                        # Method 3: Get text from the element, no href available
                        title = item_title.get_text(strip=True)
                        print(f"No <a> tag found, only text: {title}")
                        
                    if not exception_words == []:
                        for word in exception_words:
                            if word in title:
                                print(f"Skipping title due to exception words: {title}")
                            continue
                    
                    # Convert relative URLs to absolute URLs
                    if href and not href.startswith('http'):
                        if href.startswith('/'):
                            href = web_url.rstrip('/') + href
                        else:
                            href = web_url.rstrip('/') + '/' + href

                    if title and title.strip():  # Only add non-empty titles
                        website_data = {
                            'title': title.strip(),
                            'url': href if href else 'No URL found',
                            'page_num_str': page_num_str
                        }
                        website_data_list.append(website_data)
                        print(f"Found: {website_data}")
                
                # Add delay to be respectful to the server
                time.sleep(1)
                
            else:
                print(f"Failed to retrieve page {page_num}. Status code: {response.status_code}")
                if response.status_code == 404:
                    break  # No more pages
                    
        except Exception as e:
            print(f"Error processing page {page_num}: {e}")
            continue

    print(f"Total titles found: {len(website_data_list)}")
    logging.info(f"Total titles found: {len(website_data_list)}")
    
    if not website_data_list:
        print("No titles found. Please check the search word or the website URL.")
        logging.info("No titles found. Please check the search word or the website URL.")
        return []
    
    return website_data_list

def csv_writer(website_data_list, search_word):

    with open(os.path.join(results_dir, f'websites_{search_word}.csv'), 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Write header
        writer.writerow(['Title', 'URL', 'Page_Number'])
        
        # Write data
        for website_data in website_data_list:
            writer.writerow([website_data['title'], website_data['url'], website_data['page_num_str']])

    print("Websites data written to the CSV file. Please check the results directory.")
    print(f"File location: {os.path.abspath(f'./results/websites_{search_word}.csv')}")
    logging.info(f"CSV file created successfully for search word: {search_word}")

if __name__ == "__main__":
    # Create results directory if it doesn't exist
    current_dir = os.path.dirname(os.path.abspath(__file__))
    results_dir = os.path.join(current_dir, 'results')
    os.makedirs(results_dir, exist_ok=True)

    for search_word in []:
        web_data_list = list_websites(search_word)
        csv_writer(web_data_list, search_word)
    logging.info(f"Websites list completed for search words: {', '.join([''])}")