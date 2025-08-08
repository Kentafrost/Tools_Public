import json
import csv
import os
import pandas as pd
from datetime import datetime
import logging
import boto3

def get_chrome_bookmarks_path():
    """Get the Chrome bookmarks file path based on OS"""
    import platform
    
    system = platform.system()
    if system == "Windows":
        # Default Chrome bookmarks path on Windows
        username = os.getenv('USERNAME')
        ssm = boto3.client('ssm')
        username = ssm.get_parameter(Name='My_Name', WithDecryption=True)['Parameter']['Value']
        
        chrome_path = f"C:\\Users\\{username}\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Bookmarks"
    elif system == "Darwin":  # macOS
        username = os.getenv('USER')
        chrome_path = f"/Users/{username}/Library/Application Support/Google/Chrome/Default/Bookmarks"
    else:  # Linux
        username = os.getenv('USER')
        chrome_path = f"/home/{username}/.config/google-chrome/Default/Bookmarks"
    
    return chrome_path

def extract_bookmarks(bookmarks_data, folder_name="", bookmarks_list=None):
    """Recursively extract bookmarks from Chrome bookmarks JSON structure"""
    if bookmarks_list is None:
        bookmarks_list = []
    
    if isinstance(bookmarks_data, dict):
        if bookmarks_data.get('type') == 'url':
            # This is a bookmark
            bookmark = {
                'name': bookmarks_data.get('name', ''),
                'url': bookmarks_data.get('url', ''),
                'folder': folder_name,
                'date_added': bookmarks_data.get('date_added', ''),
                'date_modified': bookmarks_data.get('date_modified', '')
            }
            bookmarks_list.append(bookmark)
        
        elif bookmarks_data.get('type') == 'folder':
            # This is a folder, recurse into its children
            current_folder = bookmarks_data.get('name', '')
            full_folder_path = f"{folder_name}/{current_folder}" if folder_name else current_folder
            
            children = bookmarks_data.get('children', [])
            for child in children:
                extract_bookmarks(child, full_folder_path, bookmarks_list)
        
        # Handle root level structures
        if 'children' in bookmarks_data and not bookmarks_data.get('type'):
            children = bookmarks_data.get('children', [])
            for child in children:
                extract_bookmarks(child, folder_name, bookmarks_list)
    
    elif isinstance(bookmarks_data, list):
        for item in bookmarks_data:
            extract_bookmarks(item, folder_name, bookmarks_list)
    
    return bookmarks_list

def convert_chrome_timestamp(timestamp):
    """Convert Chrome timestamp to readable format"""
    try:
        if timestamp:
            # Chrome timestamps are microseconds since 1601-01-01
            # Convert to seconds since 1970-01-01
            timestamp_seconds = (int(timestamp) - 11644473600000000) / 1000000
            return datetime.fromtimestamp(timestamp_seconds).strftime('%Y-%m-%d %H:%M:%S')
    except:
        return ''
    return ''

def export_bookmarks_to_csv(output_file='bookmarks_export.csv'):
    """Main function to export Chrome bookmarks to CSV"""
    
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        # Get Chrome bookmarks file path
        bookmarks_path = get_chrome_bookmarks_path()
        
        if not os.path.exists(bookmarks_path):
            logging.error(f"Chrome bookmarks file not found at: {bookmarks_path}")
            print(f"Chrome bookmarks file not found at: {bookmarks_path}")
            print("Please check if Chrome is installed and you have bookmarks saved.")
            return False
        
        # Read Chrome bookmarks JSON file
        with open(bookmarks_path, 'r', encoding='utf-8') as file:
            bookmarks_data = json.load(file)
        
        logging.info("Successfully loaded Chrome bookmarks file")
        
        # Extract bookmarks from different root folders
        all_bookmarks = []
        
        # Chrome typically has 'bookmark_bar', 'other', and 'synced' folders
        roots = bookmarks_data.get('roots', {})
        
        for root_name, root_data in roots.items():
            if root_data and isinstance(root_data, dict):
                root_folder_name = root_data.get('name', root_name)
                extracted = extract_bookmarks(root_data, root_folder_name)
                all_bookmarks.extend(extracted)
        
        if not all_bookmarks:
            logging.warning("No bookmarks found in the Chrome bookmarks file")
            print("No bookmarks found in the Chrome bookmarks file")
            return False
        
        # Convert timestamps to readable format
        for bookmark in all_bookmarks:
            bookmark['date_added_readable'] = convert_chrome_timestamp(bookmark['date_added'])
            bookmark['date_modified_readable'] = convert_chrome_timestamp(bookmark['date_modified'])
        
        # Create DataFrame
        df = pd.DataFrame(all_bookmarks)
        
        # Reorder columns for better readability
        column_order = ['name', 'url', 'folder', 'date_added_readable', 'date_modified_readable', 'date_added', 'date_modified']
        df = df.reindex(columns=column_order, fill_value='')
        
        # Export to CSV
        current_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(current_dir, output_file)
        
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        logging.info(f"Successfully exported {len(all_bookmarks)} bookmarks to: {output_path}")
        print(f"Successfully exported {len(all_bookmarks)} bookmarks to: {output_path}")
        
        # Display summary
        print(f"\nSummary:")
        print(f"Total bookmarks: {len(all_bookmarks)}")
        print(f"Folders found: {len(df['folder'].unique())}")
        print(f"Output file: {output_path}")
        
        return True
        
    except FileNotFoundError:
        logging.error(f"Chrome bookmarks file not found: {bookmarks_path}")
        print(f"Chrome bookmarks file not found: {bookmarks_path}")
        return False
    except json.JSONDecodeError as e:
        logging.error(f"Error parsing Chrome bookmarks JSON: {e}")
        print(f"Error parsing Chrome bookmarks JSON: {e}")
        return False
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        print(f"Unexpected error: {e}")
        return False

def list_bookmark_folders(bookmarks_path=None):
    """List all bookmark folders for reference"""
    try:
        if not bookmarks_path:
            bookmarks_path = get_chrome_bookmarks_path()
        
        with open(bookmarks_path, 'r', encoding='utf-8') as file:
            bookmarks_data = json.load(file)
        
        folders = set()
        all_bookmarks = []
        
        roots = bookmarks_data.get('roots', {})
        for root_name, root_data in roots.items():
            if root_data and isinstance(root_data, dict):
                root_folder_name = root_data.get('name', root_name)
                extracted = extract_bookmarks(root_data, root_folder_name)
                all_bookmarks.extend(extracted)
        
        for bookmark in all_bookmarks:
            if bookmark['folder']:
                folders.add(bookmark['folder'])
        
        print("Bookmark folders found:")
        for folder in sorted(folders):
            print(f"  - {folder}")
        
        return sorted(folders)
        
    except Exception as e:
        print(f"Error listing bookmark folders: {e}")
        return []

if __name__ == "__main__":
    print("Chrome Bookmarks Exporter")
    print("=" * 30)
    
    # Export bookmarks to CSV
    success = export_bookmarks_to_csv('chrome_bookmarks_export.csv')
    
    if success:
        print("\n" + "=" * 50)
        print("Bookmark folders structure:")
        list_bookmark_folders()
    else:
        print("Export failed. Please check the error messages above.")
