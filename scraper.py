# File: scraper.py
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import os

url = 'https://davjehanabad.in'
folder_path = r"BD"
os.makedirs(folder_path, exist_ok=True)
headers = {'User-Agent': 'Mozilla/5.0'}
HASH_FILE = 'last_hash.txt'

def get_last_hash():
    """Reads the last saved hash from a file."""
    try:
        with open(HASH_FILE, 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        return None

def save_new_hash(new_hash):
    """Saves the new hash to a file."""
    with open(HASH_FILE, 'w') as f:
        f.write(str(new_hash))

def get_webpage_content():
    """Fetches the content from the website."""
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        print("Request successful")
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        return None

def process_content(html_content):
    """Extracts the relevant div from the HTML."""
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        div = soup.find('div', id='pnlBirthdayDescipBox2')
        return str(div) if div else None
    except Exception as e:
        print(f"Content processing error: {e}")
        return None

# --- THIS IS THE MODIFIED FUNCTION ---
def save_content(content):
    """Saves the content to a new HTML file named with the current date (YYYY-MM-DD.html)."""
    # Generate the filename based on the current date
    today_date_str = datetime.now().strftime('%Y-%m-%d')
    filename = f"{today_date_str}.html"
    file_path = os.path.join(folder_path, filename)
    
    # We can still use a more precise timestamp for the title inside the HTML file
    full_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    html_template = f"""<!DOCTYPE html>
<html>
<head><title>Birthday Info ({today_date_str})</title></head>
<body>
<!-- Last updated: {full_timestamp} -->
{content}
</body>
</html>"""
    
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(html_template)
    print(f"New content found. Saved/Updated file: {filename}")

def main():
    """Main function to run the scrape-and-save process once."""
    print("Starting scraper...")
    last_content_hash = get_last_hash()
    
    html_content = get_webpage_content()
    if html_content:
        processed_content = process_content(html_content)
        if processed_content:
            current_hash = str(hash(processed_content))
            
            if current_hash != last_content_hash:
                save_content(processed_content)
                save_new_hash(current_hash)
            else:
                print("No changes found on the website.")
        else:
            print("Could not find the target content div on the page.")
    
    print("Scraper finished.")

if __name__ == "__main__":
    main()
