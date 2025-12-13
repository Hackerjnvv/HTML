import os
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.utils import get_column_letter

# Month abbreviations to month numbers
MONTHS = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
}

def clean_name(name):
    """Cleans and normalizes names by removing unnecessary spaces."""
    return ' '.join(name.split())

def extract_data_from_html(html_content):
    """Extracts data from HTML content."""
    soup = BeautifulSoup(html_content, 'lxml')
    data = []

    for card in soup.select('.card'):
        date = student_name = father_name = mother_name = class_info = section = ''

        try:
            date = card.select_one('.date').text.strip() if card.select_one('.date') else ''
            student_name = card.select_one('.card-title').text.strip() if card.select_one('.card-title') else ''

            card_text = card.select_one('.card-text')
            if card_text:
                lines = [line.strip() for line in card_text.stripped_strings]

                if len(lines) >= 1:  # Father's name
                    father_name = clean_name(lines[0].split('/')[0])
                if len(lines) >= 2:  # Mother's name
                    mother_name = clean_name(lines[1])
                if len(lines) >= 3:  # Class info
                    class_info = lines[2].split(':')[-1].split('/')[0].strip()
                if len(lines) >= 4:  # Section
                    section = lines[3].split(':')[-1].strip()

            data.append([date, student_name, father_name, mother_name, class_info, section])

        except Exception as e:
            print(f"Error parsing card: {e}")

    return data

def parse_day_month(date_str):
    """Parses a date string in 'DD,MMM' format."""
    try:
        day, month_abbr = date_str.split(',')
        day = int(day.strip())
        month = MONTHS.get(month_abbr.strip(), 0)
        return (month, day)
    except ValueError:
        return (0, 0)

def save_to_excel(data, file_path):
    """Saves the extracted data to an Excel file."""
    try:
        # Create or load the workbook
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            # Check if headers exist
            if sheet.max_row == 0:
                sheet.append(["Date", "Student Name", "Father's Name", "Mother's Name", "Class", "Section"])
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = "Birthday Data"
            sheet.append(["Date", "Student Name", "Father's Name", "Mother's Name", "Class", "Section"])
        
        # Add new data (skip duplicates)
        existing_entries = set()
        for row in sheet.iter_rows(min_row=2, values_only=True):
            existing_entries.add(tuple(row))
        
        new_entries_added = 0
        for entry in data:
            if tuple(entry) not in existing_entries:
                sheet.append(entry)
                new_entries_added += 1
        
        # Auto-adjust column widths
        for column in sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            sheet.column_dimensions[column_letter].width = adjusted_width
        
        # Save the workbook
        wb.save(file_path)
        print(f"Successfully saved {new_entries_added} new entries to {file_path}")
        return True
    
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return False

def process_html_files(directory):
    """Processes HTML files and saves data to Excel."""
    all_data = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.html'):
                file_path = os.path.join(root, file)
                print(f"Processing: {file_path}")
                with open(file_path, 'r', encoding='utf-8') as f:
                    all_data.extend(extract_data_from_html(f.read()))

    # Remove duplicates and sort by date
    unique_sorted_data = sorted(set(tuple(row) for row in all_data), key=lambda x: parse_day_month(x[0]))
    
    # Convert to list of lists (from tuples)
    processed_data = [list(item) for item in unique_sorted_data]
    
    # Save to Excel file
    excel_path = "Birthday Data Master.xlsx"
    save_to_excel(processed_data, excel_path)

# Configuration
HTML_DIRECTORY = 'BD'  # Directory containing HTML files

# Run the process
process_html_files(HTML_DIRECTORY)
