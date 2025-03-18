import os
import sys
import email
import extract_msg
from bs4 import BeautifulSoup
import re
from datetime import datetime, timezone
import dateutil.parser

def eml_to_text(input_folder, output_folder):
    """
    Converts .eml and .msg files from input_folder to a single output.txt file
    with all messages in chronological order.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List to store (date, text, filename) tuples for chronological sorting
    message_data = []

    for filename in os.listdir(input_folder):
        input_path = os.path.join(input_folder, filename)

        if not os.path.isfile(input_path):
            continue

        try:
            date = None
            text = None
            
            if filename.endswith(".eml"):
                with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
                    msg = email.message_from_file(f)
                date = extract_date_from_email(msg)
                # Improved extraction to avoid duplication
                text = extract_text_from_email(msg)
            elif filename.endswith(".msg"):
                msg = extract_msg.Message(input_path)
                date = msg.date
                text = msg.body
                msg.close()
            else:
                print(f"Skipping {filename}: Unsupported file type")
                continue

            # Try to extract date from filename if not found in message
            if date is None:
                date = extract_date_from_filename(filename)
            
            # Format the date for display
            formatted_date = format_date(date) if date else "Unknown Date"
            
            # Add date header to the text content
            full_text = f"Date: {formatted_date}\n\n{text}"

            # Store for chronological sorting - ensure date is normalized
            normalized_date = normalize_datetime(date) if date else None
            message_data.append((normalized_date, full_text, filename))
            
            print(f"Processed {filename}")

        except Exception as e:
            print(f"Error processing {filename}: {e}")

    # Create combined output file in chronological order
    create_combined_output(message_data, output_folder)

def normalize_datetime(dt):
    """
    Normalize datetime to ensure all are either naive or aware.
    Convert all to naive by removing timezone info for sorting purposes.
    """
    if dt is None:
        return None
    
    if dt.tzinfo is not None:
        # Convert to UTC then remove timezone info
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    
    return dt

def extract_date_from_email(msg):
    """Extract date from email message header."""
    date_str = msg.get('Date')
    if date_str:
        try:
            return dateutil.parser.parse(date_str)
        except:
            return None
    return None

def extract_date_from_filename(filename):
    """Extract date from filename using patterns like '2014 Aug 1' or similar."""
    # Pattern for "YYYY Mon DD" format
    pattern = r'(\d{4})\s+([A-Za-z]{3,})\s+(\d{1,2})'
    match = re.search(pattern, filename)
    if match:
        try:
            year, month, day = match.groups()
            # Convert month name to number
            date_str = f"{year} {month} {day}"
            return dateutil.parser.parse(date_str)
        except:
            pass
    return None

def format_date(date):
    """Format date object to string."""
    if date:
        return date.strftime("%Y-%m-%d %H:%M:%S")
    return "Unknown Date"

def create_combined_output(message_data, output_folder):
    """Create a single output file with all messages in chronological order."""
    output_path = os.path.join(output_folder, "output.txt")
    
    # Sort by date (handling None dates)
    sorted_data = sorted(
        message_data, 
        key=lambda x: (x[0] is None, x[0] or datetime.min)  # None dates go last
    )
    
    with open(output_path, 'w', encoding='utf-8') as outfile:
        for i, (date, text, filename) in enumerate(sorted_data):
            # Add separator between messages
            if i > 0:
                outfile.write("\n\n" + "="*80 + "\n\n")
            
            # Add source filename reference
            outfile.write(f"Source: {filename}\n")
            outfile.write(text)
    
    print(f"Created combined output file: {output_path}")

def extract_text_from_email(msg):
    """
    Extract text from an email message, prioritizing plain text and avoiding duplication.
    """
    # Try to get plain text first
    plain_text = ""
    html_text = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get('Content-Disposition') or '')
            
            if content_type == 'text/plain' and 'attachment' not in content_disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    plain_text += payload.decode('utf-8', errors='ignore')
            
            elif content_type == 'text/html' and 'attachment' not in content_disposition and not plain_text:
                # Only use HTML if plain text isn't available
                payload = part.get_payload(decode=True)
                if payload:
                    html = payload.decode('utf-8', errors='ignore')
                    html_text = BeautifulSoup(html, 'html.parser').get_text(separator='\n')
    else:
        # Non-multipart email
        content_type = msg.get_content_type()
        if content_type == 'text/plain':
            payload = msg.get_payload(decode=True)
            if payload:
                plain_text = payload.decode('utf-8', errors='ignore')
        elif content_type == 'text/html':
            payload = msg.get_payload(decode=True)
            if payload:
                html = payload.decode('utf-8', errors='ignore')
                html_text = BeautifulSoup(html, 'html.parser').get_text(separator='\n')
    
    # Return plain text if available, otherwise use HTML text
    if plain_text.strip():
        return plain_text
    else:
        return html_text

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python eml_to_text.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    eml_to_text(input_folder, output_folder)
