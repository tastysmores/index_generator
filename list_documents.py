import os
import pandas as pd
from pathlib import Path
import email
from email import policy
from email.parser import BytesParser
from extract_msg import Message  # For extracting data from .msg files
from pytz import timezone
from datetime import datetime
import re
from urllib.parse import unquote

def decode_imceaex(address: str) -> str:
    if not address.upper().startswith("IMCEAEX-"):
        return address  # normal SMTP address

    # Strip IMCEAEX- prefix and domain
    local_part = address.split("@", 1)[0][8:]

    # Convert underscores to slashes
    local_part = local_part.replace("_", "/")

    # Decode +XX hex escapes
    local_part = re.sub(
        r"\+([0-9A-Fa-f]{2})",
        lambda m: bytes.fromhex(m.group(1)).decode("latin-1"),
        local_part,
    )

    # The actual mailbox is usually after the final dash
    if "-" in local_part:
        local_part = local_part.split("-")[-1]

    return local_part

def parse_email_address(address):

    #print(address)

    if address is not None:

        if "<" in address:

            contact_name = address.split("<", 1)[0]
            email_address = address.split("<", 1)[1].rstrip(">")

            email_address = decode_imceaex(email_address)

            contact_name = contact_name.strip()

        else:
            contact_name = address
            email_address = "(No Email)"


    else:
        contact_name = "(No Name)"
        email_address = "(No Email)"

    return contact_name, email_address

def parse_dt(s):
    for fmt in (
        "%Y-%m-%d %H:%M:%S.%f%z",
        "%Y-%m-%d %H:%M:%S%z",
        "%a, %d %b %Y %H:%M:%S %z",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unrecognized datetime format: {s}")

def extract_email_fields_from_eml(eml_file):
    try:
        with open(eml_file, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # Extract email fields
        received_date = msg.get('Date') or "(No Date)"
        subject = msg.get('Subject') or "(No Subject)"
        sender_name, sender_email = parse_email_address(msg.get('From')) or "(No Sender)"
        recipient_name, recipient_email = parse_email_address(msg.get('To')) or "(No Recipients)"
        cc_name, cc_email = parse_email_address(msg.get('Cc')) or "(No CC)"
        bcc_name, bcc_email = parse_email_address(msg.get('Bcc')) or "(No BCC)"
        
        return received_date, subject, sender_name, sender_email, recipient_name, recipient_email, cc_name, cc_email, bcc_name, bcc_email
    except Exception as e:
        print(f"Error extracting data from {eml_file}: {e}")
        return "(Error extracting date)", "(Error extracting subject)", "(Error extracting sender)", "(Error extracting sender)","(Error extracting recipients)","(Error extracting recipients)", "(Error extracting CC)","(Error extracting CC)", "(Error extracting BCC)", "(Error extracting BCC)"

def extract_email_fields_from_msg(msg_file):
    try:
        msg = Message(msg_file)

        # Extract email fields
        received_date = msg.date or "(No Date)"
        
        subject = msg.subject or "(No Subject)"
        subject = subject.strip()

        sender_name, sender_email = parse_email_address(msg.sender) or "(No Sender)"
        recipient_name, recipient_email = parse_email_address(msg.to) or "(No Recipients)"
        cc_name, cc_email = parse_email_address(msg.cc) or "(No CC)"
        bcc_name, bcc_email = parse_email_address(msg.bcc) or "(No BCC)"
        
        return received_date, subject, sender_name, sender_email, recipient_name, recipient_email, cc_name, cc_email, bcc_name, bcc_email
    except Exception as e:
        print(f"Error extracting data from {msg_file}: {e}")
        return "(Error extracting date)", "(Error extracting subject)", "(Error extracting sender)", "(Error extracting sender)","(Error extracting recipients)","(Error extracting recipients)", "(Error extracting CC)","(Error extracting CC)", "(Error extracting BCC)", "(Error extracting BCC)"

def export_folder_contents_to_excel(folder_path, index_output_path, extract_email_details, ebrief):
    folder_path = Path(folder_path).resolve()

    # Verify the folder exists
    if not folder_path.is_dir():
        print(f"Error: '{folder_path}' is not a valid folder.")
        return
    
    # Prepare a list to hold file data
    file_data = []
    
    def strip_ebrief_prefix(name):
        #ebrief_number = name.split(" ")[0]
        #print(ebrief_number)
        
        return name.split(" ", 1)[1], name.split(" ", 1)[0]

    # Function to recursively iterate through the folder and its subfolders
    def iterate_folder(folder):
        for item in folder.iterdir():
            if item.is_file():  # Only include files, not directories
                # Get the relative path of the file
                relative_path = item.relative_to(folder_path).as_posix()
                
                
                ebrief_number = ""
                
                if ebrief and "LR" in item.name:
                    item_name, ebrief_number = strip_ebrief_prefix(item.name)
                    item_name = item.name
                else:
                    item_name = item.name

                # Extract email fields if it's an .eml or .msg file
                if extract_email_details and item.suffix.lower() == '.eml':
                    received_date, subject, sender_name, sender_email, recipient_name, recipient_email, cc_name, cc_email, bcc_name, bcc_email = extract_email_fields_from_eml(item)
                    
                    if "error" not in str(received_date).lower():
                        #received_date_datetime = datetime.strptime(str(received_date), "%a, %d %b %Y %H:%M:%S %z")
                        received_date_datetime = parse_dt(str(received_date))
                        received_date_datetime = received_date_datetime.astimezone(timezone('Australia/Melbourne'))
                        received_date_string = received_date_datetime.strftime("%a, %d %b %Y %H:%M:%S")

                        suggested_file_name = received_date_datetime.strftime("%Y-%m-%d") + " - Email from " + sender_name + " to " + recipient_name + " '" + subject + "'" + ".eml"
                    #print(received_date_datetime)
                    else:
                        suggested_file_name = item.name
                        received_date_string = "(Error extracting date)"
                    
                    #suggested_name = "Email from " + sender_name + " to " + recipient_name + " on " + str(received_date)
                elif extract_email_details and item.suffix.lower() == '.msg':
                    received_date, subject, sender_name, sender_email, recipient_name, recipient_email, cc_name, cc_email, bcc_name, bcc_email = extract_email_fields_from_msg(item)
                    
                    if "error" not in str(received_date).lower():
                        #received_date_datetime = datetime.strptime(str(received_date), "%Y-%m-%d %H:%M:%S%z")
                        received_date_datetime = parse_dt(str(received_date))
                        received_date_datetime = received_date_datetime.astimezone(timezone('Australia/Melbourne'))
                        received_date_string = received_date_datetime.strftime("%a, %d %b %Y %H:%M:%S")
                        suggested_file_name = received_date_datetime.strftime("%Y-%m-%d") + " - Email from " + sender_name + " to " + recipient_name + " '" + subject + "'" + ".msg"
                    else:
                        suggested_file_name = item.name
                        received_date_string = "(Error extracting date)"
                else:
                    received_date_string = "(Not an email)"
                    subject = "(Not an email)"
                    sender_name = "(Not an email)"
                    sender_email = "(Not an email)"
                    recipient_name = "(Not an email)"
                    recipient_email = "(Not an email)"
                    cc_name = "(Not an email)"
                    cc_email = "(Not an email)"
                    bcc_name = "(Not an email)"
                    bcc_email = "(Not an email)"
                    suggested_file_name = item_name
                
                # Append file data (relative path and email metadata)
                file_data.append([relative_path, received_date_string, subject, sender_name, sender_email, recipient_name, recipient_email, cc_name, cc_email, bcc_name, bcc_email, suggested_file_name])
            
            elif item.is_dir():  # Recurse into subfolders
                iterate_folder(item)

    # Start the iteration from the original folder
    iterate_folder(folder_path)

    # Convert the file data into a DataFrame
    df = pd.DataFrame(file_data, columns=['Relative Path', 'Received Date', 'Subject', 'From Name', 'From Email', 'To Name', 'To Email', 'CC Name', 'CC Email', 'BCC Name', 'BCC Email', 'New File Name'])

    # Convert the 'Received Date' to a readable datetime format if it's available


    #df['Received Date'] = pd.to_datetime(df['Received Date'], errors='coerce', utc=True)

    # Remove timezone information by converting to naive datetime (if applicable)
    #df['Received Date'] = df['Received Date'].dt.tz_localize(utc=True)
    #df['Received Date'] = df['Received Date'].dt.tz_convert('Australia/Melbourne')
    #df['Received Date'] = df['Received Date'].dt.tz_localize(None)

    # Create the output Excel file
    #output_file = folder_path / f"{folder_path.name}_folder_contents.xlsx"
    df.to_excel(index_output_path, index=False)

    print(f"Excel file created successfully: {index_output_path}")

# Prompt for the folder path and run the script
#folder_path = input("Enter the path of the folder: ").strip()
#export_folder_contents_to_excel(folder_path)
