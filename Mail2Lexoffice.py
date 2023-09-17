import imaplib
import email
import os
import requests
import pdfplumber
import re
import json
import logging
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
import threading

root = tk.Tk()
root.iconbitmap(r'./icon.ico')
root.title("Mail to Lexoffice")

console = tk.Text(root, wrap=tk.WORD, state="disabled")
console.pack(pady=15, padx=10)

# Mail account data
email_address_entry_Input = tk.StringVar()
password_entry_Input = tk.StringVar()
imap_server_entry_Input = tk.StringVar()

# API-Token
access_token_entry_Input = tk.StringVar()

def update_console(text):
    console.configure(state="normal")
    console.insert(tk.END, text + "\n")
    console.configure(state="disabled")
    console.see(tk.END)

def log_info(message):
    logging.info(message)
    update_console(message)

def process_emails():

    imap_server = imap_server_entry_Input.get()
    email_address = email_address_entry_Input.get()
    password = password_entry_Input.get()
    access_token = access_token_entry_Input.get()

    # Connect to IMAP-Server
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, password)
    except Exception as e:
        update_console("Fehler beim Verbinden zum IMAP-Server: " + str(e)) 
        exit(1)

    # Config file for processed mails - #ToDo: Add to DB + track mail logic
    config_file = "processed_emails.json"

    # Check if config file exists - create otherwise
    if not os.path.exists(config_file):
        with open(config_file, "w") as f:
            json.dump([], f)

    # Load Mail IDs from configfile
    with open(config_file, "r") as f:
        processed_emails = json.load(f)

    # Initialize Logger
    log_file = "email_processing_log.txt"
    logging.basicConfig(filename=log_file, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    try:
        mail.select("inbox")

        # Set search criteria for the emails
        search_criteria = 'ALL'
        result, data = mail.search(None, search_criteria)

        # Create list for found attachments
        attachments = []

        # Invoice recognition keywords
        keywords = ["Betrag", "Gesamtbetrag", "Netto", "Brutto", "Rechnung", "Invoice", "Beleg", "Zahlung"]

        total_attachments = len(data[0].split())
        processed_attachments = 0

        # Start logic for getting attachments
        log_info("Search emails for invoices...")
        for num in data[0].split():
            email_id = num.decode("utf-8")
            if email_id in processed_emails:
                continue

            processed_attachments += 1
            progress = processed_attachments / total_attachments * 100

            result, msg_data = mail.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "application/pdf":
                    filename = part.get_filename()
                    if filename:
                        attachments.append((filename, part.get_payload(decode=True)))

            log_info(f"Processed: {processed_attachments}/{total_attachments} ({progress:.2f}%)")

        mail.logout()
        log_info("\nSearching the found attachments for invoices...")

        # Upload attachments to the lexoffice
        upload_url = "https://api.lexoffice.io/v1/files"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        }
        
        processed_attachments = 0

        if attachments:
            for attachment in attachments:
                processed_attachments += 1
                progress = processed_attachments / len(attachments) * 100

                filename, data = attachment
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                temp_file.write(data)
                temp_file.close()

                with pdfplumber.open(temp_file.name) as pdf:
                    pdf_text = ""
                    for page in pdf.pages:
                        pdf_text += page.extract_text()

                    # Check for keywords
                    found_keywords = [keyword for keyword in keywords if re.search(rf"\b{keyword}\b", pdf_text, re.IGNORECASE)]
                    if found_keywords:
                        files = {"file": (filename, open(temp_file.name, "rb"), "multipart/form-data")}
                        payload = {"type": "voucher"}

                        response = requests.post(upload_url, headers=headers, files=files, data=payload)

                        if response.status_code == 202:
                            log_info(f"Invoice {filename} successfully uploaded.                  ")
                        else:
                            log_info(f"Error uploading {filename}. Status code: {response.status_code}         ")
                    else:
                        log_info(f"{filename} is not an invoice. Will be ignored.                  ")

                log_info(f"Processed: {processed_attachments}/{len(attachments)} ({progress:.2f}%)")

                # Mark the email as processed
                processed_emails.append(email_id)
                with open(config_file, "w") as f:
                    json.dump(processed_emails, f)

                temp_file.close()

        else:
            log_info("No PDF attachments found.")
    except Exception as e:
        log_info("Error processing the e-mails: " + str(e))
    finally:
        #TODO - end connection to IMAP server
        log_info("Successfully Finished.")
        
def main():

    imap_label = ttk.Label(root, text="IMAP Server:")
    imap_server_entry = ttk.Entry(root, textvariable=imap_server_entry_Input)
    email_label = ttk.Label(root, text="E-Mail Adresse:")
    email_address_entry = ttk.Entry(root, textvariable=email_address_entry_Input)
    password_label = ttk.Label(root, text="Passwort:")
    password_entry = ttk.Entry(root, textvariable=password_entry_Input, show="*")
    token_label = ttk.Label(root, text="Access Token:")
    access_token_entry = ttk.Entry(root, textvariable=access_token_entry_Input, show="*")
    
    imap_label.pack(pady=5)
    imap_server_entry.pack(pady=5)
    email_label.pack(pady=5)
    email_address_entry.pack(pady=5)
    password_label.pack(pady=5)
    password_entry.pack(pady=5)
    token_label.pack(pady=5)
    access_token_entry.pack(pady=5)

    process_button = ttk.Button(root, text="E-Mails verarbeiten", command=start_processing_thread)
    process_button.pack(pady=10)

    # Add separator
    separator = ttk.Separator(root)
    separator.pack(pady=10)


    footer_label = ttk.Label(root, text="Made by BiOz", anchor="center")
    footer_label.pack(pady=10)

    root.mainloop()

def start_processing_thread():
    # Start processing the emails in a separate thread
    processing_thread = threading.Thread(target=process_emails)
    processing_thread.start()

if __name__ == "__main__":
    main()
