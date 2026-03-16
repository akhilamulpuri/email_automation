import win32com.client
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os
import re

# -----------------------------
# CONFIGURATION
# -----------------------------

ATTACHMENT_FOLDER = r"C:\Users\Akhila.Mulpuri\OneDrive - GEP\Documents\invoice_Processing"
EXCEL_FILE = r"C:\Users\Akhila.Mulpuri\OneDrive - GEP\Documents\invoice process.xlsx"

POPPLER_PATH = r"C:\Users\Akhila.Mulpuri\OneDrive - GEP\Documents\Release-25.12.0-0\poppler-25.12.0\Library\bin"
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Akhila.Mulpuri\OneDrive - GEP\Documents\Tesseract-OCR\tesseract.exe"

os.makedirs(ATTACHMENT_FOLDER, exist_ok=True)

columns = [
"Invoice Number",
"Vendor Name",
"Invoice Date",
"Due Date",
"Payment Terms",
"Amount",
"Currency",
"Sender Email",
"Email Received Time",
"Assigned To",
"Mailbox",
"Attachment Name"
]

# -----------------------------
# LOAD EXISTING EXCEL
# -----------------------------

if os.path.exists(EXCEL_FILE):
    existing_df = pd.read_excel(EXCEL_FILE)
    processed_files = set(existing_df["Attachment Name"].astype(str))
else:
    existing_df = pd.DataFrame(columns=columns)
    processed_files = set()

# -----------------------------
# FUNCTION: EXTRACT DATA
# -----------------------------

def extract_invoice_data(text, filename, sender, received_time):

    data = {col: "" for col in columns}

    invoice = re.search(r'Invoice\s*(No|Number)?[:\s]*([A-Z0-9\-\/]+)', text, re.I)
    if invoice:
        data["Invoice Number"] = invoice.group(2)

    date = re.search(r'(\d{2}/\d{2}/\d{4})', text)
    if date:
        data["Invoice Date"] = date.group(1)

    amount = re.search(r'([\d,]+\.\d+)', text)
    if amount:
        data["Amount"] = amount.group(1)

    currency = re.search(r'(INR|USD|Rs\.?|₹)', text, re.I)
    if currency:
        data["Currency"] = currency.group(1)

    lines = text.split("\n")
    if len(lines) > 0:
        data["Vendor Name"] = lines[0].strip()

    data["Attachment Name"] = filename
    data["Sender Email"] = sender
    data["Email Received Time"] = str(received_time)
    data["Mailbox"] = "Inbox"

    return data


# -----------------------------
# FUNCTION: WRITE TO EXCEL
# -----------------------------

def write_to_excel(row):

    global existing_df

    existing_df = pd.concat([existing_df, pd.DataFrame([row])], ignore_index=True)

    existing_df.to_excel(EXCEL_FILE, index=False)


# -----------------------------
# CONNECT TO OUTLOOK
# -----------------------------

print("Connecting to Outlook...")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

# -----------------------------
# PROCESS EMAILS
# -----------------------------

for msg in messages:

    try:
        subject = str(msg.Subject)

        if "invoice" in subject.lower():

            sender = msg.SenderEmailAddress
            received_time = msg.ReceivedTime

            print("\nInvoice Email Found:", subject)

            if msg.Attachments.Count > 0:

                for attachment in msg.Attachments:

                    if attachment.FileName.lower().endswith(".pdf"):

                        # Skip if already processed
                        if attachment.FileName in processed_files:
                            print("Already processed:", attachment.FileName)
                            continue

                        file_path = os.path.join(ATTACHMENT_FOLDER, attachment.FileName)

                        attachment.SaveAsFile(file_path)

                        print("Saved:", file_path)

                        # -----------------------------
                        # OCR PROCESS
                        # -----------------------------

                        images = convert_from_path(file_path, poppler_path=POPPLER_PATH)

                        text = ""

                        for img in images:
                            text += pytesseract.image_to_string(img)

                        # -----------------------------
                        # DATA EXTRACTION
                        # -----------------------------

                        data = extract_invoice_data(
                            text,
                            attachment.FileName,
                            sender,
                            received_time
                        )

                        print("Extracted Data:", data)

                        write_to_excel(data)

                        processed_files.add(attachment.FileName)

            else:
                print("No attachments found")

    except Exception as e:
        print("Email error:", e)

print("\nProcess completed. Excel updated.")
