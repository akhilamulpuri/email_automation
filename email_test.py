import streamlit as st
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

def load_excel():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        processed_files = set(df["Attachment Name"].astype(str))
    else:
        df = pd.DataFrame(columns=columns)
        processed_files = set()
    return df, processed_files

# -----------------------------
# EXTRACT DATA FROM OCR
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
# WRITE TO EXCEL
# -----------------------------

def write_to_excel(df, row):
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    return df

# -----------------------------
# MAIN PROCESS FUNCTION
# -----------------------------

def process_invoices():

    df, processed_files = load_excel()

    st.write("Connecting to Outlook...")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    processed_count = 0

    for msg in messages:

        try:
            subject = str(msg.Subject)

            if "invoice" in subject.lower():

                sender = msg.SenderEmailAddress
                received_time = msg.ReceivedTime

                st.write("Invoice Email Found:", subject)

                if msg.Attachments.Count > 0:

                    for attachment in msg.Attachments:

                        if attachment.FileName.lower().endswith(".pdf"):

                            if attachment.FileName in processed_files:
                                st.write("Already processed:", attachment.FileName)
                                continue

                            file_path = os.path.join(ATTACHMENT_FOLDER, attachment.FileName)

                            attachment.SaveAsFile(file_path)

                            st.write("Saved:", attachment.FileName)

                            images = convert_from_path(file_path, poppler_path=POPPLER_PATH)

                            text = ""

                            for img in images:
                                text += pytesseract.image_to_string(img)

                            data = extract_invoice_data(
                                text,
                                attachment.FileName,
                                sender,
                                received_time
                            )

                            df = write_to_excel(df, data)

                            processed_files.add(attachment.FileName)

                            processed_count += 1

        except Exception as e:
            st.error(e)

    st.success(f"{processed_count} new invoices processed")

    return df


# -----------------------------
# STREAMLIT UI
# -----------------------------

st.title("Invoice Automation System")

st.write("Process Outlook invoice emails and extract invoice data.")

if st.button("Process Invoice Emails"):

    with st.spinner("Processing invoices..."):
        df = process_invoices()

    st.success("Processing complete!")

    st.subheader("Invoice Data")
    st.dataframe(df)
