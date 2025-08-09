import streamlit as st
import imaplib
import email
from email.header import decode_header
import pytesseract
import fitz 
from PIL import Image
import io

IMAP_SERVERS = {
    "Gmail": "imap.gmail.com",
    "Outlook / Hotmail": "outlook.office365.com",
    "Yahoo": "imap.mail.yahoo.com",
    "iCloud": "imap.mail.me.com",
    "Zoho": "imap.zoho.com"
}

def log(message):
    """Helper to log debug info to Streamlit UI."""
    st.write(f"🛠 {message}")

def connect_to_email(username, password, provider):
    imap_server = IMAP_SERVERS[provider]
    log(f"Connecting to {imap_server}...")
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(username, password)
    log("✅ Logged in successfully.")
    return mail

def fetch_recent_pdfs(mail, since_time):
    """Fetch recent PDF attachments from inbox with detailed logging."""
    log("Selecting inbox...")
    mail.select("inbox")

    log("Fetching all recent message IDs...")
    status, messages = mail.search(None, "ALL")
    if status != "OK":
        log("❌ Failed to search mailbox.")
        return []

    message_ids = messages[0].split()
    log(f"Found {len(message_ids)} messages in inbox.")
    
    recent_ids = message_ids[-5:]
    pdf_files = []

    for num in reversed(recent_ids):
        status, data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            log(f"⚠ Skipping message {num} (fetch failed).")
            continue
        
        msg = email.message_from_bytes(data[0][1])
        msg_date = msg.get("Date", "Unknown")
        log(f"Checking email dated: {msg_date}")

        try:
            msg_datetime = email.utils.parsedate_to_datetime(msg_date)
        except:
            log("⚠ Could not parse email date, skipping.")
            continue

        if msg_datetime < since_time:
            log("⏩ Email is older than given duration, skipping.")
            continue

        # Process parts for attachments
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", "")).lower()
            content_type = part.get_content_type()

            if "attachment" in content_disposition or content_type == "application/pdf":
                filename = part.get_filename()
                if filename:
                    decoded_filename = decode_header(filename)[0][0]
                    if isinstance(decoded_filename, bytes):
                        decoded_filename = decoded_filename.decode()
                    if decoded_filename.lower().endswith(".pdf"):
                        log(f"📄 Found PDF: {decoded_filename}")
                        pdf_data = part.get_payload(decode=True)
                        pdf_files.append((decoded_filename, pdf_data))
            else:
                log("No PDF attachment in this part.")

    return pdf_files

def pdf_to_text(pdf_bytes):
    """Convert PDF to text using OCR."""
    text_output = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
        for page_num in range(len(pdf)):
            pix = pdf[page_num].get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img, lang="eng")
            text_output.append(text)
    return "\n".join(text_output)