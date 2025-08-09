import streamlit as st
from email.header import decode_header
import pytesseract
import fitz
from PIL import Image
import io
import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from utils import * 

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Dhruv\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

st.set_page_config(page_title="Email PDF OCR Service", layout="wide")

st.title("📧 Email PDF OCR Service")
st.write("Fetch PDF attachments from your mailbox and extract text using OCR.")

email_user = st.text_input("Email Address")

label_col, icon_col = st.columns([6, 0.2])
with label_col:
    st.markdown("**Email App Password**")
with icon_col:
    open_help = st.button("ℹ️", help="How to get Gmail App Password", key="help_btn", use_container_width=True)

email_pass = st.text_input(
    "Email App Password",
    type="password",
    label_visibility="collapsed"
)

provider = st.selectbox("Email Provider", list(IMAP_SERVERS.keys()))

duration_hours = st.number_input("Duration (hours)", min_value=1, max_value=24, value=1)

if open_help and provider == "Gmail":
    with st.popover("📜 Gmail App Password Instructions"):
        st.markdown("""
        Gmail does **not** allow your normal password for third-party apps.  
        You need to generate an **App Password**:

        1. Go to [Google Account Security](https://myaccount.google.com/security).
        2. Turn on **2-Step Verification** (if not already enabled).
        3. In **App passwords**, choose:
            - **App:** Mail  
            - **Device:** Windows Computer (or any option)
        4. Click **Generate**.
        5. Copy the 16-character password shown.
        6. Paste that here instead of your normal Gmail password.

        ⚠️ This is only for Gmail. Outlook, Yahoo, etc. have different steps.
        """)


if st.button("Fetch & OCR PDFs"):
    if not email_user or not email_pass:
        st.error("Please provide email credentials.")
    else:
        try:
            since_time = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=duration_hours)

            with st.spinner("Connecting to email server..."):
                mail_conn = connect_to_email(email_user, email_pass, provider)

            with st.spinner("Fetching recent PDF attachments..."):
                pdf_files = fetch_recent_pdfs(mail_conn, since_time)

            if not pdf_files:
                st.warning("No recent PDF attachments found.")
            else:
                with st.spinner("Running OCR on PDFs concurrently..."):
                    with ThreadPoolExecutor(max_workers=10) as executor:
                        future_to_filename = {executor.submit(pdf_to_text, pdf_data): filename for filename, pdf_data in pdf_files}
                        for future in as_completed(future_to_filename):
                            filename = future_to_filename[future]
                            try:
                                text = future.result()
                                st.subheader(f"📄 {filename}")
                                st.text_area(f"Extracted Text - {filename}", text, height=300)
                            except Exception as e:
                                st.error(f"Error processing {filename}: {str(e)}")

            mail_conn.logout()
        except Exception as e:
            st.error(f"Error: {str(e)}")