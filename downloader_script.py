import imaplib
import email
from email.header import decode_header
import os
import re
from dotenv import load_dotenv

# Load credentials from .env
load_dotenv()
EMAIL = os.getenv("EMAIL_ADDRESS")
APP_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")

# Constants
IMAP_SERVER = "imap.gmail.com"
SAVE_DIR = "attachments"
FILENAME_PATTERN = r"^report_\d{4}-\d{2}-\d{2}\.csv$"  # Customize as needed

# Ensure attachment folder exists
os.makedirs(SAVE_DIR, exist_ok=True)

# Connect to Gmail
print("Connecting to Gmail...")
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, APP_PASSWORD)
mail.select("inbox")

# Search all emails (you can filter by sender or subject later)
status, message_ids = mail.search(None, "ALL")
email_ids = message_ids[0].split()

print(f"Found {len(email_ids)} emails.")

# Process each email (starting from the latest)
for eid in reversed(email_ids):
    _, msg_data = mail.fetch(eid, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Parse the subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(errors="ignore")

            print(f"Checking email: {subject}")

            # Process attachments
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        # Decode filename
                        filename = decode_header(filename)[0][0]
                        if isinstance(filename, bytes):
                            filename = filename.decode(errors="ignore")

                        # Check if it matches the filename pattern
                        if re.match(FILENAME_PATTERN, filename):
                            filepath = os.path.join(SAVE_DIR, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"Saved attachment: {filename}")
                        else:
                            print(f"Skipped (no match): {filename}")
    break  # Test ONLY the most recent email. Remove this line to process all emails

mail.logout()
