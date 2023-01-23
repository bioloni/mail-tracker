import imaplib
import pandas as pd
import email
import re

# Connect to the Outlook account
imap_host = 'imap-mail.outlook.com'
imap_user = 'RoboButlerBot@outlook.com'
imap_pass = ':;?=)!/"&qweorkfjsa54894'

mail = imaplib.IMAP4_SSL(imap_host)
mail.login(imap_user, imap_pass)

# File name for the CSV file
file_name = "emails.csv"

# Select the inbox
mail.select('inbox')

# Search for all emails in the inbox
status, data = mail.search(None, 'ALL')
mail_ids = data[0]
id_list = mail_ids.split()

# Create a list to store email data
email_data = []

# Iterate through new emails
for email_id in id_list:
    print(f"Fetching message with ID: {email_id}")
    # Extract email data
    status, data = mail.fetch(email_id, '(RFC822)')
    # check if the data is not empty
    if not data:
        print(f"Message with ID: {email_id} not found.")
        continue
    msg = email.message_from_bytes(data[0][1])
    sender = msg['From']
    sender = re.search(r'<(.*)>', sender).group(1)  # Extract the full email from the sender field
    cc = msg['CC']
    if not cc:
        cc = ""
    else:
        cc = re.search(r'<(.*)>', cc).group(1)  # Extract the full email from the cc field
    subject = msg['Subject']
    body = ""
    for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body += part.get_payload(decode=True).decode('utf-8').strip()
                except:
                    try:
                        body += part.get_payload(decode=True).decode('latin-1').strip()
                    except:
                        body += part.get_payload(decode=True).decode('iso-8859-1').strip()
    email_data.append([sender, cc, subject, body])




# Create a DataFrame from email data
df = pd.DataFrame(email_data, columns=["Sender", "CC", "Subject", "Body"])

# Open the existing CSV file
try:
    existing_df = pd.read_csv(file_name)
except FileNotFoundError:
    existing_df = pd.DataFrame(columns=["Sender", "CC", "Subject", "Body"])

# Check for existing emails in the CSV file
existing_emails = set(existing_df["Sender"])
new_emails = set(df["Sender"])
unique_emails = new_emails - existing_emails

# Filter the new DataFrame to include only new emails
df = df[df["Sender"].isin(unique_emails)]
if not df.empty:
    # Save the new email data to the CSV file
    with open(file_name, 'a') as f:
        df.to_csv(f, header=f.tell()==0)
    print(f"{len(df)} new email(s) appended to {file_name}")
else:
    print("No new emails found.")
