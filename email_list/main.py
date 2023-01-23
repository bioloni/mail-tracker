import imaplib
import pandas as pd
import email
import re
import time

# Connect to the Outlook account
imap_host = 'imap-mail.outlook.com'
imap_user = 'RoboButlerBot@outlook.com'
imap_pass = ':;?=)!/"&qweorkfjsa54894'

mail = imaplib.IMAP4_SSL(imap_host)
mail.login(imap_user, imap_pass)

# File name for the CSV file
file_name = "emails.csv"

# Continuously listen for new emails
while True:
    # Check for new emails
    mail.select('inbox')
    status, data = mail.search(None, 'UNSEEN')
    mail_ids = data[0]
    id_list = mail_ids.split()

    # Create a list to store email data
    email_data = []
    # Iterate through new emails
    for email_id in id_list:
        # Extract email data
        status, data = mail.fetch(email_id, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        sender = msg['From']
        cc = msg['CC']
        subject = msg['Subject']
        body = msg.get_payload(decode=True).decode('utf-8','ignore').strip()
        # Append email data to list
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

    # Check for emails containing "re" in the cc field
    re_emails = df[df["CC"].str.contains("re", case=False, na=False)]
    for email in re_emails.itertuples():
        reply_to = email.Sender
        reply_subject = "Thank you for your reply!"
        reply_body = "Thank you for your reply!"
        msg = MIMEMultipart()
        msg['From'] = imap_user
        msg['To'] = reply_to
        msg['Subject'] = reply_subject
        msg.attach(MIMEText(reply_body, 'plain'))
        mail.sendmail(imap_user, reply_to, msg.as_string())
        print(f"Sent reply to {reply_to}")

    # Add a delay before checking for new emails again
    time.sleep(300)