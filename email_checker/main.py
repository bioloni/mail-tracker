import pandas as pd
import os
import json
import email
import imaplib
import re

#Define helper functions

#Email extractor
def find_email(text):
    email = re.findall(r'[\w\.-]+@[\w\.-]+',str(text))
    return ",".join(email)

#Column selector function
def col_sel(x):
    if x['From'] == credentials["username"]: return x['To']
    else: return x['From']





credentials = {}

# Try to get credentials from credentials.json
try:
    with open("credentials.json") as f:
        credentials = json.load(f)
except FileNotFoundError:
    credentials["username"] = input("Outlook username:")
    credentials["password"] = input("Outlook password:")

# Connect to Outlook using the credentials

#Debugging lines#####
#print("Username:", credentials["username"])
#print("Password:", credentials["password"])
###############

M = imaplib.IMAP4_SSL("imap-mail.outlook.com")
M.login(credentials["username"], credentials["password"])

# Fetch sent emails
M.select("Sent")
status, sent_emails = M.search(None, "ALL")
sent_email_ids = sent_emails[0].split()

# Create dataframe with sent emails
sent_emails_list = []
for email_id in sent_email_ids:
    status, email_data = M.fetch(email_id, "(RFC822)")
    email_message = email.message_from_bytes(email_data[0][1])
    sent_emails_list.append([email_message['Date'], credentials["username"], email_message['To'], email_message['Subject'], email_message.get_payload(), "Sent"])
sent_df = pd.DataFrame(sent_emails_list, columns=["Date", "From", "To","Subject", "Body", "Status"])

# Fetch inbox emails
M.select("Inbox")
status, inbox_emails = M.search(None, "ALL")
inbox_email_ids = inbox_emails[0].split()

# Append inbox emails to dataframe
inbox_emails_list = []
for email_id in inbox_email_ids:
    status, email_data = M.fetch(email_id, "(RFC822)")
    email_message = email.message_from_bytes(email_data[0][1])
    inbox_emails_list.append([email_message['Date'], email_message['From'], email_message['To'], email_message['Subject'], email_message.get_payload(), "Received"])
inbox_df = pd.DataFrame(inbox_emails_list, columns=["Date", "From", "To", "Subject","Body", "Status"])

# Append the dataframe to the main dataframe
emails_df = sent_df.append(inbox_df)

#Extract email from column
emails_df['From']=emails_df['From'].apply(lambda x: find_email(x))
emails_df['To']=emails_df['To'].apply(lambda x: find_email(x))

#Generate client column
emails_df['Client']=emails_df.apply(lambda x: col_sel(x),axis=1)

#Download the raw dataframe
emails_df.to_csv("emails_raw.csv", index=False)

# Group by user email and select the latest interaction
df = emails_df.sort_values("Date", ascending=False).groupby("Client").first().reset_index()

# Download the dataframe to a local csv file
df.to_csv("latest_emails.csv", index=False)
