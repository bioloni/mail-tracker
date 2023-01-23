import pandas as pd
import os
import json
import email
import imaplib
import re
from datetime import timedelta, datetime
from tkinter import *

def on_submit():
    credentials["username"] = username_entry.get()
    credentials["password"] = password_entry.get()
    M = imaplib.IMAP4_SSL("imap-mail.outlook.com")
    M.login(credentials["username"], credentials["password"])
    M.select("Sent")
    status, sent_emails = M.search(None, "ALL")
    sent_email_ids = sent_emails[0].split()
    sent_emails_list = []
    for email_id in sent_email_ids:
        status, email_data = M.fetch(email_id, "(RFC822)")
        email_message = email.message_from_bytes(email_data[0][1])
        sent_emails_list.append([email_message['Date'], credentials["username"], email_message['To'], email_message['Subject'], email_message.get_payload(), "Sent"])
    sent_df = pd.DataFrame(sent_emails_list, columns=["Date", "From", "To","Subject", "Body", "Status"])
    M.select("Inbox")
    status, inbox_emails = M.search(None, "ALL")
    inbox_email_ids = inbox_emails[0].split()
    inbox_emails_list = []
    for email_id in inbox_email_ids:
        status, email_data = M.fetch(email_id, "(RFC822)")
        email_message = email.message_from_bytes(email_data[0][1])
        inbox_emails_list.append([email_message['Date'], email_message['From'], email_message['To'], email_message['Subject'], email_message.get_payload(), "Received"])
    inbox_df = pd.DataFrame(inbox_emails_list, columns=["Date", "From", "To", "Subject","Body","Status"])
    latest_emails_csv = pd.concat([sent_df, inbox_df])
    latest_emails_csv['To/From'] = latest_emails_csv.apply(col_sel, axis=1)
    latest_emails_csv['Action'] = latest_emails_csv.apply(action_gen, axis=1)
    latest_emails_csv.to_csv("latest_emails.csv", index=False)
    table = Text(root, height=20, width=60)
    table.pack()
    table.insert(END, latest_emails_csv.to_string())

    root = Tk()
    root.title("Outlook Email Manager")
    username_label = Label(root, text="Username:")
    username_label.pack()
    username_entry = Entry(root)
    username_entry.pack()
    password_label = Label(root, text="Password:")
    password_label.pack()
    password_entry = Entry(root, show="*")
    password_entry.pack()
    submit_button = Button(root, text="Submit", command=on_submit)
    submit_button.pack()
    root.mainloop()
