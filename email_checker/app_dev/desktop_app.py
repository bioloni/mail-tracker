import pandas as pd
import os
import json
import email
import imaplib
import re
from datetime import timedelta, datetime
from tkinter import *
from tkinter import ttk
from functools import partial
import keyring

#Duplicate client remover
def duprem(x):
    a=x['Client'].split(",")
    print(a)
    b=list(set(a))
    print(b)
    c=", ".join(b)
    print(c)
    return c

#Email extractor
def find_email(text):
    email = re.findall(r'[\w\.-]+@[\w\.-]+',str(text))
    return ",".join(email)

#Column selector function
def col_sel(x):
    if x['From'] == credentials["username"]: return x['To']
    else: return x['From']

def action_gen(x):
    today = datetime.today()
    grace_period=3
    date_diff=(today - x['Date']).days
    if (x['Status'] == "Received") and (date_diff >= grace_period): return "URGENT Client is expecting an answer"
    elif (x['Status'] == "Received") and (date_diff < grace_period): return "Client is expecting an answer"
    elif (x['Status'] == "Sent") and (date_diff >= grace_period): return "Follow up on the email you sent"
    elif (x['Status'] == "Sent") and (date_diff < grace_period): return "Give the client more time to answer"
    else: return "Error"

def on_submit():
    global credentials
    credentials={}
    credentials["username"] = username_entry.get()
    credentials["password"] = password_entry.get()
    keyring.set_password("Outlook", credentials["username"], credentials["password"])
    M = imaplib.IMAP4_SSL("imap-mail.outlook.com")
    M.login(credentials["username"], keyring.get_password("Outlook", credentials["username"]))
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
    # Append the dataframe to the main dataframe
    emails_df = pd.concat([sent_df,inbox_df],ignore_index=True)
    

    #Post processing
    #Format date column
    emails_df['Date'] = emails_df['Date'].apply(lambda x: datetime.strptime(x, "%a, %d %b %Y %H:%M:%S %z"))
    emails_df['Date'] = emails_df['Date'].apply(lambda x: x.replace(tzinfo=None))
    #Extract email from column
    emails_df['From']=emails_df['From'].apply(lambda x: find_email(x))
    emails_df['To']=emails_df['To'].apply(lambda x: find_email(x))
    #Generate client column
    emails_df['Client']=emails_df.apply(lambda x: col_sel(x),axis=1)
    #Remove duplicate clients
    emails_df['Client']=emails_df.apply(lambda x: duprem(x),axis=1)
    print(emails_df)
    #Generate action column
    emails_df['Action']=emails_df.apply(lambda x: action_gen(x),axis=1)


    #Download the raw dataframe
    emails_df.to_csv("emails_raw.csv", index=False)
    # Group by user email and select the latest interaction
    emails_df = emails_df.sort_values("Date", ascending=False).groupby("Client").first().reset_index()
    # Download the dataframe to a local csv file
    emails_df.to_csv("latest_emails.csv", index=False)

    root = Tk()
    root.title("Outlook Email Manager")
    root.geometry("800x600") # increase the size of the window
    tree = ttk.Treeview(root)
    tree["columns"]=("one","two","three","four","five","six")
    tree.column("#0", width=270, minwidth=270, stretch=NO)
    tree.column("one", width=150, minwidth=150, stretch=NO)
    tree.column("two", width=400, minwidth=200)
    tree.column("three", width=80, minwidth=50, stretch=NO)
    tree.column("four", width=150, minwidth=150, stretch=NO)
    tree.column("five", width=200, minwidth=200)
    tree.heading("#0",text="Date",anchor=W)
    tree.heading("one", text="Client",anchor=W)
    tree.heading("two", text="Subject",anchor=W)
    tree.heading("three", text="Status",anchor=W)
    tree.heading("four", text="Action",anchor=W)


    for index,row in emails_df.iterrows():
        tree.insert("",index,text=row['Date'],values=(row['Client'],row['Subject'],row['Status'],action_gen(row)))

    login_window.withdraw()
    tree.pack(side=LEFT, fill=BOTH)
    root.mainloop()



# GUI code
login_window = Tk()
login_window.geometry("300x200")
login_window.title("Outlook Email Manager")

#username label and text entry box
username_label = Label(login_window, text="Username:")
username_label.pack()
username = StringVar()
username_entry = Entry(login_window, textvariable=username)
username_entry.pack()

#password label and password entry box
password_label = Label(login_window, text="Password:")
password_label.pack()
password = StringVar()

password_entry = Entry(login_window, textvariable=password, show='*')
password_entry.pack()



submit_button = Button(login_window, text="Submit", command=on_submit)
submit_button.pack()
login_window.mainloop()


