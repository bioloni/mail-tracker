import tkinter as tk
from tkinter import filedialog
import pandas as pd
import datetime

# Create the window
root = tk.Tk()
root.title("Email App")

# Create the new window
master_window = tk.Toplevel(root)
master_window.title("Master Dataframe")
master_window.geometry("800x600") # set the window size
master_window.withdraw()

# Create a DataFrame widget to display the master dataframe
df_widget = ttk.Treeview(master_window, columns=master.columns)
df_widget.grid(row=0, column=0, sticky="nsew")

# Create the input fields and upload buttons
inbox_label = tk.Label(root, text="Inbox emails:")
inbox_label.grid(row=0, column=0)
inbox_input = tk.Entry(root)
inbox_input.grid(row=0, column=1)
inbox_button = tk.Button(root, text="Upload", command=lambda: select_file(inbox_input))
inbox_button.grid(row=0, column=2)

sent_label = tk.Label(root, text="Sent emails:")
sent_label.grid(row=1, column=0)
sent_input = tk.Entry(root)
sent_input.grid(row=1, column=1)
sent_button = tk.Button(root, text="Upload", command=lambda: select_file(sent_input))
sent_button.grid(row=1, column=2)

# Function to open the file browser and save the selected file path
def select_file(input_field):
    filepath = filedialog.askopenfilename()
    input_field.delete(0, tk.END)
    input_field.insert(0, filepath)


def action_gen(x):
    grace_period=3
    date_diff=(today - x['Date']).days
    if (x['Status'] == "Received") and (date_diff >= grace_period): return "URGENT Client is expecting an answer"
    elif (x['Status'] == "Received") and (date_diff < grace_period): return "Client is expecting an answer"
    elif (x['Status'] == "Sent") and (date_diff >= grace_period): return "Follow up on the email you sent"
    elif (x['Status'] == "Sent") and (date_diff < grace_period): return "Give the client more time to answer"
    else: return "Error"

# Function to import and process the data
def process_data():
    # Import the data from the csv files
    inbox = pd.read_csv(inbox_input.get())
    sent = pd.read_csv(sent_input.get())

    # Add the "Email type" column
    inbox["Status"] = "Received"
    sent["Status"] = "Sent"

        # Rename the "Sent" and "Received" columns to "Date"
    inbox = inbox.rename(columns={"Received": "Date"})
    sent = sent.rename(columns={"Sent": "Date"})

    # Add the "To" and "From" columns
    inbox["To"] = "Me"
    sent["From"] = "Me"

    # Add the "Client" columns
    inbox["Client"] = inbox["From"]
    sent["Client"] = sent["To"]

    # Remove the "Size" and "Categories" columns
    inbox = inbox.drop(columns=["Size", "Categories"])
    sent = sent.drop(columns=["Size", "Categories"])

    # Concatenate the dataframes and sort by date
    master = pd.concat([inbox, sent], ignore_index=True)
    master["Date"] = pd.to_datetime(master["Date"])
    master = master.sort_values(by=["Date"])

    # Group by "Client" and keep only the latest row
    master = master.drop_duplicates(subset='Client', keep='last')

    # Add the status column
    grace_period = 3
    #Generate action column
    master['Action']=master.apply(lambda x: action_gen(x),axis=1)

    # Populate the dataframe widget with the master dataframe
    for col in master.columns:
        df_widget.heading(col, text=col)
        df_widget.column(col, width=tkfont.Font().measure(col.title()))

    for row in master.itertuples():
        df_widget.insert("", "end", values=row[1:])

    #Open the master window
    master_window.deiconify()
    root.withdraw()

# Create the submit button
submit_button = tk.Button(root, text="Submit", command=process_data)
submit_button.grid(row=2, column=1)
