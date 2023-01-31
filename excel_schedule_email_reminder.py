import smtplib, ssl
from email.message import EmailMessage
import pandas
from datetime import datetime
import numpy as np
import toml

def is_negative(num):
    return num < 0

def nearest_index(items, pivot):
    time_diff = np.abs([date - pandas.Timestamp(pivot) for date in items])

    check = np.array([date - pandas.Timestamp(pivot) for date in items])

    return time_diff.argmin(0)

def contruct_reminder_message(data_frame: pandas.DataFrame, column_exclusion_list):

    now = datetime.now()
    given_date = now.strftime("%Y-%m-%d %H:%M:%S")
    dates = data_frame["DATE"]
    t_index = nearest_index(dates, given_date)


    # If the nearest saturday has already passed then set the index to the next one
    delta = data_frame.loc[t_index, "DATE"] - datetime.strptime(given_date, "%Y-%m-%d %H:%M:%S")
    if is_negative(delta.days):
        t_index = t_index + 1

    
    message = ""
    for i, col in enumerate(data_frame.columns):
        value = data_frame.loc[t_index, col]
        if col in column_exclusion_list:
            continue
        if str(value) == "nan":
            continue
        if i != 0:
            message = message + "\n" + col + ": " + str(value)
        else:
            message = message + col + ": " +  str(value) + "\n"

    print(message)
    return message


with open("config.toml", "r") as f:
    config = toml.load(f)


smtp_server = config["variables"]["smtp_server"]
port = 587  
password = config["variables"]["password"]

sender_email = config["variables"]["sender_email"]
receiver_email = config["variables"]["receiver_email"]
subject = config["variables"]["subject"]

excel_file = config["variables"]["excel_file"]
column_exclusion_list = config["variables"]["exclusion_list"]



# Create a secure SSL context
context = ssl.create_default_context()


# Try to log in to server and send email
try:
    data_frame = pandas.read_excel(excel_file)

    message = contruct_reminder_message(data_frame, column_exclusion_list)

    server = smtplib.SMTP(smtp_server,port)
    server.ehlo() # Can be omitted
    server.starttls(context=context) # Secure the connection
    server.ehlo() # Can be omitted
    server.login(sender_email, password)

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content(message)

    server.send_message(msg)

except Exception as e:
    # Print any error messages to stdout
    print("error: ", e)
finally:
    server.quit() 