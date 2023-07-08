import pyodbc
import pandas as pd
# import tkinter as tk
# from tkinter import filedialog
import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

server = '14.192.27.190'
database = 'FMCOMP021'
username = 'sa'
password = 'Race123'


# Establish the connection
conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password)
cursor = conn.cursor()
df = pd.read_sql_query("SELECT * FROM V25M2DMSOutLog where status='E'", conn)
df.drop(['Request','Response'],axis=1,inplace=True)
df.to_csv('FlairTemp.csv')
grouped = df.groupby('ModuleID').size().reset_index(name='Count')
grouped.to_csv("Count.csv")
cursor.execute("UPDATE V25M2DMSOutLog SET Status = 'M' WHERE Status = 'E'")
cursor.commit()

sender = 'kevalflair@green-mystery-282317.iam.gserviceaccount.com'
recipient = 'edp@flairpens.com'
subject = 'DMS Error response'
# Create the email message
message = MIMEMultipart()
message['From'] = sender
message['To'] = recipient
message['Subject'] = subject

if(len(df)>0):
    body = "Please find attached the CSV files. FlairTemp file contains all the rows with the status 'E' and Count File contains the count of each ModuleID"
    # Attach the body text
    message.attach(MIMEText(body, 'plain'))
    # Attach the CSV file
    csv_file_paths = ['FlairTemp.csv','Count.csv']
    for csv_file_path in csv_file_paths:
        print(csv_file_path)
        attachment = open(csv_file_path, 'rb')
        csv_part = MIMEBase('application', 'octet-stream')
        csv_part.set_payload((attachment).read())
        encoders.encode_base64(csv_part)
        csv_part.add_header('Content-Disposition', "attachment; filename= %s" % csv_file_path)
        message.attach(csv_part)
    print(1)
else:
    body = "There are no new rows with the status 'E'"
    # Attach the body text
    message.attach(MIMEText(body, 'plain'))
    print(0)

# Connect to the SMTP server
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'n9023722@gmail.com'
smtp_password = 'xkirxpkfirfkabeh'

with smtplib.SMTP(smtp_server, smtp_port) as server:
    # Start TLS encryption
    server.starttls()
    
    # Authenticate (if required)
    server.login(smtp_username, smtp_password)
    
    # Send the email
    server.sendmail(sender, recipient, message.as_string())

print('Email sent successfully.')