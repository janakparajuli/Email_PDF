# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 09:54:27 2020

@author: janak
Adapted from: https://stackoverflow.com/questions/58575615/how-do-i-attach-separate-pdfs-to-contact-list-email-addresses-using-python
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from string import Template
import pandas as pd

#e = pd.read_csv("Contacts.csv")
e = pd.read_excel(r"path/excel_file.xlsx")

server = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
server.ehlo()
# server.starttls()
server.login('yourgmail@gmail.com','yourpassword')

body = ("""
Dear ... , 

Write your message here
""")
subject = "your subject"
fromaddr='yourgmail@gmail.com'
#body = "Subject: {}\n\n{}".format(subject,msg)
#Emails,PDF
for index, row in e.iterrows():
    print('Sending email to: ')
    print (row["Email"]+' '+row["PDF"])
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    filename = row["PDF"]
    toaddr = row["Email"]
    attachment = open(row["PDF"], "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    #server.sendmail('sender_email',emails,body)

print("Emails sent successfully")

server.quit()