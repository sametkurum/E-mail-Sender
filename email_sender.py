# BASIC E-MAIL SENDER, W\Gmail
# -*- coding: utf-8 -*-

import smtplib
from email.header import Header

# text your sender e-mail address. For example: "mygmail@gmail.com"
sender = "" 
# text receivers' e-mail addresses. For example: "hisgmail@gmail.com", "hergmail@gmail.com"
receivers = []
# text recievers' names. For example: "Rafa Silva", "Ciro Immobile"
receiver_name = []
#You should get Google APP password 
password = ""
#Text your mail's subject
subject = ""
#Text your mail's body
body = ""

# Part of connecting to SMTP server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender, password)
print("Logged in...")

for i in range(len(receivers)):
    # Code template and message as UTF-8
    subject_encoded = Header(subject, 'utf-8').encode()
    #Text sender name. exp: "From: Samet Kurum"
    message = f"""From: Name <{sender}> 
To: {receiver_name[i]} <{receivers[i]}>
Subject: {subject_encoded}\n
{body}
"""

    try:
        # Send the message coded as UTF-8 (for Turkish characters)
        server.sendmail(sender, receivers[i], message.encode('utf-8'))
        print(f"Email sent to {receivers[i]}!")
    except smtplib.SMTPRecipientsRefused as e:
        print(f"Failed to send email to {receivers[i]}: {e}")
    except Exception as e:
        print(f"An error occurred while sending to {receivers[i]}: {e}")

# Cancel the server connection
server.quit()
