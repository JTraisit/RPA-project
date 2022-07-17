import os
from email.message import EmailMessage
import ssl
import smtplib
import datetime


# import ssl


# SERVER = "smtp.yandex.ru"
# FROM = "traisit.m@outlook.com"
# TO = ["traisit.m@outlook.com"] # must be a list

# SUBJECT = "Hello!"
# TEXT = "This is a test of emailing through smtp of example.com."

# # Prepare actual message
# message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

# %s
# """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

# # Send the mail
# import smtplib
# server = smtplib.SMTP(SERVER)
# server.login("traisit.m@outlook.com", "password")
# server.sendmail(FROM, TO, message)
# server.quit()





email_sender = "traisit.m@outlook.com"
email_password = "password"
email_receiver = "traisit.m@outlook.com"

sys_date_now = datetime.datetime.now().strftime("%Y-%m-%d")

subject = 'Python project automate sending email'

body = """
[This is automatic mail, please don't reply]
Starting program ...
"""

em = EmailMessage()
em['From'] = email_sender
em['To'] = email_receiver
em['Subject'] = subject
em.set_content(body)


context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.office365.com',587, context=context) as smtp:
    smtp.login(email_sender,email_password)
    smtp.sendmail(email_sender,email_receiver,em.as_string())


