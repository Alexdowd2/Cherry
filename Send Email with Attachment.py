#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import smtplib
from email.mime.text import MIMEText

# For guessing MIME type
import mimetypes

# Import the email modules we'll need
import email
import email.mime.application


# In[ ]:


your_email = 'alexdowd2@gmail.com'
password = 'xnwbozmkmqqovdmh'


# In[ ]:


recipients = pd.read_excel('/Users/alexdowd/Documents/contactEmails_test.xlsx')

names = recipients['Name']
emails = recipients['Email']
subject = 'Python Email'


# In[ ]:


server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, password)

for i in range(len(emails)):
    name = names[i]
    email = emails[i]
    
    message = MIMEText('Good morning ' + name + ',' + """

This is a test email, have a great day!""")
    message['Subject'] = subject
    message['From'] = your_email
    message['To'] = email
    
    server.sendmail(your_email, email, message.as_string())
    
server.close()


# In[ ]:




