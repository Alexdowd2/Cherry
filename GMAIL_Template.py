#!/usr/bin/env python
# coding: utf-8

# In[]:


#library for communicating with Google server
import smtplib


recipients = ['recipient1@emial.com','recipient2@email.com']

#your email address
gmail_user = 'me@gmail.com'

#your gmail password.
#you may have to create an App Password
gmail_password = 'password'
sent_from = gmail_user
to = recipients
subject = 'Hooray for Python!'
body = """
Now there's a subject line! 
If you want to send some more emails 
though, I have the program to do it!
"""

email_text = """From: %s
To: %s
Subject: %s

%s
""" % (sent_from,", ".join(to), subject, body)

#we are making an encrypted connetion with SMTP_SSL and port 465
server = smtplib.SMTP_SSL('smtp.gmail.com',465)    

server.ehlo() #ehlo authenticates us to the gmail server

server.login(gmail_user, gmail_password)

server.sendmail(sent_from, to, email_text)

server.close()
    
print('email sent!')

    


# In[ ]:




