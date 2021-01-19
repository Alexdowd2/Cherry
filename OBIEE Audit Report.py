#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import datetime         # for date operations.
import os               # operating system library to communicate with the operating system.
import win32com.client  # windows API for Windows 95 and beyond. Allows communication with Outlook.


path = os.path.expanduser("path_to_save_attachment")
today = datetime.date.today()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)#.Folders.Item("OBIEE Audit Report")
                                     # The code above is commented out but would otherwise point to a specific Outlook folder  
messages = inbox.Items

# this function establishes the set of steps which will retrieve the email attachment based on the email subject
def saveattachments(subject):
    for message in messages:
        if message.Subject == subject and message.Senton.date() == today:
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveAsFile(os.path.join(path, str(attachment)))

# this executes the function on the email found in the Inbox today with that subject           
saveattachments('Email Subject')

print('attachment downloaded')


# In[ ]:


# this block of code reads the email attachment into excel and excludes the Oracle formatting (starts at row 2)
# it then write the file, as a text file, to the O:Drive where Tableau will read from

import pandas as pd
import numpy as np

download_file = r'path_to_excelfile.xlsx' 

df = pd.read_excel(download_file,
                  header = 2)
df.to_csv(r'path_to_excel_file.csv',
           index = None)

# delete original download file
os.remove(download_file)
print('report save complete')


# ### The next two blocks of code dosome of the text modification required for Hadoop. Specifically, re-codes the line endings and dates. 

# In[ ]:


audit_report = pd.read_csv(r'path_to_excel_file.csv')
print('audit report read')

print('reformatting dates')
audit_report = audit_report.astype({'AUDIT_TIME': 'datetime64[ns]', 'NEW_END_DATE_ACTIVE': 'datetime64[ns]', 
                                    'OLD_END_DATE_ACTIVE': 'datetime64[ns]', 'NEW_TAX_VERIFICATION_DATE': 'datetime64[ns]',
                                   'OLD_TAX_VERIFICATION_DATE': 'datetime64[ns]', 'NEW_START_DATE': 'datetime64[ns]',
                                   'OLD_START_DATE': 'datetime64[ns]', 'NEW_END_DATE': 'datetime64[ns]', 
                                    'OLD_END_DATE': 'datetime64[ns]'})
print('reading to txt')
audit_report.to_csv(r'path_to_excel_file.txt',
                   index = False, sep = '\t')


# In[ ]:


path = 'path_to_excel_file.txt'
WINDOWS_LINE_ENDING = b'\r\n'
UNIX_LINE_ENDING = b'\n'

# file path
with open(path, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(path, 'wb') as open_file:
    open_file.write(content)
    
print('Complete')

