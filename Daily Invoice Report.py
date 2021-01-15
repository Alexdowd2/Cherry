#!/usr/bin/env python
# coding: utf-8

# 
# ### CHANGE THE SUBMIT DATE TO YESTERDAY
# 
# #### This program does the following things, and where relevant, there are comments offering additional details
# 
# ##### 1. Opens Chrome, navigates to Ariba Downstream finds and then downloads the daily medium-dollar invoice report
# ##### 2. Goes to the Downloads folder and finds the most recent .csv file (which will be the report that just downloaded), opens it, formats it, and emails it to Beth Hoffman, Mike Zinone and Ladonna Rose-Hawkins. 

# ### This segment of code will import selenium webdriver which allows Python to navigate web browsers, in this case chrome. There are delays built into the program (the time.sleep() command) to prevent errors that are the result of slow page openings. At the conclusion of the report download, Python closes the browser window.

# In[ ]:


print('importing Time')
import time
print('Time imported')
# importing webdriver from selenium
print('importing webdriver')
from selenium import webdriver
print('webdriver imported')
print("")
print("Web Navigation Steps")
print("")

driver = webdriver.Chrome(executable_path = 'O:\Procurement Planning\QA\Python\chromedriver.exe')
 
# URL of website
url = "https://channele.corp.etradegrp.com/communities/initiative/epic/Pages/Environments.aspx"
 
# Open the website
driver.get(url)
driver.maximize_window()

#click Downstream
downstream = driver.find_element_by_css_selector('#WebPartWPQ3 > div.ms-rtestate-field > table > tbody > tr:nth-child(1) > td > p:nth-child(6) > span > a > span')
downstream.click()
print('Downstream clicked')

time.sleep(10)

#click 'Manage'
manage = driver.find_element_by_xpath('//*[@id="_ydbfdb"]')
manage.click()
print('Manage clicked')

time.sleep(2)

#click 'Public Reports'
public_reports = driver.find_element_by_css_selector('#_b28imd')
public_reports.click()
print('public reports clicked, now waiting')

time.sleep(3)

#click 'Invoice Reports' 
invoice_reports = driver.find_element_by_xpath('//*[@id="_sgkkbd"]')
invoice_reports.click()
print("invoice report clicked, now waiting")

time.sleep(2)

#click 'Open'
open_button = driver.find_element_by_css_selector('#_if6\$bd > b')
open_button.click()
print('Open clicked, now waiting')

time.sleep(2)

#click 'Medium-Dollar Invoices for Audit'
audit_file = driver.find_element_by_css_selector('#_rnq0h')
audit_file.click()
print('audit file clicked, now waiting')

time.sleep(2)

#click 'Export'
export = driver.find_element_by_css_selector('#_jiuf9d')
export.click()
print('Export clicked')

time.sleep(10)
driver.close()


# ### This segment of code waits 5 seconds, then it finds the most recently downloaded .csv file in the Downloads folder, which will be the report that was just downloaded from Ariba, and reads it into a pandas dataframe. 

# In[ ]:


time.sleep(5)
import glob          # glob is used to retrieve files/pathnames matching a specified pattern
import os            # this module allows python to communicate with operating systems
import pandas as pd  # Pandas is a library for reading tables

list_of_files = glob.glob(r'C:\Users\adowd\Downloads\*.csv') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)

df = pd.read_csv(latest_file, thousands=",")
df.head()


# ### The rest of the blocks of code will format the report and ultimately email the report to the designated recipients. In the event there are no invoices that match the audit criteria, and email will go to the designated recipients explaining that no such invoices were found. Otherwise, the audit file will be attached in the email. 

# In[ ]:


# convert the csv file to an xlsx file and begin formatting

df.to_excel(r'O:\Procurement Planning\QA\QA Reviews\Daily Medium-Dollar Invoice Report\Daily Invoice Report.xlsx', index = False)
print('file read to O:Drive')


# In[ ]:


import numpy as np
from openpyxl import load_workbook
import win32com.client as win32
import datetime as dt
from pandas.tseries.offsets import BDay

path = r'O:\Procurement Planning\QA\QA Reviews\Daily Medium-Dollar Invoice Report\Daily Invoice Report.xlsx'

df = pd.read_excel(path)
df.replace('Unclassified',np.nan, inplace=True)

df['Invoice Submit Date - Date'] = pd.to_datetime(df['Invoice Submit Date - Date'])
df['Invoice Date Created - Date'] = pd.to_datetime(df['Invoice Date Created - Date'])
df['Invoice Date - Date'] = pd.to_datetime(df['Invoice Date - Date'])
df['Approved Date - Date'] = pd.to_datetime(df['Approved Date - Date'])

df.head()


# In[ ]:


# filter on yesterday's submitted invoices

today = dt.datetime.today()
df['previous_day'] = (today - BDay(1))
df['previous_day'] = df['previous_day'].dt.normalize()
previous_day = df['previous_day']


df = df[(df['Invoice Submit Date - Date']) == previous_day]
df.head()


# In[ ]:


# Function to create the 'criteria met' column

def f(row):
    if row['sum(Invoice Spend)'] >= 10000 and row['sum(Invoice Spend)'] < 100000:
        val = 'Yes'
    else:
        val = 'No'
    
    return val

# add the criteria met column
df['Criteria Met'] = df.apply(f,axis=1)

print(len(df))
df.head()


# In[ ]:


# Filter on criteria met
df = df[(df['Criteria Met'] == 'Yes')]

print(len(df))
df.head()


# In[ ]:


# randomly sample data. 
# The IF argument checks if there are more or less than 10 rows. If there are 10 or fewer then Python selects distinct
# records based on Invoice ID. If there are more than 10 Python selects a random 10. 

if len(df) > 10:
    df = df.sample(10)
elif len(df) <= 10:
     df = df.groupby('Invoice ID').apply(lambda df: df.sample(1))
        
print(len(df))        
df

# this IF argument determines which email goes out
if len(df) == 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'recipient1@email.com; recipient2@email.com'
    mail.Subject = 'Daily Medium-Dollar Invoice Report for Audit'
    mail.Body = "Mike, no records were identified that met the audit criteria. Let me know if you have any questions or if this is incorrect."
    mail.Send()
    print('no files email sent')

else:
  # drop the criteria met and previous day columns from the df and begin to write the audit file
    df.drop(['Criteria Met'],axis=1, inplace = True)
    df.drop(['previous_day'],axis=1, inplace = True)
    
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book

    # hide the original export sheet
    hide_sheet = book.get_sheet_names()
    ws = book.get_sheet_by_name('Sheet1')
    ws.sheet_state = 'hidden'

    # create new 'Sample' sheet
    df.to_excel(writer, sheet_name='Sample',index=False)
    writer.save()
    writer.close()

    print('audit sample file created')
    
    # send email with the audit file
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'recipient1@email.com; recipient2@email.com'
    mail.Subject = 'Daily Medium-Dollar Invoice Report for Audit'
    mail.Body = "Mike, please see attached for invoices for audit. Let me know if you have any questions."

    # attach the audit file to the email
    
    mail.Attachments.Add(path)

    mail.Send()

    print('email sent')  

