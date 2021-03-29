#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#import the libraries
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait #tells webdriver to wait until a button appears.
                                                        #it will wait for up to 100 seconds
from selenium.webdriver.support import expected_conditions as EC

import time
import glob          #glob is used to retrieve files/pathnames matching a specified pattern
import os            #this module allows python to communicate with operating systems
import pandas as pd  #Pandas is a library for reading tables
import win32com.client as win32
import numpy as np


# ## Download and save the reports

# In[ ]:


#specify the webdriver (broswer)
driver = webdriver.Chrome(executable_path = 'path_to_driver')
 
#URL of website
url = "url.com"
 
#Open the website
driver.get(url)
driver.maximize_window()

#click Upstream
downstream = driver.find_element_by_css_selector('#WebPartWPQ3 > div.ms-rtestate-field > table > tbody > tr:nth-child(1) > td > p:nth-child(4) > a')
downstream.click()
print('Upstream clicked')

#click Manage
manage_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_s2d3v")))  
manage_button.click()
print('Manage clicked')

#click Public Reports
public_reports_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_sr\$opc")))  
public_reports_button.click()
print('Public Reports clicked')

time.sleep(3)
#Click Contract Reports
contract_reports_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_peqkcb"]')))  
contract_reports_button.click()

print('Contract Reports clicked')

#Click Open
open_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_xobkvc > b")))  
open_button.click()
print('Open clicked')

#Click Document Report
document_report_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="__vxjpd"]')))  
document_report_button.click()
print('Document Report clicked')

#Click Export
export_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, '#_xfsxnd')))  
export_button.click()
print('Export clicked')

#wait for Done 
done_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_qjzssb")))  
done_button.click()
print('Document report downloaded')


# In[ ]:


list_of_files = glob.glob(r'C:\Users\username\Downloads\*.csv') #* means all if need specific format then *.csv
latest_document_file = max(list_of_files, key=os.path.getctime)

contract_documents = pd.read_csv(latest_document_file)
print("document report read into pandas")


# In[ ]:


time.sleep(5)
#Click Contract Migration - MS
contract_migration_report = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_vps4u"]')))  
contract_migration_report.click()
print('Contract Migration - MS clicked')

#Click Export
contract_export_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_afxoib")))  
contract_export_button.click()
print('Export clicked')

#wait for Done then close the browser
done_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_qjzssb")))
driver.close()
print('contract details report downloaded')


# In[ ]:


#read the contract details report into pandas from the download folder
contract_details = glob.glob(r'C:\Users\username\Downloads\*.csv') #* means all if need specific format then *.csv
latest_file_details = max(contract_details, key=os.path.getctime)

contract_details = pd.read_csv(latest_file_details)
print('contract details report read into pandas')


# ## Format the Document Report

# In[ ]:


#import the original Documents report with hardcoded data
orig_report = pd.read_excel(r'filepath.xlsx')


# In[ ]:


# define function to create the 'Doc Tag' column in the new report from Ariba

def doc_tag(row):
    if "Signed_" in row['[DOC] Name'] and ".pdf" in row['[DOC] Name']:
        val = "Signed .PDF"
    
    elif "Signed_" in row['[DOC] Name']:
        val = "Signed .PDF"
    
    elif ".DOCX" in row['[DOC] Name'] or ".DO" in row['[DOC] Name'] or ".do" in row['[DOC] Name']:
        val = "Not .PDF"
    
    elif ".msg" in row['[DOC] Name']: 
        val = "EMAIL"
    
    elif "Signed_" not in row['[DOC] Name'] and ".pdf" in row['[DOC] Name']:
        val = "Other PDF"
    
    elif "Legal Approval" in row['[DOC] Name']:
        val = "Legal Approval"
    
    else:
        val = 'Other'
        
    return val

#add the Doc Tag column
contract_documents["Document Tag"] = contract_documents.apply(doc_tag, axis = 1)

#re-order the columns so Document Tag is after Document Name
contract_documents = contract_documents[['[DOC]Project (Project Id)','[DOC]Project (Project Type)','[DOC] Document Id','[DOC] Name','Document Tag',
        '[DOC] Status','[DOC] Type','count(Document)']]
print('DOC Type column added')


# In[ ]:


#concatenate the new data with the original hardcoded data, in the case of duplicates, 
#keep the original value and drop the new one

contract_documents_concat = pd.concat([orig_report,
                                       contract_documents]).drop_duplicates(subset = ['[DOC] Document Id'], 
                                                                     keep = 'first').reset_index(drop=True)


# ## Format the Contract Details Report

# In[ ]:


#read the report into pandas from the download folder

contract_details = glob.glob(r'C:\Users\username\Downloads\*.csv') #* means all if need specific format then *.csv
latest_file_details = max(contract_details, key=os.path.getctime)

contract_details = pd.read_csv(latest_file_details)
contract_details.fillna("Unclassified", inplace = True)
contract_details = contract_details.drop_duplicates(subset = ['Project - Project Id'],keep = 'first').reset_index(drop=True)


# In[ ]:


#define functions for the calculated columns TOM Status, Adjusted TOM status and Supplier Status, then apply them to the table

TOM = pd.to_datetime('11/1/2021')
contract_details['Contract - Expiration Date'] = contract_details['Contract - Expiration Date'].replace("Unclassified", np.nan)
contract_details['Contract - Expiration Date'] = pd.to_datetime(contract_details['Contract - Expiration Date'])


def tom_status(a):
    if a['Contract - Expiration Date'] < TOM: 
        global tomStatus
        tomStatus = "Before TOM"
    elif a['Contract - Expiration Date'] >= TOM:
        tomStatus = "After TOM"
    elif pd.isnull(a['Contract - Expiration Date']):
        tomStatus = 'After TOM'
    return tomStatus

def supplier_status(b):
    if b['Affected Parties - Common Supplier'] != "Unclassified":
        supStatus = "Active"
    else:
        supStatus = "Inactive"
    
    return supStatus


contract_details['TOM Status'] = contract_details.apply(tom_status, axis = 1)
contract_details['Supplier Status'] = contract_details.apply(supplier_status, axis = 1)

#create the adjusted tom status column after the other 2 because it is dependent on the TOM Status column
contract_details['Renewal Decision'].fillna("Unclassfied", inplace = True)
def adj_tom(c):
    if "Renew" in c['Renewal Decision'] and c['TOM Status'] == "Before TOM":
        adj_tom = 'After TOM'
      
    elif "Terminate" in c['Renewal Decision'] and c['TOM Status'] == 'After TOM':
        adj_tom = 'Before TOM'
      
    else:
        adj_tom = c['TOM Status']
        
    return adj_tom

contract_details['Adjusted TOM Status'] = contract_details.apply(adj_tom, axis = 1)


# In[ ]:


#Parent Agreement Type L1 Column

hierarchy_df = contract_details[['Project - Project Id','Hierarchy Type']]
hierarchy_df.rename(columns={'Hierarchy Type': 'Parent Agreement Type L1'},  inplace = True)



#perform the internal VLOOKUP
contract_details = pd.merge(contract_details, 
                            hierarchy_df,
                            how = 'left',
                            left_on = 'Parent Agreement - Project Id',
                            right_on = 'Project - Project Id')

contract_details.drop(['Project - Project Id_y'], 
                      axis = 1, 
                      inplace = True)

#replace the NaN with "No Parent"
contract_details['Parent Agreement Type L1'].fillna('No Parent', inplace = True)

#rename columns
contract_details.rename(columns={'Project - Project Id_x':'Project - Project Id',
                                'Parent Agreement - Project Id_x':'Parent Agreement - Project Id'}, 
                        inplace = True)

contract_details = contract_details.drop_duplicates(subset = ['Project - Project Id'],keep = 'first').reset_index(drop=True)


# In[ ]:


# Create conditions for Master Agreement column

masteragreement_df = contract_details[['Project - Project Id', 'Parent Agreement - Project Id']]
                             
contract_details = pd.merge(contract_details,
                            masteragreement_df,
                            how = 'left',
                            left_on = 'Parent Agreement - Project Id',
                            right_on = 'Project - Project Id')

contract_details.rename(columns={'Parent Agreement - Project Id_y':'Master Agreement_lookup',
                                 'Project - Project Id_x':'Project - Project Id',
                                 'Parent Agreement - Project Id_x':'Parent Agreement - Project Id'}, inplace = True)

contract_details.drop(['Project - Project Id_y'], axis = 1, inplace = True)

contract_details = contract_details.drop_duplicates(subset = ['Project - Project Id'],keep = 'first').reset_index(drop=True)


# In[ ]:


# Create the Master Agreement Column

def master(d):
    if d['Parent Agreement Type L1'] == 'Master Agreement':
        global master
        master = d['Parent Agreement - Project Id']
    elif d['Parent Agreement Type L1'] == 'Sub Agreement':
        master = d['Master Agreement_lookup']
    elif d['Parent Agreement Type L1'] == 'No Parent':
        master = 'N/A'
    return master

contract_details['Master Agreement'] = contract_details.apply(master, axis = 1)
contract_details.drop(['Master Agreement_lookup'], axis = 1, inplace = True)


# In[ ]:


# Master Agreement Current Status column

master_status = contract_details[['Project - Project Id','Master Agreement', 'Contract Status']]

contract_details = pd.merge(contract_details,
                           master_status,
                           how = 'left',
                           left_on = 'Master Agreement',
                           right_on = 'Project - Project Id')

contract_details.drop(['Master Agreement_y','Project - Project Id_y'], axis = 1, inplace = True)

contract_details.rename(columns={'Contract Status_x':'Contract Status',
                                'Master Agreement_x':'Master Agreement', 
                                'Project - Project Id_x':'Project - Project Id',
                                'Contract Status_y':'Master Agreement Current Status'}, inplace = True)

contract_details = contract_details.drop_duplicates(subset = ['Project - Project Id'],keep = 'first').reset_index(drop=True)


# In[ ]:


#Master Agreement TOM Status

master_tom = contract_details[['Project - Project Id','Master Agreement','TOM Status']]

contract_details = pd.merge(contract_details,
                           master_tom,
                           how = 'left',
                           left_on = 'Master Agreement',
                           right_on = 'Project - Project Id')



contract_details.rename(columns={'TOM Status_x':'TOM Status',
                       'Master Agreement_x':'Master Agreement',
                       'TOM Status_y' : 'Master Agreement TOM Status',
                       'Project - Project Id_x':'Project - Project Id'}, inplace = True)

contract_details.drop(['Master Agreement_y','Project - Project Id_y'], axis = 1, inplace = True)
contract_details.head()


# In[ ]:


#Master Agreement Adjusted TOM Status column

masteradj_tom = contract_details[['Project - Project Id','Master Agreement','Adjusted TOM Status']]

contract_details = pd.merge(contract_details,
                           masteradj_tom,
                           how = 'left',
                           left_on = 'Master Agreement',
                           right_on = 'Project - Project Id')



contract_details.rename(columns={'TOM Status_x':'TOM Status','Master Agreement_x':'Master Agreement',
                                 'TOM Status_y' : 'Master Agreement TOM Status','Project - Project Id_x':'Project - Project Id',
                                 'Adjusted TOM Status_x':'Adjusted TOM Status',
                                 'Adjusted TOM Status_y':'Master Agreement Adjusted TOM Status'}, inplace = True)

contract_details.drop(['Master Agreement_y','Project - Project Id_y'], axis = 1, inplace = True)
contract_details.head()


# ### Write both reports to Excel and email the workbook

# In[ ]:


print('writing reports')
from openpyxl import load_workbook

#delete the existing file, if applicable
try:
    os.remove(r'filepath.xlsx')
except:
    pass
    
contract_documents.to_excel(r"filepath.xlsx",
                            index = False)

path = r"filepath.xlsx"

book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

hide_sheet = book.get_sheet_names()
ws = book.get_sheet_by_name('Sheet1')
ws.sheet_state = 'hidden'

contract_documents.to_excel(writer, sheet_name = 'Document Details', index = False)
contract_details.to_excel(writer, sheet_name = 'Contract Details', index = False)
writer.save()
writer.close()

print("writing finished")


# In[ ]:


#email the report
path = r"filepath.xlsx"

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email1@email.com; email2@email.com'
mail.Subject = 'ETRADE Contract Details and Document report'
mail.Body ="""Good morning,

Attached is the current ETRADE contract details and document details report.

Let me know if you have any questions"""

#attach the report to the email
mail.Attachments.Add(path)
mail.Send()

print('email sent')  


# In[ ]:




