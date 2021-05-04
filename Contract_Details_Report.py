#!/usr/bin/env python
# coding: utf-8

# ### Contract Workspace Document Report
#  
# #### The program will download 2 reports from Ariba, the Contract Migration - MS and Contract Documents - MS from Public Reports. It will then do some formatting on these reports before saving them to 1 Excel workbook and emailing that workbook to the end user upon completion. 
# 
# #####  Open Chrome, navigate to Ariba Public Reports and download the Document Report
# #####  Go back to Public Reports and download the Contract Migration - MS report
# 
# ##### "Contract Documents - MS"
# ##### 1. Read both the original document report (to retain hard-coded document tags) and the new Contract Documents - MS report into Pandas
# ##### 2. Create the function that will determine what a document is tagged as, based on conditions set by Nicole. This applies to the new Ariba report only. 
# ##### 3. Concatenate the original report with the new report, drop duplicates based on last value in the Document Id column. This will default to Nicole's original file in the event of a duplicate, preserving the hard-coding.
# 
# ##### "Contract Migration - MS"
# ##### 1. Read the Contract Migration - MS report into Pandas
# ##### 2. Convert Date columns to Datetime
# ##### 3. Create 7 new columns: TOM, Supplier Status, Parent Agreement Type L1, Master Agreement, Master Agreement Current Status, Master Agreement TOM Status and Exclusions.
# 
# 
# #####  Write both dataframes (Contract Migration and Contract Documents) to excel with each dataframe being on sheets titled "ETRADE Contract Details Report" and "Document Details Report" respectively
# 
# #####  Email the new file to end user on Monday and Thursday ~9AM once complete

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


# In[ ]:


# Download and save the reports

#specify the webdriver (broswer)
driver = webdriver.Chrome(executable_path = 'chromedriverpath\chromedriver.exe')
 
#URL of website
url = "sharepoint.aspx"
 
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

#Download Document Report

#Click Document Report
document_report_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_t7o7wc"]')))  
document_report_button.click()
print('Document Report clicked')

#Click Export
export_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, '#_1\$lyoc')))  
export_button.click()
print('Export clicked')

#wait for Done 
done_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_qjzssb")))  
done_button.click()
print('Document report downloaded')


# In[ ]:


#read document report into pandas from the downloads folder
list_of_files = glob.glob(r'C:\Users\User\Downloads\*.csv') #* means all if need specific format then *.csv
latest_document_file = max(list_of_files, key=os.path.getctime)

contract_documents = pd.read_csv(latest_document_file)
print("document report read into pandas")


# In[ ]:


#now download the contract details report

#time.sleep(30)
#Click Contract Migration - MS
contract_migration_report = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_ghlkrb"]')))  
contract_migration_report.click()
print('Contract Migration - MS clicked')

#Click Export
contract_export_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_4lont")))  
contract_export_button.click()
print('Export clicked')

#wait for Done then close the browser
done_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_qjzssb")))
done_button.click()

print('contract details report downloaded')


# In[ ]:


#read the contract details report into pandas from the download folder
contract_details = glob.glob(r'C:\Users\User\Downloads\*.csv') #* means all if need specific format then *.csv
latest_file_details = max(contract_details, key=os.path.getctime)

contract_details = pd.read_csv(latest_file_details)
print('contract details report read into pandas')


# ## Format the Document Report
# 

# In[ ]:


#import the original Documents report with hardcoded data
orig_report = pd.read_excel(r'Static Documents File.xlsx')


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

contract_documents_concat = pd.concat([orig_report,contract_documents]).drop_duplicates(
                                                        subset = ['[DOC] Document Id'], keep = 'first').reset_index(drop=True)
contract_documents_concat.head()


# ## Format the Contract Details Report

# In[ ]:


#read the contract details report into pandas from the download folder

contract_details = glob.glob(r'C:\Users\adowd\Downloads\*.csv') #* means all if need specific format then *.csv
latest_file_details = max(contract_details, key=os.path.getctime)

contract_details = pd.read_csv(latest_file_details)
contract_details.fillna("Unclassified", inplace = True)

contract_details.drop_duplicates(subset = ['Project - Project Id'],keep = 'first').reset_index(drop=True)


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
contract_details.columns


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

#contract_details.drop(['Project - Project Id_x'], 
                      #axis = 1, 
                      #inplace = True)

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


# In[ ]:


# create Exclusions column based on the Contract Type and whether the Expiration Date is before 11/1/2018

ex_date = pd.to_datetime('11/1/2018')

def exclude(y):
    if y['Contract Type'] == 'Statement of Work' and y['Contract - Expiration Date'] < ex_date:
        ex = 'SOW Excl'
    else:
        ex = 'Unclassified'
    return ex

contract_details['Exclusions'] = contract_details.apply(exclude, axis = 1)


# In[ ]:


# Download the Commodity and Team report for the additional tabs

#Click Commodity and Team Report - MS
team_report = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_gb4uuc"]')))  
team_report.click()
print('Commodity and Team Report - MS clicked')

#Click Export
export_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="_8k95nd"]')))  
export_button.click()
print('Export clicked')

#wait for Done then close the browser
done_button = WebDriverWait(driver,1000).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#_qjzssb")))
driver.close()


# In[ ]:


#import the report to pandas

two_tabs = glob.glob(r'C:\Users\User\Downloads\*.csv') #* means all if need specific format then *.csv
latest_file_teams = max(two_tabs, key=os.path.getctime)

teams_and_commodities = pd.read_csv(latest_file_teams)
teams_and_commodities.columns


# In[ ]:


# create the report for the Commodity tab

commodity_df = teams_and_commodities[['[PCW] Contract Id','[PCW]Project (Project Id)',
                                 '[PCW]Commodity (Commodity)', '[PCW]Commodity (Commodity ID)',
                                     '[PCW]Contract (Contract)', '[PCW] Description']]
commodity_df = commodity_df.drop_duplicates(subset = ['[PCW] Contract Id', 
                                        '[PCW]Commodity (Commodity ID)']).reset_index(drop=True)


# In[ ]:


# create the Teams tab

team_df = teams_and_commodities[['[PCW]Project (Project Id)', '[PCW] Contract Id','[PGP] Member Name', '[PGP] Member ID']]


# ### Write both reports to Excel and email the workbook

# In[ ]:


print('writing reports')
from openpyxl import load_workbook

#delete the existing file, if applicable
try:
    os.remove(r'pathtoexistingfile.xlsx')
except:
    pass
    
contract_documents_concat.to_excel(r"pathtonewfile.xlsx",
                            index = False)

path = r"pathtocontract_documents.xlsx"

book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

hide_sheet = book.get_sheet_names()
ws = book.get_sheet_by_name('Sheet1')
ws.sheet_state = 'hidden'

contract_documents_concat.to_excel(writer, sheet_name = 'Document Details', index = False)
contract_details.to_excel(writer, sheet_name = 'ETRADE Contract Details Report', index = False)
commodity_df.to_excel(writer, sheet_name = 'Commodity', index = False)
team_df.to_excel(writer, sheet_name = 'Teams', index = False)

writer.save()
writer.close()

print("writing finished")


# In[ ]:


#email the report to Nicole

path = r"path_to_report_location.xlsx"

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email1@email.com; email2@email.com'
mail.Subject = 'Contract Details and Document report'
mail.Body ="""Good morning,

Attached is the current contract details and document details report.

Let me know if you have any questions"""

#attach the report to the email
mail.Attachments.Add(path)
mail.Send()

print('email sent')  


# In[ ]:




