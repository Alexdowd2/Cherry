#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# take the files that UiPath downloads and read them into pandas

import pandas as pd

contract_report = pd.read_excel('O:\Procurement Planning\Tableau\RPA Download\PROD\Contracts_Master.xlsx')
print('contract report complete')

requisition_report = pd.read_excel('O:\Procurement Planning\Tableau\RPA Download\PROD\Requisition_Master.xlsx')
print('requisition report complete')

PO_report = pd.read_excel('O:\Procurement Planning\Tableau\RPA Download\PROD\PO_Master.xlsx')
print('PO report complete')

invoice_report = pd.read_excel('O:\Procurement Planning\Tableau\RPA Download\PROD\Invoice Report for Tableau_Master.xlsx')
print('invoice report complete')

invoice_spend_report = pd.read_excel('O:\Procurement Planning\Tableau\RPA Download\PROD\Total Invoice Spend_Master.xlsx')
print('invoice spend report complete')

print("")
print('UiPath download files read')
print("")


# ### Move the "Description" column to the end so that it doesn't compromise Hadoop

# In[ ]:


cols = list(contract_report.columns.values)
cols.pop(cols.index('Description'))
contract_report = contract_report[cols+['Description']]

cols = list(requisition_report.columns.values)
cols.pop(cols.index('Description'))
requisition_report = requisition_report[cols+['Description']]

cols = list(PO_report.columns.values)
cols.pop(cols.index('Description'))
PO_report = PO_report[cols+['Description']]

cols = list(invoice_report.columns.values)
cols.pop(cols.index('Description'))
invoice_report = invoice_report[cols+['Description']]

cols = list(invoice_spend_report.columns.values)
cols.pop(cols.index('Description'))
invoice_spend_report = invoice_spend_report[cols+['Description']]


# In[ ]:


# save the files to Alex's Tableau Folder

contract_report.to_excel('O:\Procurement Planning\Tableau\Alex Tableau\Datasources\EPIC Operational Metrics\Contracts_Master.xlsx',
                         index=False, 
                         sheet_name = 'Contract_Master')
print('contract report relocated')

requisition_report.to_excel('O:\Procurement Planning\Tableau\Alex Tableau\Datasources\EPIC Operational Metrics\Requisition_Master.xlsx',
                           index=False, 
                           sheet_name = 'Requisition_Master')
print('requisition report relocated')

PO_report.to_excel('O:\Procurement Planning\Tableau\Alex Tableau\Datasources\EPIC Operational Metrics\PO_Master.xlsx',
                   index=False, 
                   sheet_name = 'PO_Master')
print('PO report relocated')

invoice_report.to_excel('O:\Procurement Planning\Tableau\Alex Tableau\Datasources\EPIC Operational Metrics\Invoice Report for Tableau_Master.xlsx',
                        index=False, 
                        sheet_name = 'Invoice Report_Master')
print('invoice report relocated')

invoice_spend_report.to_excel('O:\Procurement Planning\Tableau\Alex Tableau\Datasources\EPIC Operational Metrics\Total Invoice Spend_Master.xlsx',
                              index=False, 
                              sheet_name = 'Total Invoice Spend_Master')

print('total invoice spend report relocated')
print("")
print('files relocated')


# ### Convert files to .txt

# In[ ]:


contract_report.to_csv(r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\EPIC Metrics\Contracts.txt")
requisition_report.to_csv(r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\EPIC Metrics\Requisitions.txt")
PO_report.to_csv(r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\EPIC Metrics\POs.txt")
invoice_report.to_csv(r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\EPIC Metrics\Invoice.txt")
invoice_spend_report.to_csv(r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\EPIC Metrics\Invoice_Spend.txt")


# ### Change line endings for Hadoop upload, if necessary

# In[ ]:


path1 = r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\Daily Invoice Reports\12_Month_Invoice_Report_1.7.2021.txt"
path2 = r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\Daily Invoice Reports\12_Month_Invoice_Report_1.8.2021.txt"
path3 = r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\Daily Invoice Reports\12_Month_Invoice_Report_12.9.2020.txt"
path4 = r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\Daily Invoice Reports\12_Month_Invoice_Report_12.11.2020.txt"
path5 = r"O:\Procurement Planning\Tableau\Alex Tableau\Hadoop\Hadoop Files (Current)\Txt Files\Daily Invoice Reports\12_Month_Invoice_Report_12.15.2020.txt"

# replacement strings
WINDOWS_LINE_ENDING = b'\r\n'
UNIX_LINE_ENDING = b'\n'

# file path
file_path1 = path1
with open(file_path1, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(file_path1, 'wb') as open_file:
    open_file.write(content)
    

file_path2 = path2
with open(file_path2, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(file_path2, 'wb') as open_file:
    open_file.write(content)
    

file_path3 = path3
with open(file_path3, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(file_path3, 'wb') as open_file:
    open_file.write(content)
        
    
file_path4 = path4
with open(file_path4, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(file_path4, 'wb') as open_file:
    open_file.write(content)
        
file_path5 = path5
with open(file_path5, 'rb') as open_file:
    content = open_file.read()

content = content.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

with open(file_path5, 'wb') as open_file:
    open_file.write(content)
    
print('Complete')

