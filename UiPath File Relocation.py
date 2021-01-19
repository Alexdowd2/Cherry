#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# take the files that UiPath downloads and read them into pandas

import pandas as pd

contract_report = pd.read_excel('filepath.xlsx')
print('contract report complete')

requisition_report = pd.read_excel('filepath.xlsx')
print('requisition report complete')

PO_report = pd.read_excel('filepath.xlsx')
print('PO report complete')

invoice_report = pd.read_excel('filepath.xlsx')
print('invoice report complete')

invoice_spend_report = pd.read_excel('filepath.xlsx')
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

contract_report.to_excel('filepath.xlsx.xlsx',
                         index=False, 
                         sheet_name = 'Contract_Master')
print('contract report relocated')

requisition_report.to_excel('filepath.xlsx.xlsx',
                           index=False, 
                           sheet_name = 'Requisition_Master')
print('requisition report relocated')

PO_report.to_excel('filepath.xlsx.xlsx',
                   index=False, 
                   sheet_name = 'PO_Master')
print('PO report relocated')

invoice_report.to_excel('filepath.xlsx.xlsx',
                        index=False, 
                        sheet_name = 'Invoice Report_Master')
print('invoice report relocated')

invoice_spend_report.to_excel('filepath.xlsx.xlsx',
                              index=False, 
                              sheet_name = 'Total Invoice Spend_Master')

print('total invoice spend report relocated')
print("")
print('files relocated')


# ### Convert files to .txt

# In[ ]:


contract_report.to_csv(r"filepath.txt")
requisition_report.to_csv(r"filepath.txt")
PO_report.to_csv(r"filepath.txt")
invoice_report.to_csv(r"filepath.txt")
invoice_spend_report.to_csv(r"filepath.txt")


# ### Change line endings for Hadoop upload, if necessary

# In[ ]:


path1 = r"filepath.txt.txt"
path2 = r"filepath.txt"
path3 = r"filepath.txt"
path4 = r"filepath.txt"
path5 = r"filepath.txt"

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

