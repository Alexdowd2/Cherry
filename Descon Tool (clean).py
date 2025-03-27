#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import os
import numpy as np
import datetime as dt
import glob
import openpyxl

sheetName = 'RFP Scoring Summary'
path = r"folder path to excel files\\"

bid_master = pd.read_excel(r'path to excel',
                              sheet_name = name of sheet to read)

li = []
for file in glob.glob(path + '*.xlsm'):
    li.append(file)

final_li = [x for x in li if '~' not in x]

  #check if sheet needs to be renamed
for i in final_li:    
    wb = openpyxl.load_workbook(i)
    sheet = wb.active
    
    if sheet.title != sheetName:
        sheet.title = sheetName
        wb.save(i)
        wb.close()

init = []

for f in final_li:
    print('reading {}'.format(f))
  
    ###Create dataframes for macro project details
    file = f 
    cols = ['Presenation Date', 
        'Sourcing Manager', 
        'Requestor', 
        'Key Stakeholdres',
        'Project Manager',
        'Ariba WS',
        'CIIPS',
        'MER',
        'Recommendation']
    cols_2 = ['Project Type',
        'Project Address',
        'Project City',
        'Project State',
        'Project Zip Code',
        'Project Region',
        'Notify Banking?',
        'Project Start Date',
        'Project End Date']
    df = pd.DataFrame(data = cols)
    
    data_1 = pd.read_excel(f, sheet_name = sheetName, usecols = 'C', skiprows = 3, nrows = 9).transpose()
    data_2 = pd.read_excel(f, sheet_name = sheetName, usecols = 'G', skiprows = 3, nrows = 9).transpose()

    data_1.columns = cols
    data_2.columns = cols_2

    macro_df = data_1.merge(data_2, how = 'cross')
    
    ### obtain bidder data

    df_vendors = pd.read_excel(f, sheet_name = sheetName, usecols = 'B:B, D:I', skiprows = 20, nrows = 6).transpose().reset_index()
    df_vendors = df_vendors.dropna(axis=1, how ='all')

    #convert first row to column headers
    new_header = df_vendors.iloc[0].tolist()
    df_vendors = df_vendors[1:]
    df_vendors.columns = new_header

    #remove new line separator from participating bidders
    df_vendors = df_vendors.rename(columns={'Other Considerations':'Participating Bidders', 'Innovation Assessment \nLow - Vendor does not provide innovation \nMedium - Vendor provides some innovation\nHigh - Vendor provides market-leading innovation':'Innovation Assessment'})
    df_vendors['Participating Bidders'] = df_vendors['Participating Bidders'].str.replace('\n', ' ')

    #delete any empy vendor rows
    df_vendors = df_vendors[~df_vendors['Participating Bidders'].str.contains('Unnamed', na=False)]
    #df_vendors
    
    #Bidder Due Diligence
    df_bid_dd = pd.read_excel(f, sheet_name = sheetName, usecols = 'B:B, C:H', skiprows = 35, nrows = 7).transpose().reset_index()

    #convert first row to column headers
    new_header_2 = df_bid_dd.iloc[0].tolist()
    df_bid_dd = df_bid_dd[1:]
    df_bid_dd.columns = new_header_2

    #delete any empy vendor rows
    df_bid_dd = df_bid_dd[~df_bid_dd['Participating Bidders'].str.contains('Unnamed', na=False)]
    df_bid_dd = df_bid_dd[df_bid_dd['Participating Bidders'] != 'N/A']
    try:
        df_bid_dd = df_bid_dd.drop(df_bid_dd.loc[:,'Bidder Branch Address' : 'Bidder Phone'].columns, axis = 1)
    except:
        pass
    #remove /n line separator
    df_bid_dd['Participating Bidders'] = df_bid_dd['Participating Bidders'].str.replace('\n', ' ')

    ###Budget Amount vs. Award Amount
#     df_budget = pd.read_excel(f, sheet_name = sheetName, usecols = 'B:C', skiprows = 45, nrows = 6).transpose().reset_index()

#     #convert first row to column headers
#     new_header_budget = df_budget.iloc[0].tolist()
#     df_budget = df_budget[1:]
#     df_budget.columns = new_header_budget
#     try:
#         df_budget = df_budget.rename(columns = {'Contingency ': 'Contingency'})
#     except:
#         pass
    
    
    ###RFP Process
    df_rfp_pro = pd.read_excel(f, sheet_name = sheetName, usecols = 'B', skiprows = 53, nrows = 1).transpose().reset_index()
    df_rfp_pro = df_rfp_pro.dropna(axis = 1, how = 'all')
    df_rfp_pro = df_rfp_pro.rename(columns={'index': 'RFP Process'})
    
    ###RFP Data Summary

    df_rfp = pd.read_excel(f, sheet_name = sheetName, usecols = 'B:H', skiprows = 57, nrows = 10).transpose().reset_index()

    #convert first row to column headers
    new_header_rfp = df_rfp.iloc[0].tolist()
    df_rfp = df_rfp[1:]
    df_rfp.columns = new_header_rfp

    #remove new line separator from participating bidders
    df_rfp['Participating Bidders'] = df_rfp['Participating Bidders'].str.replace('\n', ' ')
    df_rfp = df_rfp.drop(df_rfp.columns[3], axis =1)
    df_rfp = df_rfp[df_rfp['Participating Bidders'] !='N/A']
    try:
        df_rfp = df_rfp.drop('Weighted scoring (50% pricing, 50% non-pricing)', axis = 1)
    except:
        pass
    df_rfp = df_rfp[~df_rfp['Participating Bidders'].str.contains('Unnamed', na=False)] 
    
    ### Award Recommendation

    df_award = pd.read_excel(f, sheet_name = sheetName, usecols = 'B:C', skiprows = 69, nrows = 2).transpose()
    #df_award.insert(0, 'Pipe1', ['|'])
    #df_award.insert(2, 'Pipe2', ['|'])

    df_award[1] = df_award[1].str.replace('\n',' ')

    df_award = df_award.reset_index()
    #df_award.rename(columns={'index': 'Column'})

    #convert first row to column headers
    new_header = df_award.iloc[0].tolist()
    df_award = df_award[1:]
    df_award.columns = new_header
    
    ##merge all dataframes
    df_1 = df_vendors.merge(macro_df, how = 'cross') 
    df_2 = df_award.merge(df_rfp, how = 'cross')

    #df_2.drop(['Participating Bidders'], axis = 1, inplace = True)
    df_1 = df_1.merge(df_2, right_on = 'Participating Bidders', left_on = 'Participating Bidders')
    df_1 = df_1[df_1['Participating Bidders'].str.contains('Unnamed') == False]
    df_1 = df_rfp_pro.merge(df_1, how = 'cross')
    #df_1 = df_1.merge(df_3, how = 'cross')
    df_1 = df_1.merge(df_bid_dd, how = 'inner')

    df_1 = df_1.rename(columns={'Award Amount_x' : 'Award Amount', 
                                'Participating Bidders_x' : 'Participating Bidders',
                               'Recommendation_y':'Recommendation'})
    
    df_1['Ariba WS'] = df_1['Ariba WS'].str.replace('\n', '')
    df_1['Ariba WS'] = df_1['Ariba WS'].str.strip()
    
    #function to populate contract year
    import pyodbc
    conn = pyodbc.connect(r'Driver={SQL Server};'
                          r'Server=#server name;'
                          r'Database=database name;'
                          r'Trusted_Connection=yes;'
                         )
    cursor = conn.cursor()


    df_cwyear = pd.read_sql('select distinct [Parent Project - Project Id], YEAR([Effective Date - Date]) as [Contract Year] from tbl_ariba_cw', conn)
    #df_cwyear.head()

    df_1 = df_1.merge(df_cwyear, how = 'left', left_on = 'Ariba WS', right_on = 'Parent Project - Project Id')
    
    df_1 = df_1.loc[:,['Participating Bidders',
                  'Presenation Date',
                  'Sourcing Manager',
                  'Requestor',
                  'Key Stakeholdres', #rename to correct spelling
                  #blank
                  #'CMS#', try to look up from SQL DB
                  'Ariba WS',
                  'CIIPS',
                  'MER',
                  'Project Type',
                  'Project Address',
                  'Project City',
                  'Project State', 
                  #'Project State - City', manually create
                  'Project Zip Code',
                  'Project Region',
                  'Notify Banking?',
                  'Project Start Date',
                  'Project End Date',
                  'Recommendation',
                  'Contract Year',
                  #'Project Description',
                 # 'Bidder Branch Address',
                 # 'Bidder Contact Name',
                 # 'Bidder Email',
                 # 'Bidder Phone',
                 # 'Bidder Source',
                   'Diverse Supplier',
                 # 'Diversity Certified',
                #  'Client',
                #  'Decision to Bid',
                 # 'MER Budget Amount',
#                   'Construction Budget',
#                   'Realigned funds',
#                   'Updated Construction Budget',
#                   'Contingency',
                #  'MER Variance',
                  'Award Amount',
                #  'Alternates',
                 # 'TI Allowance',
                 # 'TI Allowance Details',- get source from Ro/Alex G
                  'RFP Process',
                  'Score (Out of 5)',
                  '(Non-Pricing) Rank',
                  'Total Final Bid',
                  'Total Change Order',
                  'Justification',
                  'Environmental, Social & Governance Summary',
                  'Innovation Assessment',
                  'Bidder Source',
                  'Client',
                  'Decision to Bid',
                  'Total Scoring (out of 5)',
                  'Overall Rank' 
                   ]]
                 # 'Award Recommendation Summary' - rename Merged to this]]
                  #Date Entry was Made
                  #Entry Made by User]
                
    df_1 = df_1.rename(columns = {'Key Stakeholdres': 'Key Stakeholders', 
                              #'Merged' : 'Award Recommendation Summary',
                              'Presenation Date' : 'Presentation Date',
                              'Participating Bidders' : 'Vendor Name',
                              #'Recommendation' : 'Top Ranked Vendor',
                               'Diverse Supplier' : 'Diversity Certified'})
    df_1 = df_1.loc[:,~df_1.columns.duplicated()]            
    init.append(df_1)
print("")
print("Assembling final dataframe")
df_final = pd.concat(init)

#insert column to preserve legacy data
df_final.insert(0,'Date Entry was Made', '')
df_final.insert(0, 'Entry Made by User', '')
df_final.insert(5, 'Blank', '')
df_final.insert(6, 'CMS #', '')
df_final.insert(20, 'Project Description', '')
df_final.insert(21, 'Bidder Branch Address', '')
df_final.insert(22, 'Bidder Contact Name', '')
df_final.insert(23, 'Bidder Email', '')
df_final.insert(24, 'Bidder Phone', '')
df_final.insert(29, 'MER Budget Amount', '')
df_final.insert(30, 'Contingency', '')
df_final.insert(31, 'MER Variance', '')
df_final.insert(33, 'Alternates', '')
df_final.insert(34, 'TI Allowance', '')
df_final.insert(35, 'TI Allowance Details', '')
df_final.insert(14, 'Project State - City', '')

def recommend(row):
    if row['Vendor Name'].upper().strip() in row['Recommendation'].upper().strip():
        val = 'Yes'
    else:
        val = ''
    return val

#df_final.drop(['Recommendation'], axis = 1)
df_final['Recommendation (Y/N)'] = df_final.apply(recommend, axis = 1)

#final column ordering before insert
df_final = df_final.loc[:,[
'Vendor Name',
 'Presentation Date',
 'Sourcing Manager',
 'Requestor',
 'Key Stakeholders',
 'Blank',
 'CMS #',
 'Ariba WS',
 'CIIPS',
 'MER',
 'Project Type',
 'Project Address',
 'Project City',
 'Project State',
 'Project State - City',
 'Project Zip Code',
 'Project Region',
 'Notify Banking?',
 'Project Start Date',
 'Project End Date',
 'Contract Year',
 'Project Description',
 'Bidder Branch Address',
 'Bidder Contact Name',
 'Bidder Email',
 'Bidder Phone',
 'Bidder Source',
 'Diversity Certified',
 'Client',
 'Decision to Bid',
 'MER Budget Amount',
 'Contingency',
 'MER Variance',
 'Award Amount',
 'Alternates',
 'TI Allowance',
 'TI Allowance Details',
 'RFP Process',
 'Score (Out of 5)',
 '(Non-Pricing) Rank',
 'Total Final Bid',
 'Total Change Order',
 'Recommendation',   
 'Recommendation (Y/N)',
 'Justification',
'Date Entry was Made', 
'Entry Made by User',
'Environmental, Social & Governance Summary',
'Innovation Assessment',
'Total Scoring (out of 5)',
'Overall Rank'    
]]
print("")
print("compiling bid data and refreshing excel")



#remove duplicated columns from final df
df_final = df_final.loc[:,~df_final.columns.duplicated()].copy()

biddata = pd.concat([bid_master,df_final])

file = r"path to excel"

# from pandas import ExcelWriter

# with pd.ExcelWriter(
#              file,
#              mode="a",
#              engine = 'openpyxl',
#              if_sheet_exists = 'replace') as writer:
#      biddata.to_excel(writer, sheet_name = 'BidData', index = False)
    
import os

print('Moving files to Archive')

destination = r'archive folder path'

allfiles = os.listdir(path)

for f in allfiles:
    src_path = os.path.join(path, f)
    dst_path = os.path.join(destination, f)
    try:
        os.rename(src_path, dst_path)
    except:
        pass
    
print("")    
print("loading bid data and refreshing pivot tables")

import xlwings as xw

sheet_name = 'BidData'

#file = r"path to excel"
app = xw.App(visible = False)

wb = xw.Book(file)
ws = xw.sheets[sheet_name]

ws.range("A1").value = [biddata.columns.tolist()] + biddata.values.tolist()
wb.api.RefreshAll()
wb.save()
wb.close()
app.quit()
# import win32com.client

# office = win32com.client.Dispatch("Excel.Application")
# wb = office.Workbooks.Open(file)

# count = wb.Sheets.Count
# for i in range(count):
#     ws = wb.Worksheets[i]
#     pivotCount = ws.PivotTables().Count
#     for j in range(1, pivotCount+1):
#         ws.PivotTables(j).PivotCache().Refresh()

# wb.Save()        
# wb.Close()

print("")
print('Complete')

