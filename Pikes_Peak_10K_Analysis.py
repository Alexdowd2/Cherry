#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# this code will create a combined document of males and females with gender idenfitying columns.
# the remainder of the analysis will be in excel

import pandas as pd
import numpy as np


# the file is a tab-delimited txt file encoded with ISO not UTF-8
# so we must specify the encoding for pandas to read it correctly

males = pd.read_csv(r'PikesPeak_Males.txt',
                   sep = '\t',
                   encoding = 'ISO-8859-1')
females = pd.read_csv(r'PikesPeak_Females.txt',
                     sep = '\t',
                     encoding = 'ISO-8859-1')


# In[ ]:


# create the male gender column

males['Gender'] = 'M'
males.head()


# In[ ]:


# create the female gender column
females['Gender'] = 'F'
females.head()


# In[ ]:


#combine the dataframes into one dataframe

df = pd.concat([males, females])


# In[ ]:


# create the division colum

def division(row):
    global val
    if row['Ag'] <= 14:
        val = '0 - 14'
    elif row['Ag'] >= 15 and row['Ag'] <= 19:
        val = '15 - 19'
    elif row['Ag'] >= 20 and row['Ag'] <= 29:
        val = '20 - 29'
    elif row['Ag'] >= 30 and row['Ag'] <= 39:
        val = '30 - 39'
    elif row['Ag'] >= 40 and row['Ag'] <= 49:
        val = '40 - 49'
    elif row['Ag'] >= 50 and row['Ag'] <= 59:
        val = '50 - 59'
    elif row['Ag'] >= 60 and row['Ag'] <= 69:
        val = '60 - 69'
    elif row['Ag'] >= 70 and row['Ag'] <= 79:
        val = '70 - 79'
    elif row['Ag'] >= 80 and row['Ag'] <= 89:
        val = '80 -89'
    return val

df['Division'] = df.apply(division, axis = 1)
df.head()


# In[ ]:


#remove spaces in column names and the '/' from the Div/Tot column

df.columns = df.columns.str.replace(" ","")
df.columns = df.columns.str.replace("/","")


# In[ ]:


#split the Div/Tot column into 2 new columns (Div Place and Total in Div) then drop the Div/Tot column

df[['Div Place','Total in Div']] = df.DivTot.str.split("/", expand = True)
df = df.drop(['DivTot'], axis = 1)


# In[ ]:


# remove the'#' from the NetTim column

df['NetTim'] = df.NetTim.str.replace('#',"")
df.head()


# In[ ]:


#remove periods from Hometown column

df['Hometown'] = df.Hometown.str.replace(".","")
df.head()


# In[ ]:


# clean the issue in the GunTim column of state initials being included

#first, create 2 columns based on a space

df[['initial','time']] = df['GunTim'].str.split(" ", n = 1, expand = True)
df.tail(10)


# In[ ]:


# Create a function to only concatenate the letters to Hometown

def htletters(a):
    global name
    if a['initial'] in 'A' or a['initial'] in 'D' or a['initial'] in 'M' or a['initial'] in 'N' or a['initial'] in 'V':
        return a['Hometown'] + a['initial']
    else:
        return a['Hometown']

df['HT2'] = df.apply(htletters, axis = 1)
df.tail(10)


# In[ ]:


# create consolidated guntim column

def guntim(b):
    if b['initial'] in 'A' or b['initial'] in 'D' or b['initial'] in 'M' or b['initial'] in 'N' or b['initial'] in 'V':
        col = b['time']
    else:
        col = b['initial']
        
    return col

df['GunTim2'] = df.apply(guntim, axis = 1)

df.tail()


# In[ ]:


# drop the unnecessary 'helper' columns

df = df.drop(['Hometown','initial','time', 'GunTim'], axis =1)
df = df.rename(columns={'GunTim2': 'GunTim'})
df.tail()


# In[ ]:


# rename the HT2 column
df = df.rename(columns={'HT2': 'Hometown'})
df.tail()


# In[ ]:


# remove the * from NetTim column

df['NetTim'] = df.NetTim.str.replace('*',"")
df.head(10)


# In[ ]:


# write the df to a csv 

df.to_csv('new_file.csv', index = False) 


# In[ ]:




