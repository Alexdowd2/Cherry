#!/usr/bin/env python
# coding: utf-8




#import libraries
import pandas as pd
from pathlib import Path
import glob

#establish empty list
init_df = []

#for loop to iterate through files in folder, read into pandas, then append to list
for i in glob.glob('filePath\\*.xlsx'):
    this_df = pd.read_excel(i, sheet_name = 'mySheet')
    init_df.append(this_df)

#create a new dataframe from the init_df list
master_df = pd.concat(init_df)

#write the file to excel
master_df.to_excel('filePath', index = False)



