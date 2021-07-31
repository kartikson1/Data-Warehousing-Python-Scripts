import pandas as pd
import numpy as np
import os
import sys
import csv
import glob

#location of folder containing subfolders and excel files in quotes
folder = r'/content/drive/MyDrive/Test files 2'

# Get all excel files from all subfolders into variable 'f
paths = []
for root, dirs, files in os.walk("/content/drive/MyDrive/Test files 2"):
    for file in files:
        if file.endswith(".xlsx"):
             #print(os.path.join(root, file))
             s = os.path.join(root, file)
             print(s)
             paths.append(s)

all_data = pd.DataFrame()
for f in paths:
    # df = pd.read_excel(f)
    # all_data = all_data.append(df,ignore_index=True)
    print(f)

j= 'a'
k='b'
for file in paths:

  dfs = pd.read_excel(os.path.join(folder,file), sheet_name = None)
  for df in dfs.values():
    i=0
    if 'order_address' in df:
      df['Pincode'] = ""
      for address in df['order_address']:
        addr = str(address)
        if addr[-1].isdigit():
          df['Pincode'][i] = address[-6:]
          i=i+1
        else:
          df['Pincode'][i] = ''
          i=i+1
      i=0
      for address in df['order_address']:
        addr = str(address)
        address = addr[0:-7]
        df['order_address'][i] = address
        i = i+1
      df[['AddressLine', 'City', 'State']] = df['order_address'].str.rsplit(',', 2, expand=True)
      
      # print('if stmt')
      df.to_excel(j+ 'output.xlsx', sheet_name=j)
      j = j + 'a'
    
    else:
      # print('else stmt')
      # print(df)
      df.to_excel(k+ 'output.xlsx', sheet_name=k)
      j = j + 'a'
      k = k + 'b'

