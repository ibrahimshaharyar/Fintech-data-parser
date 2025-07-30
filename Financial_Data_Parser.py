#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import openpyxl
from IPython.display import display
import sqlite3
import re
from datetime import datetime
from decimal import Decimal
import locale


# ### Phase 1

# In[34]:


class ExcelProcessor:

    def __init__(self, file_paths):
        self.files = file_paths
        self.workbooks = {}

    def load_files(self):
        for file in self.files:
            excel_obj = pd.ExcelFile(file, engine='openpyxl')
            self.workbooks[file] = excel_obj
    
    def get_sheet_info(self):
        for file, excel_obj in self.workbooks.items():
            print(f"\nFile: {file}")
            for sheet in excel_obj.sheet_names:
                df = excel_obj.parse(sheet)
                print(f"Sheet: {sheet}")
                print(f"Shape: {df.shape}")
                print(f"Columns: {list(df.columns)}")
    
    def drop_all_null_columns(self, df):
        return df.dropna(axis=1, how='all')

    def preview_data(self, rows=5):
        for file, excel_obj in self.workbooks.items():
            print(f"\nPreview from: {file}")
            for sheet in excel_obj.sheet_names:
                df = excel_obj.parse(sheet)
                print(f"\nSheet: {sheet}")
                display(df.head(rows))
    


    


# In[35]:


files = ['KH_Bank.XLSX', 'Customer_Ledger_Entries_FULL.xlsx']

processor = ExcelProcessor(files)


# In[36]:


processor.load_files()


# In[5]:


processor.get_sheet_info()


# In[6]:


processor.preview_data()


# ### After adding dropping all null columns 

# In[28]:




# ### Phase 2

# In[37]:


from dateutil.parser import parse
import re


# In[38]:


class DataTypeDetector:

    def detect_column_type(self, series):
        data = series.dropna().astype(str)

        if data.empty:
            return 'Unknown'

        # Now i will classify dates

        date_success = 0

        for val in data[:100]:
            try:
                parse(val, fuzzy=False)
                date_success += 1
            
            except:
                continue
        
        if date_success / len(data[:100]) > 0.6:
            return 'Date'
    

        # Now i am classifying numbers

        num_success = 0

        for val in data[:100]:
            cleaned = re.sub(r'[^\d\.\-\(\)]', '', val)

            try:
                float(cleaned)
                num_success += 1
            
            except:
                continue
        
        if num_success / len(data[:100]) > 0.6:
            return 'Number'
        

        return 'String'


# In[39]:


detector = DataTypeDetector()

for file, excel_obj in processor.workbooks.items():
    print(f"\nFile: {file}")
    for sheet in excel_obj.sheet_names:
        df = excel_obj.parse(sheet)
        print(f"\nSheet: {sheet}")
        for col in df.columns:
            dtype = detector.detect_column_type(df[col])
            print(f"{col}: {dtype}")


# ### Phase 3

# In[40]:


import re
from dateutil import parser
from datetime import datetime, timedelta


# In[41]:


class FormatParser:
    
    def parse_amount(self, value):
        if pd.isnull(value):
            return None
        val = str(value).strip()

        # Handling paranthesis as negative

        if val.startswith('(') and val.endswith(')'):
            val = '-' + val[1:-1]

        #removing currency symbols and letters

        val = re.sub(r'[^\d.,\-KMBkmb]', "", val)

        multiplier = 1
        if val and val[-1].lower() == 'k':
            multiplier = 1e3
            val = val[:-1]

        elif val and val[-1].lower() == 'm':
            multiplier = 1e6
            val = val[:-1]
        
        elif val and val[-1].lower() == 'b':
            multiplier = 1e9
            val = val[:-1]

        
        if ',' in val and val.count(",") > val.count('.'):
            val = val.replace(',', "").replace(',', '.')

        # formatting indian way 
        val = val.replace(',', '')

        try:
            return float(val) * multiplier
        except:
            return None

        # handling excel serial date
    def parse_date(self, value):
        if pd.isnull(value):
            return None
        val = str(value).strip()

        if val.isdigit() and len(val) <= 5:
            try:
                return datetime(1899,12,30) + timedelta(days= int(val))
            
            except:
                pass

        # quarter formats
        q_match = re.match(r'Q([1-4])[-\s]?\s*(\d{2,4})', val, re.IGNORECASE)

        if q_match:
            q = int(q_match.group(1))
            year = int(q_match.group(2))
            if year < 100:
                year += 2000
            
            return datetime(year, 3 * q - 2, 1)
        
        try:
            return parser.parse(val, fuzzy = True)
        
        except:
            return None


# In[42]:


fp = FormatParser()


# In[43]:


df = processor.workbooks['KH_Bank.XLSX'].parse('Sheet1')


# In[44]:


# Identify amount and date columns (based on Phase 2 output)
amount_cols = [col for col in df.columns if 'amount' in col.lower()]
date_cols = [col for col in df.columns if 'date' in col.lower()]


# In[45]:


for col in amount_cols:
    df[col + "_parsed"] = df[col].apply(fp.parse_amount)

for col in date_cols:
    df[col + "_parsed"] = df[col].apply(fp.parse_date)


# In[46]:


display(df[[*amount_cols, *[c+"_parsed" for c in amount_cols]]].head())
display(df[[*date_cols, *[c+"_parsed" for c in date_cols]]].head())


# ### Phase 4

# In[47]:


class FinancialDataStore:
    def __init__(self):
        self.data = {}  # Stores DataFrames

    def add_dataset(self, name, df):
        self.data[name] = df

    def query_range(self, name, column, start, end):
        df = self.data.get(name)
        if df is None or column not in df.columns:
            return None
        return df[(df[column] >= start) & (df[column] <= end)]

    def aggregate(self, name, group_by_col, value_col):
        df = self.data.get(name)
        if df is None: return None
        return df.groupby(group_by_col)[value_col].sum().reset_index()


# In[49]:


df = processor.workbooks['KH_Bank.XLSX'].parse('Sheet1')
df = processor.drop_all_null_columns(df)


# In[50]:


for col in amount_cols:
    df[col + "_parsed"] = df[col].apply(fp.parse_amount)

for col in date_cols:
    df[col + "_parsed"] = df[col].apply(fp.parse_date)


# In[51]:


store = FinancialDataStore()
store.add_dataset("KH_Bank", df)


# In[52]:


result = store.query_range("KH_Bank", "Statement.Entry.ValueDate.Date_parsed", "2023-01-01", "2023-12-31")
display(result.head())


# In[53]:


summary = store.aggregate("KH_Bank", "Statement.Entry.Amount.Currency", "Statement.Entry.Amount.Value_parsed")
display(summary.head())


# In[54]:


from datetime import datetime

result = store.query_range(
    "KH_Bank",
    "Statement.Entry.ValueDate.Date_parsed",
    datetime(2023, 1, 1),
    datetime(2023, 12, 31)
)
display(result.head())







