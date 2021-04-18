import pandas as pd
import datetime as dt
import calendar
import plotly
import plotly.express as px

#Files list 
report_files = ['Jan.xlsx', 'Feb.xlsx', 'Mar.xlsx']

# Merged 
merged = pd.DataFrame()

# The Loop for merging 
for file in report_files:
    df = pd.read_excel(file)
    df['Date'] = df['Date'].dt.date
    df['Day'] = pd.DatetimeIndex(df['Date']).day
    df['Month'] = pd.DatetimeIndex(df['Date']).month
    df['Year'] = pd.DatetimeIndex(df['Date']).year
    merged= merged.append(df,ignore_index = True)

# export to excel 
merged.to_excel('Sales_Q2020.xlsx', index = False, sheet_name='Quarter Sales Report')

