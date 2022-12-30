# -*- coding: utf-8 -*-
"""
Created on Fri Nov 11 13:15:05 2022

@author: jrkumar
"""
# Identifying tickers which have consistent quarterly results data

import pandas as pd
from datetime import date
import numpy as np
import os
import fnmatch

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path6 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 6 (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

# Reading final list of tickers that qualified for further processing
files = []
for file in os.listdir(input_path1):    
    if fnmatch.fnmatch(file, '* - Quarter data (Income Statement).xlsx'):
        f_name = file.replace(" - Quarter data (Income Statement).xlsx","")
        files.append(f_name)

# For Rolling Quarters
# We want a minimum of 2 quarters and a maximum of 3 quarters of data to calculate CAGR
# Creating a list to store tickers which have recent 3 quarters of data
full_data = []
# Creating a list to store tickers which have either the recent two quarters of data or have q-2 and q-1 data without having the most recent quarters data
partial_data = []
# Creating a list to store tickers which have dont satify the above two conditions
no_data = []

todays_date = date.today()

if (todays_date.month > 3) & (todays_date.month <= 6):
    qtr_month = 3
    qtr_year = todays_date.year
    qtr_minus1_month = 12
    qtr_minus1_year = todays_date.year-1
    qtr_minus2_month = 9
    qtr_minus2_year = todays_date.year-1
    
elif (todays_date.month > 6) & (todays_date.month <= 9):
    qtr_month= 6
    qtr_year = todays_date.year
    qtr_minus1_month = 3
    qtr_minus1_year = todays_date.year
    qtr_minus2_month = 12
    qtr_minus2_year = todays_date.year-1
    
elif (todays_date.month > 9) & (todays_date.month <= 12):
    qtr_month = 9
    qtr_year = todays_date.year
    qtr_minus1_month = 6
    qtr_minus1_year = todays_date.year
    qtr_minus2_month = 3
    qtr_minus2_year = todays_date.year
else:
    qtr_month = 12
    qtr_year = todays_date.year-1
    qtr_minus1_month = 9
    qtr_minus1_year = todays_date.year-1
    qtr_minus2_month = 6
    qtr_minus2_year = todays_date.year-1

# Fetching Price data for final list of tickers
for t in files:
    inc_statement = pd.read_excel(input_path1+r'{} - Quarter data (Income Statement).xlsx'.format(t), sheet_name=0)
    inc_statement.rename(columns={'Unnamed: 0':'Rep_Date'}, inplace=True)
    inc_statement['Rep_date New'] = inc_statement['Rep_Date']
    
    # For Rolling Quarters 
    # Creating new dataframe for rolling quarters
    inc_statement_copy = inc_statement.copy()
    
    # First row
    if len(inc_statement_copy['Rep_date New']) >= 3:
        if len(inc_statement_copy['Rep_date New'].iloc[1]) == 9:
            first_row_month = int(inc_statement_copy['Rep_date New'].iloc[1][0])
            first_row_year = int(inc_statement_copy['Rep_date New'].iloc[1][-4:])
        elif len(inc_statement_copy['Rep_date New'].iloc[1]) == 10:
            first_row_month = int(inc_statement_copy['Rep_date New'].iloc[1][0:1])
            first_row_year = int(inc_statement_copy['Rep_date New'].iloc[1][-4:])
        
        # Second row
        if len(inc_statement_copy['Rep_date New'].iloc[2]) == 9:
            second_row_month = int(inc_statement_copy['Rep_date New'].iloc[2][0])
            second_row_year = int(inc_statement_copy['Rep_date New'].iloc[2][-4:])
        elif len(inc_statement_copy['Rep_date New'].iloc[2]) == 10:
            second_row_month = int(inc_statement_copy['Rep_date New'].iloc[2][0:1])
            second_row_year = int(inc_statement_copy['Rep_date New'].iloc[2][-4:])
        
        # Third row if available
        if len(inc_statement_copy['Rep_date New']) > 3:
            # Third row
            if len(inc_statement_copy['Rep_date New'].iloc[3]) == 9:
                third_row_month = int(inc_statement_copy['Rep_date New'].iloc[3][0])
                third_row_year = int(inc_statement_copy['Rep_date New'].iloc[3][-4:])
            elif len(inc_statement_copy['Rep_date New'].iloc[3]) == 10:
                third_row_month = int(inc_statement_copy['Rep_date New'].iloc[3][0:1])
                third_row_year = int(inc_statement_copy['Rep_date New'].iloc[3][-4:])  
            
            # Stocks with last 3 quarters data
            if (((first_row_month == qtr_month) & (first_row_year == qtr_year)) & ((second_row_month == qtr_minus1_month) & (second_row_year == qtr_minus1_year)) & ((third_row_month == qtr_minus2_month) & (third_row_year == qtr_minus2_year))):
                full_data.append(t)
            else:
                if ((((first_row_month == qtr_month) & (first_row_year == qtr_year)) & ((second_row_month == qtr_minus1_month) & (second_row_year == qtr_minus1_year))) | (((first_row_month == qtr_minus1_month) & (first_row_year == qtr_minus1_year)) & ((second_row_month == qtr_minus2_month) & (second_row_year == qtr_minus2_year)))):
                    partial_data.append(t)
                else:
                    no_data.append(t)
        # Stock with either the recent two quarters of data or have q-2 and q-1 data without having the most recent quarters data
        if len(inc_statement_copy['Rep_date New']) <= 3:
            if ((((first_row_month == qtr_month) & (first_row_year == qtr_year)) & ((second_row_month == qtr_minus1_month) & (second_row_year == qtr_minus1_year))) | (((first_row_month == qtr_minus1_month) & (first_row_year == qtr_minus1_year)) & ((second_row_month == qtr_minus2_month) & (second_row_year == qtr_minus2_year)))):
                partial_data.append(t)
            else:
                no_data.append(t)
    else:
         no_data.append(t)
    
    
    print('Ticker: {} is complete'.format(t))
    
full_data_df = pd.DataFrame(full_data, columns = ['Tickers with Full data'])
partial_data_df = pd.DataFrame(partial_data, columns = ['Tickers with Partial data'])    
no_data_df = pd.DataFrame(no_data, columns = ['Tickers with No data'])    

# Saving full_data_df to excel
writer = pd.ExcelWriter(output_path6+'Full data.xlsx', engine='xlsxwriter')
full_data_df.to_excel(writer)
writer.save()

# Saving partial_data_df to excel
writer = pd.ExcelWriter(output_path6+'Partial data.xlsx', engine='xlsxwriter')
partial_data_df.to_excel(writer)
writer.save()

# Saving no_data_df to excel
writer = pd.ExcelWriter(output_path6+'No data.xlsx', engine='xlsxwriter')
no_data_df.to_excel(writer)
writer.save()
    
# For Quarter-on-Quarter
qoq_full_data = []
qoq_no_data = []
qoq_ticker_data = full_data + partial_data

for t in qoq_ticker_data:
    inc_statement = pd.read_excel(input_path1+r'{} - Quarter data (Income Statement).xlsx'.format(t), sheet_name=0)
    inc_statement.rename(columns={'Unnamed: 0':'Rep_Date'}, inplace=True)
    inc_statement['Rep_date New'] = inc_statement['Rep_Date']
    
    # Creating new dataframe for rolling quarters
    inc_statement_copy2 = inc_statement.copy()
    
    # First row
    if len(inc_statement_copy2['Rep_date New'].iloc[1]) == 9:
        first_row_month_qoq = int(inc_statement_copy2['Rep_date New'].iloc[1][0])
        first_row_year_qoq =  int(inc_statement_copy2['Rep_date New'].iloc[1][-4:])
    elif len(inc_statement_copy2['Rep_date New'].iloc[1]) == 10:
        first_row_month_qoq =  int(inc_statement_copy2['Rep_date New'].iloc[1][0:1])
        first_row_year_qoq =  int(inc_statement_copy2['Rep_date New'].iloc[1][-4:])
    
    if ((first_row_month_qoq == qtr_month) & (first_row_year_qoq == qtr_year)):
        for j in (1,len(inc_statement_copy2['Rep_date New'])-1,1):
            if len(inc_statement_copy2['Rep_date New'].iloc[j]) == 9:
                last_year_qtr_month =  int(inc_statement_copy2['Rep_date New'].iloc[j][0])
                last_year_qtr_year =  int(inc_statement_copy2['Rep_date New'].iloc[j][-4:])
            elif len(inc_statement_copy2['Rep_date New'].iloc[j]) == 10:
                last_year_qtr_month =  int(inc_statement_copy2['Rep_date New'].iloc[j][0:1])
                last_year_qtr_year =  int(inc_statement_copy2['Rep_date New'].iloc[j][-4:])
            
            if ((last_year_qtr_month == first_row_month_qoq) & (last_year_qtr_year == first_row_year_qoq-1)):
                qoq_full_data.append(t)
            else:
                if (len(inc_statement_copy2['Rep_date New']) - 1) == j:
                    qoq_no_data.append(t)
                else:
                    pass
    else:
        qoq_no_data.append(t)
            
    print('Ticker: {} is complete'.format(t))
    
qoq_full_data_df = pd.DataFrame(qoq_full_data, columns = ['Tickers with QoQ Full data'])  

# Saving qoq_full_data_df to excel
writer = pd.ExcelWriter(output_path6+'QoQ Full data.xlsx', engine='xlsxwriter')
qoq_full_data_df.to_excel(writer)
writer.save()

# Now move to Analysis 7
