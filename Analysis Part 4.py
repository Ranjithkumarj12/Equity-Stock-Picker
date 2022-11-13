# -*- coding: utf-8 -*-
"""
Created on Sun Nov  6 21:20:05 2022

@author: jrkumar
"""

import pandas as pd
import numpy as np
from datetime import date

# Stored Path
input_path1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path4 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 4 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Industry Stats\\')

# List of Ratios used for Valuation
ratios_list_valuation = ['Earnings to Price', 'Sales to Price',	'CFO to Price',	'BV to Price', 'FCFF to Price']

# Reading final list of tickers that qualified for further processing
zscore_ticker_df = pd.read_excel(output_path3+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
zscore_ticker_df.rename(columns={'Unnamed: 0':'Tickers'}, inplace=True)
zscore_ticker_list = zscore_ticker_df['Tickers'].dropna().values.tolist()

# Price Return Part
# Creating empty data frame to store Price CAGR
price_return_df = pd.DataFrame(index = zscore_ticker_list, columns = ['Price Return'])

# Fetching Price data for final list of tickers
for t in zscore_ticker_list:
    price_df = pd.read_excel(price_path+r'{} - Daily Price data.xlsx'.format(t), sheet_name=0)
    price_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
    # Converting 'Date' in string format to date time format
    # Assigning a temporary date
    price_df['Date New'] = date(2022,10,31)
    for j in range(len(price_df['Date'])):
        d_year2 = price_df['Date'].iloc[j][:4]
        d_year2 = int(d_year2)
        d_month2 = price_df['Date'].iloc[j][5:7]
        d_month2 = int(d_month2)
        d_day2 = price_df['Date'].iloc[j][-2:]
        d_day2 = int(d_day2)
        d_date2 = date(d_year2, d_month2, d_day2)
        price_df['Date New'].iloc[j] = d_date2
    price_df.set_index('Date New',inplace = True, drop = True)

    # Reading analysis part 1 to get Rep_date
    analysis_p1_df = pd.read_excel(output_path1+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
    analysis_p1_df = analysis_p1_df['Rep_Date'] 
    
    # Assigning Start Date
    rep_count = len(analysis_p1_df)
    start_date = analysis_p1_df[rep_count - 1]
    
    # Converting string date to date format
    d_year = start_date[-4:]
    d_year_int = int(d_year)
    d_month = start_date[:2]
    if d_month[1] == '/':
       d_month = '0'+d_month[0]
       d_month_int = int(d_month)
    d_day = start_date[-7:-5]
    d_day_int = int(d_day)
    start_date = date(d_year_int, d_month_int, d_day_int)
    y = start_date.isoweekday()
    if y == 6:
        start_date = start_date.replace(day=(start_date.day -1))
    elif y == 7:
        start_date = start_date.replace(day=(start_date.day -2))
    
    # Assigning End Date
    end_date = date(2022,10,31)
    z = end_date.isoweekday()
    if z == 6:
        end_date = end_date.replace(day=(end_date.day -1))
    elif z == 7:
        end_date = end_date.replace(day=(end_date.day -2))
        
    price_df2 = price_df.loc[start_date:end_date]
    price_df3 = price_df2['Adj Close']
    
    # Calculating Price Return (Simple Return)
    fin_price = price_df3[-1]
    beg_price = price_df3[0]
    price_return = (fin_price/beg_price)-1
    print('Price Return for Ticker: {} is complete'.format(t))
    
    # Storing Price Return to dataframe
    price_return_df['Price Return'].loc[t] = price_return
    
# Price Return Ranking
for ratio in range(len(price_return_df)):
    price_return_df['Price Return - Ranking'] = price_return_df['Price Return'].rank(ascending = False, method = 'average', na_option='bottom')
    
# Saving Price Return Dataframe to excel
writer = pd.ExcelWriter(output_path4+'All Ticker Price Return.xlsx', engine='xlsxwriter')
price_return_df.to_excel(writer)
writer.save()

print('Price Return for all tickers is complete and saved as excel')
        
# Combining Price Return Rank with Z-Score Combined Rank
analysis_p4_df = pd.read_excel(output_path3+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
analysis_p4_df.rename(columns={'Unnamed: 0':'Tickers'}, inplace=True)
analysis_p4_df.set_index('Tickers',inplace = True, drop = True)

analysis_p4_df = pd.concat([analysis_p4_df, price_return_df], axis=1)

# 70% weightage to Price momentum, 30% weightage to Fundamental Z-Score
wt_price_return = 0.7
wt_zscores = 0.3
analysis_p4_df['Combined Score (with Price Return)'] = (wt_price_return * analysis_p4_df['Price Return - Ranking']) + (wt_zscores * analysis_p4_df['Combined Rank'])
analysis_p4_df['Combined Ranking (with Price Return)'] = analysis_p4_df['Combined Score (with Price Return)'].rank(ascending = True, method = 'average', na_option='bottom')

# Saving Updated Analysis Dataframe (Part 3) to excel
writer = pd.ExcelWriter(output_path4+'Z-Score and Price Return Ranking.xlsx', engine='xlsxwriter')
analysis_p4_df.to_excel(writer)
writer.save()

print('Z-Score and Price Return Ranking for all Tickers is complete and saved as excel')

# Now move to Analysis Part 5


        





    

