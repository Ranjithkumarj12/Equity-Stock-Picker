# -*- coding: utf-8 -*-
"""
Created on Tue Nov  8 17:09:28 2022

@author: jrkumar
"""

import pandas as pd
import os
import numpy as np
from datetime import date

#Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

#Creating empty list to store list of tickers
files = []

#List of Ratios used
ratios_list = ['Gross Profit Margin',	'Operating Profit Margin',	'Pre-tax Margin',	'Net Profit Margin',	
               'Return on Assets (Added Gr.Interest)',	'Operating Return on Assets',	'Return on Total Capital',	'Return on Equity',		
               'Inventory Turnover',	'Days of Inventory on hand', 'Total Asset Turnover',	'Fixed Asset Turnover',	'Working Capital Turnover',	'Current Ratio',	
               	'Cash Ratio',	'Debt-to-Equity',	'Debt-to-Capital Ratio',	'Debt-to-Assets',	'Financial Leverage',	'Interest Coverage (Income)',	
               'CFO-to-Net Revenue',	'CFO-to-Assets',	'CFO-to-Equity',	'CFO-to-Op.Income',	'CFO-to-Debt',	'Interest Coverage (CFO)', 'Earnings to Price',
               'Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

ratios_list_descending = ['Gross Profit Margin',	'Operating Profit Margin',	'Pre-tax Margin',	'Net Profit Margin',	
                          'Return on Assets (Added Gr.Interest)',	'Operating Return on Assets','Return on Total Capital',	'Return on Equity',
                          'Inventory Turnover','Total Asset Turnover',	'Fixed Asset Turnover',	'Working Capital Turnover',	'Current Ratio',	
                          'Cash Ratio','Financial Leverage','Interest Coverage (Income)',	
                          'CFO-to-Net Revenue',	'CFO-to-Assets',	'CFO-to-Equity',	'CFO-to-Op.Income',	'CFO-to-Debt',	'Interest Coverage (CFO)']

ratios_list_ascending = ['Days of Inventory on hand',
                         'Debt-to-Equity',	'Debt-to-Capital Ratio',	'Debt-to-Assets']

ratios_list_valuation = ['Earnings to Price',	'Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

ratios_not_using = ['Receivables Turnover',	'Days of Sales Outstanding','Payables Turnover',	'Days of Sales Payables','Quick Ratio',
                    'Dividend Payment Coverage (CFO)','Outflows to CFI & CFF (CFO)','Dividend to Price',	'EBITDA to EV',
                    'Defensive Interval',	'Cash Conversion Cycle']


# Getting the list of tickers in the Analysis Part 1 folder    
for file1 in os.listdir(output_path1):
    f_name1 = file1.replace(" - Analysis Part 1.xlsx","")
    files.append(f_name1)

# Computing Valuation Ratios (price as of one month prior to reporting date) and Updating it to Analysis Part 1
for t in files:
    analysis_p1_val_df = pd.read_excel(output_path1+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
    try:
        analysis_p1_val_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
    except:
        pass
    analysis_p1_val_df['Price Date'] = date(2022,10,31)
    analysis_p1_val_df['Price'] = 0.00
    analysis_p1_val_df['Earnings to Price'] = 0.00
    analysis_p1_val_df['Sales to Price'] = 0.00
    analysis_p1_val_df['CFO to Price'] = 0.00
    analysis_p1_val_df['BV to Price'] = 0.00
    analysis_p1_val_df['FCFF to Price'] = 0.00
    for d in range(len(analysis_p1_val_df['Rep_Date'])):
        # Converting string date to date format
        price_date = analysis_p1_val_df['Rep_Date'][d]
        d_year = price_date[-4:]
        d_year_int = int(d_year)
        d_month = price_date[:2]
        if d_month[1] == '/':
           d_month = '0'+d_month[0]
           d_month_int = int(d_month)
        d_day = price_date[-7:-5]
        d_day_int = int(d_day)
        price_date = date(d_year_int, d_month_int, d_day_int)
        # Considering Price one month prior to Reporting date (to ensure that actual performance is not priced in) 
        price_date.replace(day=(price_date.month -1))
        y = price_date.isoweekday()
        if y == 6:
            price_date = price_date.replace(day=(price_date.day -1))
        elif y == 7:
            price_date = price_date.replace(day=(price_date.day -2))
        
        # Fetching Price data for Rep_Dates
        analysis_p1_val_df['Price Date'][d] = price_date
        analysis_price_df = pd.read_excel(price_path+r'{} - Daily Price data.xlsx'.format(t), sheet_name=0)
        try:
            analysis_price_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
        except:
            pass
        # Converting 'Date' in string format to date time format
        # Assigning a temporary date
        analysis_price_df['Date New'] = date(2022,10,31)
        for j in range(len(analysis_price_df['Date'])):
            d_year2 = analysis_price_df['Date'].iloc[j][:4]
            d_year2 = int(d_year2)
            d_month2 = analysis_price_df['Date'].iloc[j][5:7]
            d_month2 = int(d_month2)
            d_day2 = analysis_price_df['Date'].iloc[j][-2:]
            d_day2 = int(d_day2)
            d_date2 = date(d_year2, d_month2, d_day2)
            analysis_price_df['Date New'].iloc[j] = d_date2
        analysis_price_df.set_index('Date New',inplace = True, drop = True)
        try:
            analysis_p1_val_df['Price'][d] = analysis_price_df['Adj Close'].loc[price_date]
            analysis_p1_val_df['Price Date'][d] = price_date
        except:
            try:
                price_date =  price_date.replace(day=(price_date.day -1))
                analysis_p1_val_df['Price'][d] = analysis_price_df['Adj Close'].loc[price_date]
                analysis_p1_val_df['Price Date'][d] = price_date
            except:
                try:
                    price_date =  price_date.replace(day=(price_date.day -2))
                    analysis_p1_val_df['Price'][d] = analysis_price_df['Adj Close'].loc[price_date]
                    analysis_p1_val_df['Price Date'][d] = price_date
                except:
                    try:
                        price_date =  price_date.replace(day=(price_date.day -3))
                        analysis_p1_val_df['Price'][d] = analysis_price_df['Adj Close'].loc[price_date]
                        analysis_p1_val_df['Price Date'][d] = price_date
                    except:
                        analysis_p1_val_df['Price'][d] = analysis_price_df['Adj Close'].iloc[0]
                        analysis_price_df.reset_index(inplace = True, drop = False)
                        analysis_p1_val_df['Price Date'][d] = analysis_price_df['Date New'].iloc[0]
                        analysis_price_df.set_index('Date New',inplace = True, drop = True)
        try:
            analysis_p1_val_df['Earnings to Price'][d] = ((analysis_p1_val_df['Net Income'][d]/analysis_p1_val_df['Ordinary Shares Number'][d])/analysis_p1_val_df['Price'][d])
        except:
            analysis_p1_val_df['Earnings to Price'][d] = np.NaN
        try:
            analysis_p1_val_df['Sales to Price'][d] = ((analysis_p1_val_df['Total Revenue'][d]/analysis_p1_val_df['Ordinary Shares Number'][d])/analysis_p1_val_df['Price'][d])
        except:
            analysis_p1_val_df['Sales to Price'][d] = np.NaN
        try:
            analysis_p1_val_df['CFO to Price'][d] = ((analysis_p1_val_df['Operating Cash Flow'][d]/analysis_p1_val_df['Ordinary Shares Number'][d])/analysis_p1_val_df['Price'][d])
        except:
            analysis_p1_val_df['CFO to Price'][d] = np.NaN
        try:
            analysis_p1_val_df['BV to Price'][d] = ((analysis_p1_val_df['Total Equity Gross Minority Interest'][d]/analysis_p1_val_df['Ordinary Shares Number'][d])/analysis_p1_val_df['Price'][d])
        except:
            analysis_p1_val_df['BV to Price'][d] = np.NaN
        try:
            analysis_p1_val_df['FCFF to Price'][d] = ((analysis_p1_val_df['Free Cash Flow'][d]/analysis_p1_val_df['Ordinary Shares Number'][d])/analysis_p1_val_df['Price'][d])
        except:
            analysis_p1_val_df['FCFF to Price'][d] = np.NaN
    print('Analysis Part 1 for Ticker: {} is Updated'.format(t))
    
    # Saving updated Analysis Part 1 to excel
    writer = pd.ExcelWriter(output_path1+'{} - Analysis Part 1.xlsx'.format(t), engine='xlsxwriter')
    analysis_p1_val_df.to_excel(writer)
    writer.save()
    
# Now move to Industry Stats