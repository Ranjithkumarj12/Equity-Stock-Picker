# -*- coding: utf-8 -*-
"""
Created on Tue Oct 11 18:41:54 2022

@author: jrkumar
"""

#Importing Libraries
import pandas as pd
import numpy as np
import os

#Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
ticker_input = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

#Creating empty dictionary to store stock analysis
financial_dir = {}

#Input Tickers
tickers_df = pd.read_excel(ticker_input+r'Total Ticker List - Annual.xlsx', sheet_name=0)
#Final Consideration List
final_tickers1 = tickers_df['Final Consideration List - 2'].dropna().values.tolist()
final_tickers1 = final_tickers1
#Temp - To see how many years of data are available across securities - Found out that 4 yrs of data are avl. for majority of the securities.
ind = ['row_len','col_len']
length_df = pd.DataFrame(index = ind)
for file in os.listdir(input_path1):
    t = pd.read_excel(input_path1+file, sheet_name=0)
    file = file.replace(" - Annual data.xlsx","")
    row_len = len(t)
    col_len = len(t.columns)
    length_df[file] = np.NaN
    length_df[file].loc['row_len'] = row_len
    length_df[file].loc['col_len'] = col_len
length_df = length_df.T

length_df2 = length_df.loc[(length_df['row_len'] == 4) | (length_df['row_len'] == 3)]
final_tickers2 = length_df2.index.dropna().values.tolist()

final_tickers3 = []
for i in final_tickers1:
    if i in final_tickers2:
        final_tickers3.append(i)
        
#From above analysis, we decide that we will only analyse those securties where we have a minimum of 3 yrs of FS data
for t in final_tickers3:
    
    finy1 = pd.read_excel(input_path1+r'{} - Annual data.xlsx'.format(t), sheet_name=0)
    finy1.rename(columns={'Unnamed: 0':'Year'}, inplace=True)
    finy1.set_index('Year', inplace = True)
    # Converting from "in thousands" to Actual values
    finy1[finy1.select_dtypes(include=['number']).columns] = finy1[finy1.select_dtypes(include=['number']).columns] * 1000
    # Reversing the order of Rows/Years
    finy1 = finy1.loc[::-1]
    fin_2 = finy1.copy()
    
    sumy1 = pd.read_excel(input_path_s1+r'{} - Stats & Price data.xlsx'.format(t), sheet_name=0)
    sumy1.rename(columns={'Unnamed: 0':'Year'}, inplace=True)
    sumy1.set_index('Year', inplace = True)
    sum1 = sumy1.copy()
    
# Profitability Ratios
    for i in range(len(fin_2)):
        # Gross Profit Margin
        try:
            fin_2['Gross Profit Margin'].iloc[i] = fin_2['Gross Profit'].iloc[i]/fin_2['Total Revenue'].iloc[i]
            num = fin_2['Gross Profit'].iloc[i]
            den = fin_2['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Gross Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                fin_2['Gross Profit Margin'].iloc[i] = (fin_2['Total Revenue'].iloc[i] - fin_2['Cost of Revenue'].iloc[i])/fin_2['Total Revenue'].iloc[i]
                num = (fin_2['Total Revenue'].iloc[i] - fin_2['Cost of Revenue'].iloc[i])
                den = fin_2['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    fin_2['Gross Profit Margin'].iloc[i] = np.NaN
            except:
                fin_2['Gross Profit Margin'].iloc[i] = np.NaN
        
        #fin_2['EBIT'] = fin_2['EBITDA'] - fin_2['Depreciation and Amortization']
        
        # Operating Profit Margin
        try:
            fin_2['Operating Profit Margin'].iloc[i] = fin_2['EBIT'].iloc[i]/fin_2['Total Revenue'].iloc[i]
            num = fin_2['EBIT'].iloc[i]
            den = fin_2['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Operating Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                fin_2['Operating Profit Margin'].iloc[i] = fin_2['Operating Income'].iloc[i]/fin_2['Total Revenue'].iloc[i]
                num = fin_2['Operating Income'].iloc[i]
                den = fin_2['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    fin_2['Operating Profit Margin'].iloc[i] = np.NaN
            except:
                fin_2['Operating Profit Margin'].iloc[i] = np.NaN
        
        # Pre-tax Margin
        fin_2['Pre-tax Margin'].iloc[i] = fin_2['Pretax Income'].iloc[i]/fin_2['Total Revenue'].iloc[i]
        num = fin_2['Pretax Income'].iloc[i]
        den = fin_2['Total Revenue'].iloc[i]
        if ((num < 0) & (den < 0)):
            fin_2['Pre-tax Margin'].iloc[i] = np.NaN
        
        # Net Profit Margin
        fin_2['Net Profit Margin'].iloc[i] = fin_2['Net Income'].iloc[i]/fin_2['Total Revenue'].iloc[i]
        num = fin_2['Net Income'].iloc[i]
        den = fin_2['Total Revenue'].iloc[i]
        if ((num < 0) & (den < 0)):
            fin_2['Net Profit Margin'].iloc[i] = np.NaN
        
        try:            
            if i != 0:
                fin_2['Return on Assets (Added Gr.Interest)'].iloc[i] = (fin_2['Net Income'].iloc[i] + (fin_2['Interest Expense Non Operating'].iloc[i]*(1-(fin_2['Tax Provision'].iloc[i]/fin_2['Pretax Income'].iloc[i]))))/((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)
                num = (fin_2['Net Income'].iloc[i] + (fin_2['Interest Expense Non Operating'].iloc[i]*(1-(fin_2['Tax Provision'].iloc[i]/fin_2['Pretax Income'].iloc[i]))))
                den = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)
                if ((num < 0) & (den < 0)):
                    fin_2['Return on Assets (Added Gr.Interest)'].iloc[i] = np.NaN
            else:
                fin_2['Return on Assets (Added Gr.Interest)'].iloc[i] = np.NaN
        except:
            fin_2['Return on Assets (Added Gr.Interest)'].iloc[i] = np.NaN
        
        # Operating Return on Assets
        try:
            if i != 0:
                fin_2['Operating Return on Assets'].iloc[i] = fin_2['EBIT'].iloc[i]/((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)
                num = fin_2['EBIT'].iloc[i]
                den = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)
                if ((num < 0) & (den < 0)):
                    fin_2['Operating Return on Assets'].iloc[i] = np.NaN
            else:
                fin_2['Operating Return on Assets'].iloc[i] = np.NaN
        except:
            fin_2['Operating Return on Assets'].iloc[i] = np.NaN
        
        # Modified Total Debt
        try:
            fin_2['Modified Total Debt'].iloc[i] = fin_2['Long Term Debt And Capital Lease Obligation'][i] + fin_2['Current Debt And Capital Lease Obligation'][i]
        except:
            try:    
                fin_2['Modified Total Debt'].iloc[i] = fin_2['Long Term Debt And Capital Lease Obligation'][i]
            except:
                try:    
                    fin_2['Modified Total Debt'].iloc[i] = fin_2['Current Debt And Capital Lease Obligation'][i]
                except:
                    fin_2['Modified Total Debt'].iloc[i] = np.NaN    
        
        # Return on Total Capital
        try:
            if i != 0:
                fin_2['Return on Total Capital'].iloc[i] = fin_2['EBIT'].iloc[i]/(((fin_2['Modified Total Debt'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i]) + (fin_2['Modified Total Debt'].iloc[i-1] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1]))/2)
                num = fin_2['EBIT'].iloc[i]
                den = (((fin_2['Modified Total Debt'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i]) + (fin_2['Modified Total Debt'].iloc[i-1] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1]))/2)
                if ((num < 0) & (den < 0)):
                    fin_2['Return on Total Capital'].iloc[i] = np.NaN
            else:
                fin_2['Return on Total Capital'].iloc[i] = np.NaN
        except:
            fin_2['Return on Total Capital'].iloc[i] = np.NaN
        
        # Return on Equity
        try:
            if i != 0:
                fin_2['Return on Equity'].iloc[i] = fin_2['Net Income'].iloc[i]/((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2)
                num = fin_2['Net Income'].iloc[i]
                den = ((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2)
                if ((num < 0) & (den < 0)):
                    fin_2['Return on Equity'].iloc[i] = np.NaN
            else:
                fin_2['Return on Equity'].iloc[i] = np.NaN    
        except:
            fin_2['Return on Equity'].iloc[i] = np.NaN
    
    # Activity/Utilization/Operating Efficiency Ratios
        
        # Receivables Turnover
        try:    
            if i != 0:
                fin_2['Receivables Turnover'].iloc[i] = fin_2['Total Revenue'].iloc[i]/((fin_2['Receivables'].iloc[i] + fin_2['Receivables'].iloc[i-1])/2)
                num = fin_2['Total Revenue'].iloc[i]
                den = ((fin_2['Receivables'].iloc[i] + fin_2['Receivables'].iloc[i-1])/2)
                if ((num < 0) & (den < 0)):
                    fin_2['Receivables Turnover'].iloc[i] = np.NaN
            else:
                fin_2['Receivables Turnover'].iloc[i] = np.NaN
        except:
            fin_2['Receivables Turnover'].iloc[i] = np.NaN
        
        # Days of Sales Outstanding
        fin_2['Days of Sales Outstanding'].iloc[i] = 365/fin_2['Receivables Turnover'].iloc[i]
        
        # Inventory Turnover
        try:
            if i != 0:
                fin_2['Inventory Turnover'].iloc[i] = fin_2['Cost of Revenue'].iloc[i]/((fin_2['Inventory'].iloc[i] + fin_2['Inventory'].iloc[i-1])/2)   
                num = fin_2['Cost of Revenue'].iloc[i]
                den = ((fin_2['Inventory'].iloc[i] + fin_2['Inventory'].iloc[i-1])/2)   
                if ((num < 0) & (den < 0)):
                    fin_2['Inventory Turnover'].iloc[i] = np.NaN
            else:
                fin_2['Inventory Turnover'].iloc[i] = np.NaN
        except:
            fin_2['Inventory Turnover'].iloc[i] = np.NaN
        
        # Days of Inventory on hand
        fin_2['Days of Inventory on hand'].iloc[i] = 365/fin_2['Inventory Turnover'].iloc[i]
        
        # Payables Turnover
        try:
            if i != 0:
                fin_2['Payables Turnover'].iloc[i] = (fin_2['Inventory'].iloc[i] - fin_2['Inventory'].iloc[i-1] + fin_2['Cost of Revenue'].iloc[i])/((fin_2['Accounts Payable'].iloc[i] + fin_2['Accounts Payable'].iloc[i-1])/2)
                num = (fin_2['Inventory'].iloc[i] - fin_2['Inventory'].iloc[i-1] + fin_2['Cost of Revenue'].iloc[i])
                den = ((fin_2['Accounts Payable'].iloc[i] + fin_2['Accounts Payable'].iloc[i-1])/2) 
                if ((num < 0) & (den < 0)):
                    fin_2['Payables Turnover'].iloc[i] = np.NaN
            else:
                fin_2['Payables Turnover'].iloc[i] = np.NaN
        except:
            fin_2['Payables Turnover'].iloc[i] = np.NaN
        
        # Days of Sales Payables
        fin_2['Days of Sales Payables'].iloc[i] = 365/fin_2['Payables Turnover'].iloc[i]
        
        # Total Asset Turnover
        if i != 0:
            fin_2['Total Asset Turnover'].iloc[i] = fin_2['Total Revenue'].iloc[i]/((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)   
            num = fin_2['Total Revenue'].iloc[i]
            den = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)   
            if ((num < 0) & (den < 0)):
                fin_2['Total Asset Turnover'].iloc[i] = np.NaN
        else:
            fin_2['Total Asset Turnover'].iloc[i] = np.NaN
        
        # Fixed Asset Turnover
        try:
            if i != 0:
                fin_2['Fixed Asset Turnover'].iloc[i] = fin_2['Total Revenue'].iloc[i]/((fin_2['Net PPE'].iloc[i] + fin_2['Net PPE'].iloc[i-1])/2)   
                num = fin_2['Total Revenue'].iloc[i]
                den = ((fin_2['Net PPE'].iloc[i] + fin_2['Net PPE'].iloc[i-1])/2)   
                if ((num < 0) & (den < 0)):
                    fin_2['Fixed Asset Turnover'].iloc[i] = np.NaN
            else:
                fin_2['Fixed Asset Turnover'].iloc[i] = np.NaN
        except:
            fin_2['Fixed Asset Turnover'].iloc[i] = np.NaN
        
        # Working Capital Turnover
        try:
            if i != 0:
                fin_2['Working Capital Turnover'].iloc[i] = fin_2['Total Revenue'].iloc[i]/(((fin_2['Current Assets'].iloc[i] - fin_2['Current Liabilities'].iloc[i]) + (fin_2['Current Assets'].iloc[i-1] - fin_2['Current Liabilities'].iloc[i-1]))/2)   
                num = fin_2['Total Revenue'].iloc[i]
                den = (((fin_2['Current Assets'].iloc[i] - fin_2['Current Liabilities'].iloc[i]) + (fin_2['Current Assets'].iloc[i-1] - fin_2['Current Liabilities'].iloc[i-1]))/2)     
                if ((num < 0) & (den < 0)):
                    fin_2['Working Capital Turnover'].iloc[i] = np.NaN
            else:
                fin_2['Working Capital Turnover'].iloc[i] = np.NaN
        except:
            fin_2['Working Capital Turnover'].iloc[i] = np.NaN
       
    # Liquidity Ratios
        
        # Current Ratio
        try:
            fin_2['Current Ratio'].iloc[i] = fin_2['Current Assets'].iloc[i]/fin_2['Current Liabilities'].iloc[i]
            num = fin_2['Current Assets'].iloc[i]
            den = fin_2['Current Liabilities'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Current Ratio'].iloc[i] = np.NaN
        except:
            fin_2['Current Ratio'].iloc[i] = np.NaN
        
        # Quick Ratio
        try:
            fin_2['Quick Ratio'].iloc[i] = (fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i] + fin_2['Receivables'].iloc[i])/fin_2['Current Liabilities'].iloc[i]
            num = (fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i] + fin_2['Receivables'].iloc[i])
            den = fin_2['Current Liabilities'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Quick Ratio'].iloc[i] = np.NaN
        except:
            fin_2['Quick Ratio'].iloc[i] = np.NaN
        
        # Cash Ratio
        try:
            fin_2['Cash Ratio'].iloc[i] = fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i]/fin_2['Current Liabilities'].iloc[i]
            num = fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i]
            den = fin_2['Current Liabilities'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Cash Ratio'].iloc[i] = np.NaN
        except:
            fin_2['Cash Ratio'].iloc[i] = np.NaN
        
        # Defensive Interval
        try:
            if i != 0:
                fin_2['Defensive Interval'].iloc[i] = (fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i] + fin_2['Receivables'].iloc[i])/(((fin_2['Cost of Revenue'].iloc[i] + fin_2['Operating Expense'].iloc[i]) + (fin_2['Cost of Revenue'].iloc[i-1] + fin_2['Operating Expense'].iloc[i-1]))/2)   
                num = (fin_2['Cash, Cash Equivalents & Short Term Investments'].iloc[i] + fin_2['Receivables'].iloc[i])
                den = (((fin_2['Cost of Revenue'].iloc[i] + fin_2['Operating Expense'].iloc[i]) + (fin_2['Cost of Revenue'].iloc[i-1] + fin_2['Operating Expense'].iloc[i-1]))/2) 
                if ((num < 0) & (den < 0)):
                    fin_2['Defensive Interval'].iloc[i] = np.NaN
            else:
                fin_2['Defensive Interval'].iloc[i] = np.NaN
        except:
            fin_2['Defensive Interval'].iloc[i] = np.NaN
        
        # Cash Conversion Cycle
        fin_2['Cash Conversion Cycle'].iloc[i] = fin_2['Days of Sales Outstanding'].iloc[i] + fin_2['Days of Inventory on hand'].iloc[i] - fin_2['Days of Sales Payables'].iloc[i]
        
    # Solvency Ratios
        
        # Debt-to-Equity
        try:
            fin_2['Debt-to-Equity'].iloc[i] = fin_2['Modified Total Debt'].iloc[i]/fin_2['Total Equity Gross Minority Interest'].iloc[i]
            num = fin_2['Modified Total Debt'].iloc[i]
            den = fin_2['Total Equity Gross Minority Interest'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Debt-to-Equity'].iloc[i] = np.NaN
        except:
            fin_2['Debt-to-Equity'].iloc[i] = np.NaN
        
        # Debt-to-Capital Ratio
        try:
            fin_2['Debt-to-Capital Ratio'].iloc[i] = fin_2['Modified Total Debt'].iloc[i]/(fin_2['Modified Total Debt'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i])
            num = fin_2['Modified Total Debt'].iloc[i]
            den = (fin_2['Modified Total Debt'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i])
            if ((num < 0) & (den < 0)):
                fin_2['Debt-to-Capital Ratio'].iloc[i] = np.NaN
        except:
            fin_2['Debt-to-Capital Ratio'].iloc[i] = np.NaN
        
        # Debt-to-Assets
        fin_2['Debt-to-Assets'].iloc[i] = fin_2['Modified Total Debt'].iloc[i]/fin_2['Total Assets'].iloc[i]
        num = fin_2['Modified Total Debt'].iloc[i]
        den = fin_2['Total Assets'].iloc[i]
        if ((num < 0) & (den < 0)):
            fin_2['Debt-to-Assets'].iloc[i] = np.NaN
        
        # Financial Leverage
        try:
            if i != 0:
                fin_2['Financial Leverage'].iloc[i] = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)/((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2) 
                num = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)
                den = ((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2) 
                if ((num < 0) & (den < 0)):
                    fin_2['Financial Leverage'].iloc[i] = np.NaN
            else:
                fin_2['Financial Leverage'].iloc[i] = np.NaN
        except:
            fin_2['Financial Leverage'].iloc[i] = np.NaN
        
        # Interest Coverage (Income)
        try:
            fin_2['Interest Coverage (Income)'].iloc[i] = fin_2['EBIT'].iloc[i]/fin_2['Interest Expense Non Operating'].iloc[i]
            num = fin_2['EBIT'].iloc[i]
            den = fin_2['Interest Expense Non Operating'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Interest Coverage (Income)'].iloc[i] = np.NaN
        except:
            fin_2['Interest Coverage (Income)'].iloc[i] = np.NaN
        
    # Performance Ratios
        
        # CFO-to-Net Revenue
        try:
            fin_2['CFO-to-Net Revenue'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/fin_2['Gross Profit'].iloc[i]        
            num = fin_2['Operating Cash Flow'].iloc[i]
            den = fin_2['Gross Profit'].iloc[i]  
            if ((num < 0) & (den < 0)):
                fin_2['CFO-to-Net Revenue'].iloc[i] = np.NaN
        except:
            fin_2['CFO-to-Net Revenue'].iloc[i] = np.NaN
        
        # CFO-to-Assets
        try:
            if i != 0:
                fin_2['CFO-to-Assets'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)    
                num = fin_2['Operating Cash Flow'].iloc[i]
                den = ((fin_2['Total Assets'].iloc[i] + fin_2['Total Assets'].iloc[i-1])/2)    
                if ((num < 0) & (den < 0)):
                    fin_2['CFO-to-Assets'].iloc[i] = np.NaN
            else:
                fin_2['CFO-to-Assets'].iloc[i] = np.NaN    
        except:
            fin_2['CFO-to-Assets'].iloc[i] = np.NaN
        
        # CFO-to-Equity
        try:
            if i != 0:
                fin_2['CFO-to-Equity'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2)
                num = fin_2['Operating Cash Flow'].iloc[i]
                den = ((fin_2['Total Equity Gross Minority Interest'].iloc[i] + fin_2['Total Equity Gross Minority Interest'].iloc[i-1])/2)
                if ((num < 0) & (den < 0)):
                    fin_2['CFO-to-Equity'].iloc[i] = np.NaN
            else:
                fin_2['CFO-to-Equity'].iloc[i] = np.NaN    
        except:
            fin_2['CFO-to-Equity'].iloc[i] = np.NaN
        
        # CFO-to-Op.Income
        try:
            fin_2['CFO-to-Op.Income'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/fin_2['EBIT'].iloc[i]
            num = fin_2['Operating Cash Flow'].iloc[i]
            den = fin_2['EBIT'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['CFO-to-Op.Income'].iloc[i] = np.NaN
        except:
            fin_2['CFO-to-Op.Income'].iloc[i] = np.NaN
        
    # Coverage Ratios
        
        # CFO-to-Debt
        try:
            fin_2['CFO-to-Debt'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/fin_2['Modified Total Debt'].iloc[i]
            num = fin_2['Operating Cash Flow'].iloc[i]
            den = fin_2['Modified Total Debt'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['CFO-to-Debt'].iloc[i] = np.NaN
        except:
            fin_2['CFO-to-Debt'].iloc[i] = np.NaN
        
        # Interest Coverage (CFO)
        try:
            fin_2['Interest Coverage (CFO)'].iloc[i] = (fin_2['Operating Cash Flow'].iloc[i] + fin_2['Interest Expense Non Operating'].iloc[i] + fin_2['Tax Provision'].iloc[i])/fin_2['Interest Expense Non Operating'].iloc[i]
            num = (fin_2['Operating Cash Flow'].iloc[i] + fin_2['Interest Expense Non Operating'].iloc[i] + fin_2['Tax Provision'].iloc[i])
            den = fin_2['Interest Expense Non Operating'].iloc[i]
            if ((num < 0) & (den < 0)):
                fin_2['Interest Coverage (CFO)'].iloc[i] = np.NaN
        except:
            fin_2['Interest Coverage (CFO)'].iloc[i] = np.NaN
        
        # Dividend Payment Coverage (CFO)
        try:
            fin_2['Dividend Payment Coverage (CFO)'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/-(fin_2['Cash Dividends Paid'].iloc[i])
            num = fin_2['Operating Cash Flow'].iloc[i]
            den = -(fin_2['Cash Dividends Paid'].iloc[i])
            if ((num < 0) & (den < 0)):
                fin_2['Dividend Payment Coverage (CFO)'].iloc[i] = np.NaN
        except:
            fin_2['Dividend Payment Coverage (CFO)'].iloc[i] = np.NaN
        
        # Outflows to CFI & CFF (CFO)
        try:
            fin_2['Outflows to CFI & CFF (CFO)'].iloc[i] = fin_2['Operating Cash Flow'].iloc[i]/-(fin_2['Purchase of PPE'].iloc[i] + fin_2['Purchase of Business'].iloc[i] + fin_2['Purchase of Investment'].iloc[i] + fin_2['Long Term Debt Payments'].iloc[i] + fin_2['Common Stock Payments'].iloc[i] + fin_2['Cash Dividends Paid'].iloc[i])    
            num = fin_2['Operating Cash Flow'].iloc[i]
            den = -(fin_2['Purchase of PPE'].iloc[i] + fin_2['Purchase of Business'].iloc[i] + fin_2['Purchase of Investment'].iloc[i] + fin_2['Long Term Debt Payments'].iloc[i] + fin_2['Common Stock Payments'].iloc[i] + fin_2['Cash Dividends Paid'].iloc[i])
            if ((num < 0) & (den < 0)):
                fin_2['Outflows to CFI & CFF (CFO)'].iloc[i] = np.NaN
        except:
            fin_2['Outflows to CFI & CFF (CFO)'].iloc[i] = np.NaN
        
        # Replacing inf with NaN
        fin_2.replace([np.inf, -np.inf], np.nan, inplace=True)
        
# Valuation Ratios on the latest year
    
    # Earnings to Price
    #fin_2['Earnings to Price'] = 0.00
    #fin_2['Earnings to Price'].iloc[-1] = ((fin_2['Net Income'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0])
    #fin_2['Earnings to Price'].iloc[0:(row_len-1)] = np.NaN
    
    # Sales to Price
    #fin_2['Sales to Price'] = 0.00
    #fin_2['Sales to Price'].iloc[-1] = ((fin_2['Total Revenue'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0]) 
    #fin_2['Sales to Price'].iloc[0:(row_len-1)] = np.NaN
    
    # CFO to Price
    #try:
    #    fin_2['CFO to Price'] = 0.00
    #    fin_2['CFO to Price'].iloc[-1] = ((fin_2['Operating Cash Flow'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0])
    #    fin_2['CFO to Price'].iloc[0:(row_len-1)] = np.NaN
    #except:
    #    fin_2['CFO to Price'] = np.NaN
    
    # Dividend to Price
    #try:
    #    fin_2['Dividend to Price'] = 0.00
    #    fin_2['Dividend to Price'].iloc[-1] = ((-fin_2['Cash Dividends Paid'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0])
    #    fin_2['Dividend to Price'].iloc[0:(row_len-1)] = np.NaN
    #except:
    #    fin_2['Dividend to Price'] = np.NaN
    
    # BV to Price
    #fin_2['BV to Price'] = 0.00
    #fin_2['BV to Price'].iloc[-1] = ((fin_2['Total Equity Gross Minority Interest'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0])
    #fin_2['BV to Price'].iloc[0:(row_len-1)] = np.NaN
    
    # EBITDA to EV
    #try:    
    #    fin_2['EBITDA to EV'] = 0.00
    #    fin_2['EBITDA to EV'].iloc[-1] = (fin_2['Normalized EBITDA'].iloc[-1]/sum1['Enterprise Value'][0])
    #    fin_2['EBITDA to EV'].iloc[0:(row_len-1)] = np.NaN
    #except:
    #    fin_2['EBITDA to EV'] = np.NaN
    
    # FCFF to Price
    #fin_2['FCFF to Price'] = 0.00
    #fin_2['FCFF to Price'].iloc[-1] = ((fin_2['Free Cash Flow'].iloc[-1]/fin_2['Basic Average Shares'].iloc[-1])/sum1['Price'][0])
    #fin_2['FCFF to Price'].iloc[0:(row_len-1)] = np.NaN    
    
    # Storing ticker data to dictionary
    financial_dir[t] = fin_2
    
    # Saving Analysis to Excel
    writer = pd.ExcelWriter(output_path+'{} - Analysis Part 1.xlsx'.format(t), engine='xlsxwriter')
    fin_2.to_excel(writer, sheet_name = str(t))
    writer.save()
        
    print('Analysis Part 1 for Ticker:{} saved as excel file'.format(t))
    
    # Now move to Analysis Part 1 - B