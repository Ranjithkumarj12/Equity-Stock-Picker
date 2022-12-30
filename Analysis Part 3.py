# -*- coding: utf-8 -*-
"""
Created on Sat Nov  5 21:59:13 2022

@author: jrkumar
"""

import pandas as pd
import numpy as np

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

# List of Ratios used
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
                          'CFO-to-Net Revenue',	'CFO-to-Assets',	'CFO-to-Equity',	'CFO-to-Op.Income',	'CFO-to-Debt',	'Interest Coverage (CFO)',
                          'Earnings to Price', 'Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

ratios_list_ascending = ['Days of Inventory on hand',
                         'Debt-to-Equity',	'Debt-to-Capital Ratio',	'Debt-to-Assets']

ratios_list_valuation = ['Earnings to Price',	'Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

ratios_not_using = ['Receivables Turnover',	'Days of Sales Outstanding','Payables Turnover',	'Days of Sales Payables','Quick Ratio',
                    'Dividend Payment Coverage (CFO)','Outflows to CFI & CFF (CFO)','Dividend to Price',	'EBITDA to EV',
                    'Defensive Interval',	'Cash Conversion Cycle']

# Reading Year-wise Tickers (with Z-Score)
ticker_yr_df = pd.read_excel(output_path2+r'Year-wise Tickers (with Z-Score).xlsx', sheet_name=0)
ticker_yr_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)

# Combining all tickers across years 
combined_ticker_list = []
for i in ticker_yr_df.columns:
    ticker_yr_temp_list = ticker_yr_df[i].dropna().values.tolist()
    combined_ticker_list = combined_ticker_list + ticker_yr_temp_list

# Unique list of tickers
combined_ticker_list2 = list(set(combined_ticker_list))

# Ticker has to have Z-Scores for atleast 3 years
combined_ticker_list3 = []
for tick in combined_ticker_list2:
    cnt = []
    for x in range(len(combined_ticker_list)):
        if combined_ticker_list[x] == tick:
           cnt.append(tick)
    if len(cnt) >= 3:
        combined_ticker_list3.append(tick)


# Ticker that don't have Z-Scores for atleast 3 years
combined_ticker_list4 = combined_ticker_list3.copy()
combined_ticker_list4.append(np.NaN)
ticker_yr_df2 = pd.DataFrame()
for i in range(len(ticker_yr_df.columns)):
    temp_list = []
    for j in range(len(ticker_yr_df[ticker_yr_df.columns[i]])):    
        if ticker_yr_df[ticker_yr_df.columns[i]].iloc[j] not in combined_ticker_list4:
            temp_list.append(ticker_yr_df[ticker_yr_df.columns[i]].iloc[j])
        else:
            pass
    temp_list_df = pd.DataFrame(temp_list, columns = [ticker_yr_df.columns[i]])
    ticker_yr_df2 = pd.concat([ticker_yr_df2,temp_list_df],axis=1)

# Saving Z-Score data to excel
writer = pd.ExcelWriter(output_path3+'Year-wise Tickers (without 3 yrs of Z-Scores).xlsx', engine='xlsxwriter')
ticker_yr_df2.to_excel(writer)
writer.save()

print('Year-wise Tickers (without 3 yrs of Z-Scores) is saved as excel') 
        
# Creating a column list for the zscore_aggregator dataframe
column_list = []
for i in range(len(ratios_list)):
    column_list.append(ratios_list[i] + ' - Z-Score (CAGR)')
    column_list.append(ratios_list[i] + ' - Z-Score (Trend)')
zscore_agg_df = pd.DataFrame(index = combined_ticker_list3, columns = column_list)

# Fetching Z-Scores across years and Aggregating them in zscore_aggregator dataframe
for t in combined_ticker_list3:
    # Check which years data is available for a specific ticker
    zscore_collector_df = pd.DataFrame()
    for col in ticker_yr_df.columns:
        if t in ticker_yr_df[col].dropna().values.tolist():
            zscore_collector_temp_df = pd.read_excel(output_path2+r'{} All Ticker Z-Score Data.xlsx'.format(col), sheet_name=0)
            zscore_collector_temp_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
            names = zscore_collector_temp_df.columns.tolist()
            for i in range(len(names)):
                names[i] = names[i].replace("{} ".format(col),"")
            zscore_collector_temp_df.columns = names
            zscore_collector_temp_df = zscore_collector_temp_df.loc[(zscore_collector_temp_df['Tickers'] == t)]
            zscore_collector_temp_df['Year'] = col            
            zscore_collector_df = zscore_collector_df.append(zscore_collector_temp_df)
   
    zscore_collector_df.reset_index(inplace = True, drop = True)
    yr_count = len(zscore_collector_df['Year'])-1
    
    for j in range(len(ratios_list)):
        numerator = zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[-1]
        denominator = zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[1]   
        if (numerator>0) & (denominator>0):
            div = numerator/denominator
            zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list[j])].loc[t] = ((div)**(1/yr_count)) - 1
        elif (numerator<0) & (denominator<0):
            div = abs(numerator)/abs(denominator)
            zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list[j])].loc[t] = (-1)*(((div)**(1/yr_count)) - 1)
        elif (numerator>0) & (denominator<0):
            zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list[j])].loc[t] = (((numerator + 2*abs(denominator))/abs(denominator))**(1/yr_count)) - 1
        elif (numerator<0) & (denominator>0):
            zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list[j])].loc[t] = (-1)*((((abs(numerator) + 2*denominator)/denominator)**(1/yr_count)) - 1)
        else:
            zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list[j])].loc[t] = 0
        
    for j in range(len(ratios_list)):   
        if yr_count == 3:
            if (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[1] > zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2]) & (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2] > zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[3]):
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Decreasing in last 2 leaps (4yr data)'
            elif (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[1] < zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2]) & (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2] < zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[3]):
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Increasing in last 2 leaps (4yr data)'
            else:
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Indeterminate/Fluctuating (4yr data)'
        
        elif yr_count == 2:
            if (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[1] > zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2]):
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Decreasing in last 1 leap (3yr data)'
            elif (zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[1] < zscore_collector_df['{} - Z-Score'.format(ratios_list[j])].iloc[2]):
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Increasing in last 1 leap (3yr data)'
            else:
                zscore_agg_df['{} - Z-Score (Trend)'.format(ratios_list[j])].loc[t] = 'Indeterminate/Fluctuating (3yr data)'
     
for ratio in range(len(ratios_list_descending)):
    zscore_agg_df['{} - Z-Score (CAGR) - Ranking'.format(ratios_list_descending[ratio])] = zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list_descending[ratio])].rank(ascending = False, method = 'average', na_option='bottom') 
for ratio in range(len(ratios_list_ascending)):
    zscore_agg_df['{} - Z-Score (CAGR) - Ranking'.format(ratios_list_ascending[ratio])] = zscore_agg_df['{} - Z-Score (CAGR)'.format(ratios_list_ascending[ratio])].rank(ascending = True, method = 'average', na_option='bottom') 

# Calculating Combined Score and Combined Rank
# Assigning equal weight/low weight
#count_ratios = len(ratios_list)
zscore_agg_df['Combined Score'] = ((0.0417*zscore_agg_df['Gross Profit Margin - Z-Score (CAGR) - Ranking']) + (0.0833*zscore_agg_df['Operating Profit Margin - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Pre-tax Margin - Z-Score (CAGR) - Ranking']) + (0.0833*zscore_agg_df['Net Profit Margin - Z-Score (CAGR) - Ranking']) + 
(0*zscore_agg_df['Return on Assets (Added Gr.Interest) - Z-Score (CAGR) - Ranking']) + (0.0833*zscore_agg_df['Operating Return on Assets - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Return on Total Capital - Z-Score (CAGR) - Ranking']) + (0.0333*zscore_agg_df['Return on Equity - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Inventory Turnover - Z-Score (CAGR) - Ranking']) +
(0*zscore_agg_df['Total Asset Turnover - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Fixed Asset Turnover - Z-Score (CAGR) - Ranking']) + (0.05*zscore_agg_df['Working Capital Turnover - Z-Score (CAGR) - Ranking']) + (0.025*zscore_agg_df['Current Ratio - Z-Score (CAGR) - Ranking']) + (0.025*zscore_agg_df['Cash Ratio - Z-Score (CAGR) - Ranking']) +
(0*zscore_agg_df['Financial Leverage - Z-Score (CAGR) - Ranking']) + (0.0417*zscore_agg_df['Interest Coverage (Income) - Z-Score (CAGR) - Ranking']) + (0.0677*zscore_agg_df['CFO-to-Net Revenue - Z-Score (CAGR) - Ranking']) + (0.0677*zscore_agg_df['CFO-to-Assets - Z-Score (CAGR) - Ranking']) + (0.0333*zscore_agg_df['CFO-to-Equity - Z-Score (CAGR) - Ranking']) +
(0*zscore_agg_df['CFO-to-Op.Income - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['CFO-to-Debt - Z-Score (CAGR) - Ranking']) + (0.0417*zscore_agg_df['Interest Coverage (CFO) - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Days of Inventory on hand - Z-Score (CAGR) - Ranking']) + (0.0833*zscore_agg_df['Debt-to-Equity - Z-Score (CAGR) - Ranking']) +
(0.0833*zscore_agg_df['Debt-to-Capital Ratio - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Debt-to-Assets - Z-Score (CAGR) - Ranking']) + (0.05*zscore_agg_df['Earnings to Price - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['Sales to Price - Z-Score (CAGR) - Ranking']) + (0*zscore_agg_df['CFO to Price - Z-Score (CAGR) - Ranking']) +
(0.05*zscore_agg_df['BV to Price - Z-Score (CAGR) - Ranking']) + (0.05*zscore_agg_df['FCFF to Price - Z-Score (CAGR) - Ranking']))
zscore_agg_df['Combined Rank'] = zscore_agg_df['Combined Score'].rank(ascending = True, method = 'average')

ind_stats_agg = pd.read_excel(ind_stats_output+r'Industry Stats.xlsx', sheet_name=0)
ind_stats_agg.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_unique = ind_stats_agg['Tickers'].tolist()
ticker_unique = list(set(ticker_unique))
ind_stats_agg_df = pd.DataFrame()
for ticker in ticker_unique:
    ind_stats_agg_temp = ind_stats_agg.loc[(ind_stats_agg['Tickers'] == ticker)]
    ind_stats_agg_temp = ind_stats_agg_temp.iloc[-1]     
    ind_stats_agg_df = ind_stats_agg_df.append(ind_stats_agg_temp)

zscore_agg_df.reset_index(inplace = True, drop = False)
zscore_agg_df.rename(columns={'index':'Tickers'}, inplace=True)
zscore_agg_df['Industry-Group'] = 'temp'
zscore_agg_df['Industry-Sub Group'] = 'temp'
zscore_agg_df['MCap Category'] = 'temp'
for t in range(len(zscore_agg_df)):
    for m in range(len(ind_stats_agg_df)):
        if ind_stats_agg_df['Tickers'].iloc[m] ==  zscore_agg_df['Tickers'].iloc[t]:
            zscore_agg_df['Industry-Group'].iloc[t] = ind_stats_agg_df['Industry-Group'].iloc[m]
            zscore_agg_df['Industry-Sub Group'].iloc[t] = ind_stats_agg_df['Industry-Sub Group'].iloc[m]
            zscore_agg_df['MCap Category'].iloc[t] = ind_stats_agg_df['MCap Category'].iloc[m]

zscore_agg_df.set_index('Tickers', inplace = True)
# Saving Z-Score data to excel
writer = pd.ExcelWriter(output_path3+'All Ticker Z-Score CAGR and Ranking.xlsx', engine='xlsxwriter')
zscore_agg_df.to_excel(writer)
writer.save()

print('Z-Score CAGR and Ranking for all tickers is complete and saved as excel')       

# Now move to Anlaysis Part 4
    
    
    
    
    