# -*- coding: utf-8 -*-
"""
Created on Tue Oct 25 17:25:44 2022

@author: jrkumar
"""

import pandas as pd
import os
from datetime import date

#Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
input_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')
output_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')


#Creating empty dictionary to store ticker data
sec_dir = {}
secind_na = []
secind_a = []
files = []
mcap_na = []
mcap_a = []
mcap_price_discrepancy = []

# Getting the list of tickers in the Analysis folder    
for file1 in os.listdir(output_path):
    f_name1 = file1.replace(" - Analysis Part 1.xlsx","")
    files.append(f_name1)

# Saving tickers which had a min of 3 yrs of FS to Input annual and summary excel
ticker_list_final_a = pd.read_excel(input_path2+r'Total Ticker List - Annual.xlsx', sheet_name=0)
ticker_list_final_a.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_final_s = pd.read_excel(input_path2+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list_final_s.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_final = pd.DataFrame(files, columns = ['Final Consideration List - 3'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_final],axis=1)
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_final],axis=1)

#Saving to Annual data excel
writer = pd.ExcelWriter(input_path2+'Total Ticker List - Annual.xlsx', engine='xlsxwriter')
ticker_list_final_a.to_excel(writer)
writer.save()

#Saving to Annual data excel
writer = pd.ExcelWriter(input_path2+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_final_s.to_excel(writer)
writer.save()

# Creating empty dataframes to fill data
secind_data2 = pd.DataFrame(columns = ['Industry-Group','Industry-Sub Group'])

# Collating Sector/Industry and MCap data for Industry Stats
# Sector/Industry data
secind_data1 = pd.read_excel(input_path2+r'Final Ticker Industry Mapping.xlsx', sheet_name=0)
secind_data1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
        
# MCap data for all years
# Defining MCap Criteria
# Large Cap ~ MCap > INR 20,000 Cr., Mid Cap ~ INR 20,000 Cr. > MCap > INR 5,000 Cr., Small Cap ~ MCap < INR 5,000 Cr. 
upper_thld = 200000000000
lower_thld = 50000000000
final_list2 = pd.DataFrame()
for t in files:
    mcap_data1 = pd.read_excel(output_path+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
    try:
        mcap_data1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
    except:
        pass
    mcap_data1.set_index('Year', inplace = True) 
    mcap_data1['Final_MCap'] = 0.00
    mcap_data1['Industry-Group'] = 'temp'
    mcap_data1['Industry-Sub Group'] = 'temp'
    mcap_data1['Tickers'] = 'temp'
    
    mcap_data1.reset_index(inplace = True, drop = False)
    for i in range(len(mcap_data1['Year'])):                       
        mcap_data1['Final_MCap'].iloc[i] = mcap_data1['Price'].iloc[i] * mcap_data1['Ordinary Shares Number'].iloc[i]
        for p in range(len(secind_data1)):
            if secind_data1['Tickers'].iloc[p] == t:
                mcap_data1['Industry-Group'].iloc[i] =  secind_data1['Industry-Group'].iloc[p]
                mcap_data1['Industry-Sub Group'].iloc[i] =  secind_data1['Industry-Sub Group'].iloc[p]
                mcap_data1['Tickers'].iloc[i] = t
                
    # Final dataframe which contains tickers with MCap and Sector/Industry data
    final_list = mcap_data1.copy()
    final_list = final_list[['Tickers','Year','Industry-Group','Industry-Sub Group','Final_MCap']]
       
    # Bucketing securities based on Industry and MCap
    final_list['MCap Category'] = 'temp'
    for i in range(len(final_list['Final_MCap'])):
        if final_list['Final_MCap'][i] > upper_thld:
            final_list['MCap Category'][i] = 'Large Cap Stock'
        elif (final_list['Final_MCap'][i] <= upper_thld) & (final_list['Final_MCap'][i] > lower_thld):
            final_list['MCap Category'][i] = 'Mid Cap Stock'
        else:
            final_list['MCap Category'][i] = 'Small Cap Stock'
    
    # Keep appending tickers to final_list
    final_list2 = final_list2.append(final_list)
    final_list2.reset_index(drop = True, inplace = True)

final_list2.dropna(subset=['Industry-Group','Industry-Sub Group'],axis = 0, inplace = True)    
# Saving Analysis to Excel
writer = pd.ExcelWriter(ind_stats_output+'Industry Stats.xlsx', engine='xlsxwriter')
final_list2.to_excel(writer)
writer.save()
    
print('Industry Stats saved as excel file')

# Creating dataframes for Small, Mid and Large Cap
# List of Ratios used in Analysis Part 1
ratios_list = ['Gross Profit Margin',	'Operating Profit Margin',	'Pre-tax Margin',	'Net Profit Margin',	
               'Return on Assets (Added Gr.Interest)',	'Operating Return on Assets',	'Return on Total Capital',	'Return on Equity',		
               'Inventory Turnover',	'Days of Inventory on hand', 'Total Asset Turnover',	'Fixed Asset Turnover',	'Working Capital Turnover',	'Current Ratio',	
               	'Cash Ratio',	'Debt-to-Equity',	'Debt-to-Capital Ratio',	'Debt-to-Assets',	'Financial Leverage',	'Interest Coverage (Income)',	
               'CFO-to-Net Revenue',	'CFO-to-Assets',	'CFO-to-Equity',	'CFO-to-Op.Income',	'CFO-to-Debt',	'Interest Coverage (CFO)', 'Earnings to Price',
               'Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

ratios_not_using = ['Receivables Turnover',	'Days of Sales Outstanding','Payables Turnover',	'Days of Sales Payables','Quick Ratio',
                    'Dividend Payment Coverage (CFO)','Outflows to CFI & CFF (CFO)','Dividend to Price',	'EBITDA to EV',
                    'Defensive Interval',	'Cash Conversion Cycle']

# Getting a unique list of Yahoo Industries
yr_list = final_list2['Year'].tolist()
yr_list = list(set(yr_list))

for yr in yr_list:
    
    # Getting a unique list of Yahoo Industries
    yr_df = final_list2.loc[(final_list2['Year'] == yr)]
    ind_list = yr_df['Industry-Sub Group'].tolist()
    ind_list = list(set(ind_list))
    
    # Small Cap Dataframe
    sm_ticker_ind_map = pd.DataFrame()
    for i in ind_list:
        small_df = final_list2.loc[(final_list2['Year'] == yr)]
        small_df = small_df.loc[(small_df['Industry-Sub Group'] == i) & (small_df['MCap Category'] == 'Small Cap Stock')]  
        small_list = small_df['Tickers'].values.tolist()
        small_list_df = pd.DataFrame(small_list, columns = [i])
        sm_ticker_ind_map = pd.concat([sm_ticker_ind_map,small_list_df],axis=1)
    
    # Creating a empty dataframe
    sm_data_df = pd.DataFrame(index = ratios_list, columns = ind_list)
    
    # Fetching ratios data for companies within a specific industry
    for i in sm_data_df.columns:
        for j in range(len(sm_ticker_ind_map.columns)):
            if i == sm_ticker_ind_map.columns[j]:
                if len(sm_ticker_ind_map[sm_ticker_ind_map.columns[j]].dropna().values.tolist()) != 0 :
                    ind_ticker_df = pd.DataFrame()
                    ind_ticker_list = sm_ticker_ind_map[sm_ticker_ind_map.columns[j]].dropna().values.tolist()
                    for t in ind_ticker_list:
                        ind_ticker_data1 = pd.read_excel(output_path+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
                        ind_ticker_data1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                        ind_ticker_data2 = ind_ticker_data1[ratios_list].iloc[-1]
                        ind_ticker_data2['Ticker'] = t
                        ind_ticker_df = ind_ticker_df.append(ind_ticker_data2)
                    ind_ticker_df.reset_index(drop = True, inplace = True)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)].mean(axis = 0)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)-1].std(axis = 0)
                    
                    # Saving Industry stats to excel
                    writer = pd.ExcelWriter(ind_stats_output+'{} Small Cap Stats - {}.xlsx'.format(yr,i), engine='xlsxwriter')
                    ind_ticker_df.to_excel(writer)
                    writer.save()
                    print('{} Industry Stats for Small Cap - {} saved as excel file'.format(yr,i))
                else:
                    pass
            
    # Mid Cap Dataframe    
    md_ticker_ind_map = pd.DataFrame()
    for i in ind_list:
        mid_df = final_list2.loc[(final_list2['Year'] == yr)]
        mid_df = mid_df.loc[(mid_df['Industry-Sub Group'] == i) & (mid_df['MCap Category'] == 'Mid Cap Stock')] 
        mid_list = mid_df['Tickers'].values.tolist()
        mid_list_df = pd.DataFrame(mid_list, columns = [i])
        md_ticker_ind_map = pd.concat([md_ticker_ind_map,mid_list_df],axis=1)
        
    # Creating a empty dataframe
    md_data_df = pd.DataFrame(index = ratios_list, columns = ind_list)
    
    # Fetching ratios data for companies within a specific industry
    for i in md_data_df.columns:
        for j in range(len(md_ticker_ind_map.columns)):
            if i == md_ticker_ind_map.columns[j]:
                if len(md_ticker_ind_map[md_ticker_ind_map.columns[j]].dropna().values.tolist()) != 0 :
                    ind_ticker_df = pd.DataFrame()
                    ind_ticker_list = md_ticker_ind_map[md_ticker_ind_map.columns[j]].dropna().values.tolist()
                    for t in ind_ticker_list:
                        ind_ticker_data1 = pd.read_excel(output_path+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
                        ind_ticker_data1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                        ind_ticker_data2 = ind_ticker_data1[ratios_list].iloc[-1]
                        ind_ticker_data2['Ticker'] = t
                        ind_ticker_df = ind_ticker_df.append(ind_ticker_data2)
                    ind_ticker_df.reset_index(drop = True, inplace = True)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)].mean(axis = 0)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)-1].std(axis = 0)
                    
                    # Saving Industry stats to excel
                    writer = pd.ExcelWriter(ind_stats_output+'{} Mid Cap Stats - {}.xlsx'.format(yr,i), engine='xlsxwriter')
                    ind_ticker_df.to_excel(writer)
                    writer.save()
                    print('{} Industry Stats for Mid Cap - {} saved as excel file'.format(yr,i))
                else:
                    pass
    
    # Large Cap Dataframe
    lg_ticker_ind_map = pd.DataFrame()
    for i in ind_list:
        large_df = final_list2.loc[(final_list2['Year'] == yr)]
        large_df = large_df.loc[(large_df['Industry-Sub Group'] == i) & (large_df['MCap Category'] == 'Large Cap Stock')]  
        large_list = large_df['Tickers'].values.tolist()
        large_list_df = pd.DataFrame(large_list, columns = [i])
        lg_ticker_ind_map = pd.concat([lg_ticker_ind_map,large_list_df],axis=1)  
    
    # Creating a empty dataframe
    lg_data_df = pd.DataFrame(index = ratios_list, columns = ind_list)
    
    # Fetching ratios data for companies within a specific industry
    for i in lg_data_df.columns:
        for j in range(len(lg_ticker_ind_map.columns)):
            if i == lg_ticker_ind_map.columns[j]:
                if len(lg_ticker_ind_map[lg_ticker_ind_map.columns[j]].dropna().values.tolist()) != 0 :
                    ind_ticker_df = pd.DataFrame()
                    ind_ticker_list = lg_ticker_ind_map[lg_ticker_ind_map.columns[j]].dropna().values.tolist()
                    for t in ind_ticker_list:
                        ind_ticker_data1 = pd.read_excel(output_path+r'{} - Analysis Part 1.xlsx'.format(t), sheet_name=0)
                        ind_ticker_data1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                        ind_ticker_data2 = ind_ticker_data1[ratios_list].iloc[-1]
                        ind_ticker_data2['Ticker'] = t
                        ind_ticker_df = ind_ticker_df.append(ind_ticker_data2)
                    ind_ticker_df.reset_index(drop = True, inplace = True)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)].mean(axis = 0)
                    ind_ticker_df.loc[len(ind_ticker_df.index)] =  ind_ticker_df.iloc[:len(ind_ticker_df.index)-1].std(axis = 0)
                    
                    # Saving Industry stats to excel
                    writer = pd.ExcelWriter(ind_stats_output+'{} Large Cap Stats - {}.xlsx'.format(yr,i), engine='xlsxwriter')
                    ind_ticker_df.to_excel(writer)
                    writer.save()
                    print('{} Industry Stats for Large Cap - {} saved as excel file'.format(yr,i))
                else:
                    pass

# Now Move to Analysis Part 1 - C