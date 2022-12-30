# -*- coding: utf-8 -*-
"""
Created on Thu Nov 10 14:53:31 2022

@author: jrkumar
"""

import pandas as pd
import os

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

# Industries (in Financial Services Sector) to Remove
fin_ind = ['Non Banking Financial Company (NBFC)',
'Housing Finance Company',
'Investment Company',
'Other Bank',
'Stockbroking & Allied',
'Public Sector Bank',
'Financial Technology (Fintech)',
'Other Financial Services',
'Depositories, Clearing Houses and Other Intermediaries',
'Private Sector Bank',
'Life Insurance',
'Holding Company',
'Other Capital Market related Services',
'Financial Institution',
'Exchange and Data Platform',
'Financial Products Distributor',
'Ratings',
'Asset Management Company',
'General Insurance',
'Depositories Clearing Houses and Other Intermediaries']

# Creating Empty dataframe and list to store tickers belonging to Financial Services Sector
fin_tickers_list = []
fin_tickers_df = pd.DataFrame()

# Getting the list of tickers in the Analysis Part 1 folder    
for file1 in os.listdir(output_path1):
    f_name1 = file1.replace(" - Analysis Part 1.xlsx","")
    files.append(f_name1)

# Reading Industry Stats to identify industry for the Ticker
ind_stats = pd.read_excel(ind_stats_output+r'Industry Stats.xlsx', sheet_name=0)
ind_stats.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)

# Fetching tickers which are operating in Financial Services Sector
for t in files:
    try:
        ind_stats_loc_df = pd.DataFrame()
        ind_stats_loc_df = ind_stats.loc[(ind_stats['Tickers'] == t)]
        ind_stats_loc_df.reset_index(inplace = True, drop = False)
        industry = ind_stats_loc_df['Industry-Sub Group'].iloc[len(ind_stats_loc_df['Industry-Sub Group'])-1]
        if industry in fin_ind:
            fin_tickers_list.append(t)
            # Removing Analysis Part 1 for these tickers
            os.remove(output_path1+t+' - Analysis Part 1.xlsx')
            print("Removed Ticker: {} from Analysis Part 1 as it belonged to Industry: {}".format(t,industry))
    except:
        pass
fin_tickers_df_temp = pd.DataFrame(fin_tickers_list, columns = ['Financial Indsutry Tickers'])
fin_tickers_df_temp['Industry-Sub Group'] = 'temp'
for tick in range(len(fin_tickers_df_temp['Financial Indsutry Tickers'])):
    ind_stats_loc_df2 = pd.DataFrame()
    ind_stats_loc_df2 = ind_stats.loc[(ind_stats['Tickers'] == fin_tickers_df_temp['Financial Indsutry Tickers'][tick])]
    ind_stats_loc_df2.reset_index(inplace = True, drop = False)
    industry = ind_stats_loc_df2['Industry-Sub Group'].iloc[len(ind_stats_loc_df2['Industry-Sub Group'])-1]
    fin_tickers_df_temp['Industry-Sub Group'][tick] = industry
fin_tickers_df = pd.concat([fin_tickers_df,fin_tickers_df_temp],axis=1)
    
# Saving Removed Financial Industry Tickers to excel
writer = pd.ExcelWriter(output_path2+'Financial Indsutry Tickers.xlsx', engine='xlsxwriter')
fin_tickers_df.to_excel(writer)
writer.save()

print('Removed Financial Industry Tickers stored to excel')

# Now move to Analysis Part 2    
    
    
    
    
    
    
    