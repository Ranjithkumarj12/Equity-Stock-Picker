# -*- coding: utf-8 -*-
"""
Created on Thu Nov 10 14:53:31 2022

@author: jrkumar
"""

import pandas as pd
import os

#Stored Path
input_path1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Industry Stats\\')

#Creating empty list to store list of tickers
files = []

# Industries (in Financial Services Sector) to Remove
fin_ind = ['Asset Management',
'Banks—Regional',
'Capital Markets',
'Insurance—Life',
'Insurance—Reinsurance',
'Insurance Brokers',
'Shell Companies',
'Credit Services',
'Banks—Diversified',
'Mortgage Finance',
'Financial Data & Stock Exchanges',
'Insurance—Property & Casualty',
'Insurance—Specialty',
'Insurance—Diversified',
'Financial Conglomerates']

# Creating Empty dataframe and list to store tickers belonging to Financial Services Sector
fin_tickers_list = []
fin_tickers_df = pd.DataFrame()

# Getting the list of tickers in the Analysis Part 1 folder    
for file1 in os.listdir(output_path1):
    f_name1 = file1.replace(".NS - Analysis Part 1.xlsx",".NS")
    files.append(f_name1)

# Reading Industry Stats to identify industry for the Ticker
ind_stats = pd.read_excel(ind_stats_output+r'Industry Stats Part 1.xlsx', sheet_name=0)
ind_stats.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)

# Fetching tickers which are operating in Financial Services Sector
for t in files:
    ind_stats_loc_df = pd.DataFrame()
    ind_stats_loc_df = ind_stats.loc[(ind_stats['Tickers'] == t)]
    ind_stats_loc_df.reset_index(inplace = True, drop = False)
    industry = ind_stats_loc_df['Industry'].iloc[len(ind_stats_loc_df['Industry'])-1]
    if industry in fin_ind:
        fin_tickers_list.append(t)
        # Removing Analysis Part 1 for these tickers
        os.remove(output_path1+t+' - Analysis Part 1.xlsx')
        print("Removed Ticker: {} from Analysis Part 1 as it belonged to Industry: {}".format(t,industry))

fin_tickers_df_temp = pd.DataFrame(fin_tickers_list, columns = ['Financial Indsutry Tickers'])
fin_tickers_df_temp['Industry'] = 'temp'
for tick in range(len(fin_tickers_df_temp['Financial Indsutry Tickers'])):
    ind_stats_loc_df2 = pd.DataFrame()
    ind_stats_loc_df2 = ind_stats.loc[(ind_stats['Tickers'] == fin_tickers_df_temp['Financial Indsutry Tickers'][tick])]
    ind_stats_loc_df2.reset_index(inplace = True, drop = False)
    industry = ind_stats_loc_df2['Industry'].iloc[len(ind_stats_loc_df2['Industry'])-1]
    fin_tickers_df_temp['Industry'][tick] = industry
fin_tickers_df = pd.concat([fin_tickers_df,fin_tickers_df_temp],axis=1)
    
# Saving Removed Financial Industry Tickers to excel
writer = pd.ExcelWriter(output_path2+'Financial Indsutry Tickers.xlsx'.format(t), engine='xlsxwriter')
fin_tickers_df.to_excel(writer)
writer.save()

print('Removed Financial Industry Tickers stored to excel')

# Now move to Analysis Part 2    
    
    
    
    
    
    
    
