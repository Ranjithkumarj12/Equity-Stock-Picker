# -*- coding: utf-8 -*-
"""
Created on Thu Oct 27 20:28:26 2022

@author: jrkumar
"""

import pandas as pd

#Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

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

# Finance Industry-Sub Groups        
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


# Computing Industry Mean, Std Dev, and Z-Score
# Identifying the industry for the collected tickers
ind_stats = pd.read_excel(ind_stats_output+r'Industry Stats.xlsx', sheet_name=0)
ind_stats.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ind_stats = ind_stats[~ind_stats['Industry-Sub Group'].isin(fin_ind)]

# Calculating Z-Scores (All Tickers)
yr_list = ind_stats['Year'].tolist()
yr_list = list(set(yr_list))

# Creating a dataframe to collect tickers for each year
ticker_yr_df = pd.DataFrame()
no_zscore_df = pd.DataFrame()

for yr in yr_list:
    
    # Getting a unique list of Yahoo Industries
    ticker_full_zscore_df = pd.DataFrame()
    yr_df = ind_stats.loc[(ind_stats['Year'] == yr)]
    yr_df.reset_index(inplace = True, drop = True)
    ind_list = yr_df['Industry-Sub Group'].tolist()
    ind_list = list(set(ind_list))
    
    for i in range(len(yr_df['Tickers'])):
        industry1 = yr_df['Industry-Sub Group'][i]
        mcap_category1 = yr_df['MCap Category'][i]
        mcap_category1 = mcap_category1.replace(" Stock","")
        ticker_stats1 = pd.read_excel(output_path1+r'{} - Analysis Part 1.xlsx'.format(yr_df['Tickers'][i]), sheet_name=0)
        ticker_stats1 = ticker_stats1.loc[(ticker_stats1['Year'] == yr)]
        ticker_stats1.reset_index(inplace = True, drop = True)
        ticker_ind_stats1 = pd.read_excel(ind_stats_output+r'{} {} Stats - {}.xlsx'.format(yr,mcap_category1,industry1), sheet_name=0)
        ticker_ind_stats1.drop(columns = ['Unnamed: 0','Ticker'], inplace = True, axis = 1)
        ticker_ind_stats_mean1 = ticker_ind_stats1.iloc[-1]
        ticker_ind_stats_stddev1 = ticker_ind_stats1.iloc[-2]
        ticker_stats_zscores1 = pd.Series()
        
        # Creating columns in Analysis Part 1 for Industry mean and Std. dev
        for j in ratios_list:
            ticker_stats1['{} {} - Industry-Sub Group mean'.format(yr,j)] = ticker_ind_stats_mean1[j]
            ticker_stats1['{} {} - Industry-Sub Group Stddev'.format(yr,j)] = ticker_ind_stats_stddev1[j]
            ticker_stats1['{} {} - Z-Score'.format(yr,j)] =  (ticker_stats1[j] - ticker_stats1['{} {} - Industry-Sub Group mean'.format(yr,j)])/ticker_stats1['{} {} - Industry-Sub Group Stddev'.format(yr,j)]
        
        for k in ticker_stats1.columns:
            if 'Z-Score' in k:        
                ticker_stats_zscores1[k] = ticker_stats1[k][0]
        ticker_stats_zscores1.rename(yr_df['Tickers'][i], inplace = True)
        ticker_stats_zscores_df1 = pd.DataFrame(ticker_stats_zscores1)
        ticker_full_zscore_df = pd.concat([ticker_full_zscore_df,ticker_stats_zscores_df1],axis=1)
        print('{} Z-Scores for - {} is complete'.format(yr,yr_df['Tickers'][i]))
        
    # Ranking Tickers based on Ratios
    ticker_full_zscore_df = ticker_full_zscore_df.T
    ticker_full_zscore_df.reset_index(inplace = True, drop = False)
    ticker_full_zscore_df.rename(columns={'index':'Tickers'}, inplace=True)
    
    # Removing those tickers with NaN Z-Scores (only one company in the industry)
    no_zscore = []
    for i in range(len(ticker_full_zscore_df['Tickers'])):
        check_list = []
        for ratio in ratios_list:
            if pd.isnull(ticker_full_zscore_df['{} {} - Z-Score'.format(yr,ratio)])[i] == True:
                check_list.append(ratio)
        if len(check_list) == len(ratios_list):
            no_zscore.append(ticker_full_zscore_df['Tickers'][i])  
    
    # Check which industry the ticker is from, and take mid cap stocks from that industry
    # Final output should be altering industry stats and creating analysis Part 2
    yr_df2 = yr_df.copy()
    yr_df2['Tickers_New'] = yr_df2['Tickers']
    yr_df2.set_index('Tickers',inplace = True, drop = True)
    for x in no_zscore:
        industry2 = yr_df2['Industry-Sub Group'].loc[x]
        mcap_category2 = yr_df2['MCap Category'].loc[x]
        mcap_category2 = mcap_category2.replace(" Stock","")
        ticker_ind_stats2 = pd.read_excel(ind_stats_output+r'{} {} Stats - {}.xlsx'.format(yr,mcap_category2,industry2), sheet_name=0)
        ticker_ind_stats2.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
        # Dropping the mean and std.dev rows from the ticker industry stats file
        ticker_ind_stats2 = ticker_ind_stats2.drop([1,2])
        
        # Small Cap (without Z-Scores) - Consider Small Cap + Mid Cap to calculate industry metrics
        if yr_df2['MCap Category'].loc[x] == 'Small Cap Stock':
            small_mid_df = yr_df2.loc[(yr_df2['Industry-Sub Group'] == industry2) & ((yr_df2['MCap Category'] == 'Small Cap Stock') | (yr_df2['MCap Category'] == 'Mid Cap Stock'))]
            small_mid_df.reset_index(inplace = True, drop = True)
            for tick in range(1,len(small_mid_df)-1,1):
                industry3 = small_mid_df['Industry-Sub Group'].loc[tick]
                mcap_category3 = small_mid_df['MCap Category'].loc[tick]
                mcap_category3 = mcap_category3.replace(" Stock","")
                ticker_ind_stats3 = pd.read_excel(ind_stats_output+r'{} {} Stats - {}.xlsx'.format(yr,mcap_category3,industry3), sheet_name=0)
                ticker_ind_stats3.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                ticker_ind_stats4 = ticker_ind_stats3.loc[ticker_ind_stats3['Ticker'] == small_mid_df['Tickers_New'].iloc[tick]]
                # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
                ticker_ind_stats2 = ticker_ind_stats2.append(ticker_ind_stats4)
            if len(ticker_ind_stats2)>1:
                no_zscore.remove(x)
            # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)].mean(axis = 0)
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)-1].std(axis = 0)  
    
        # Mid Cap (without Z-Scores) - Consider Small Cap + Mid Cap + large Cap to calculate industry metrics
        elif yr_df2['MCap Category'].loc[x] == 'Mid Cap Stock':
            small_mid_large_df = yr_df2.loc[(yr_df2['Industry-Sub Group'] == industry2) & ((yr_df2['MCap Category'] == 'Small Cap Stock') | (yr_df2['MCap Category'] == 'Mid Cap Stock') | (yr_df2['MCap Category'] == 'Large Cap Stock'))]
            small_mid_large_df.reset_index(inplace = True, drop = True)
            for tick in range(1,len(small_mid_large_df)-1,1):
                industry3 = small_mid_large_df['Industry-Sub Group'].loc[tick]
                mcap_category3 = small_mid_large_df['MCap Category'].loc[tick]
                mcap_category3 = mcap_category3.replace(" Stock","")
                ticker_ind_stats3 = pd.read_excel(ind_stats_output+r'{} {} Stats - {}.xlsx'.format(yr,mcap_category3,industry3), sheet_name=0)
                ticker_ind_stats3.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                ticker_ind_stats4 = ticker_ind_stats3.loc[ticker_ind_stats3['Ticker'] == small_mid_large_df['Tickers_New'].iloc[tick]]
                # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
                ticker_ind_stats2 = ticker_ind_stats2.append(ticker_ind_stats4)
            if len(ticker_ind_stats2)>1:
                no_zscore.remove(x)
            # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)].mean(axis = 0)
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)-1].std(axis = 0)      
    
    # Large Cap (without Z-Scores) - Consider Mid Cap + Large Cap to calculate industry metrics
        elif yr_df2['MCap Category'].loc[x] == 'Large Cap Stock':
            mid_large_df = yr_df2.loc[(yr_df2['Industry-Sub Group'] == industry2) & ((yr_df2['MCap Category'] == 'Mid Cap Stock') | (yr_df2['MCap Category'] == 'Large Cap Stock'))]
            mid_large_df.reset_index(inplace = True, drop = True)
            for tick in range(1,len(mid_large_df)-1,1):
                industry3 = mid_large_df['Industry-Sub Group'].loc[tick]
                mcap_category3 = mid_large_df['MCap Category'].loc[tick]
                mcap_category3 = mcap_category3.replace(" Stock","")
                ticker_ind_stats3 = pd.read_excel(ind_stats_output+r'{} {} Stats - {}.xlsx'.format(yr,mcap_category3,industry3), sheet_name=0)
                ticker_ind_stats3.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
                ticker_ind_stats4 = ticker_ind_stats3.loc[ticker_ind_stats3['Ticker'] == mid_large_df['Tickers_New'].iloc[tick]]
                # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
                ticker_ind_stats2 = ticker_ind_stats2.append(ticker_ind_stats4)
            if len(ticker_ind_stats2)>1:
                no_zscore.remove(x)
            # Append the Ratio Values of small_mid_df tickers to the original no_zscore ticker 'x''s industry stats
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)].mean(axis = 0)
            ticker_ind_stats2.loc[len(ticker_ind_stats2.index)] = ticker_ind_stats2.iloc[:len(ticker_ind_stats2.index)-1].std(axis = 0)             
    
    # Removing the Z-Scores of tickers for which Z-Score is Zero (because of single company in an industry)
    no_zscore_temp_df = pd.DataFrame(no_zscore, columns = [yr])
    ticker_full_zscore_df = ticker_full_zscore_df[~(ticker_full_zscore_df['Tickers'].isin(no_zscore))]
    ticker_full_zscore_df.reset_index(inplace = True, drop = True)
        
    # Saving Z-Score data to excel
    writer = pd.ExcelWriter(output_path2+'{} All Ticker Z-Score Data.xlsx'.format(yr), engine='xlsxwriter')
    ticker_full_zscore_df.to_excel(writer)
    writer.save()
    
    # Storing tickers to no_zscore_df 
    no_zscore_df = pd.concat([no_zscore_df,no_zscore_temp_df],axis=1)
    
    # Storing tickers to ticker_yr_df
    ticker_yr_temp_df = ticker_full_zscore_df['Tickers']
    ticker_yr_temp_df.rename(yr, inplace = True)
    ticker_yr_df = pd.concat([ticker_yr_df,ticker_yr_temp_df],axis=1)

# Saving no_zscore tickers (due to only 1 company in an industry) to excel
writer = pd.ExcelWriter(output_path2+'Year-wise Tickers (without Z-Score).xlsx', engine='xlsxwriter')
no_zscore_df.to_excel(writer)
writer.save()    

# Saving ticker_yr_df to excel
writer = pd.ExcelWriter(output_path2+'Year-wise Tickers (with Z-Score).xlsx', engine='xlsxwriter')
ticker_yr_df.to_excel(writer)
writer.save()
    
# Now move to Analysis Part 3

        






        
