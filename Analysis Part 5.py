# -*- coding: utf-8 -*-
"""
Created on Fri Nov 11 13:15:05 2022

@author: jrkumar
"""

import pandas as pd

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
input_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
input_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
output_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 1 (Yahoo)\\')
output_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 2 (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path4 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 4 (Yahoo)\\')
output_path5 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 5 (Yahoo)\\')
price_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

# Reading final list of tickers that qualified for further processing
zscore_ticker_df = pd.read_excel(output_path3+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
zscore_ticker_df.rename(columns={'Unnamed: 0':'Tickers'}, inplace=True)
zscore_ticker_list = zscore_ticker_df['Tickers'].dropna().values.tolist()

# % Shares held by Insiders Part
# Creating empty data frame to store Price CAGR
insiders_temp_df2 = pd.DataFrame(index = zscore_ticker_list, columns = ['% Held by Insiders'])

# Fetching Price data for final list of tickers
for t in zscore_ticker_list:
    insiders_temp_df = pd.read_excel(input_path_s1+r'{} - Stats & Price data.xlsx'.format(t), sheet_name=0)
    insiders_temp_df2['% Held by Insiders'].loc[t] = insiders_temp_df['% Held by Insiders'][0]
    
# % Shares held by Insiders Ranking
insiders_temp_df2['% Held by Insiders - Ranking'] = insiders_temp_df2['% Held by Insiders'].rank(ascending = False, method = 'average', na_option='bottom')
    
# Saving Price Return Dataframe to excel
writer = pd.ExcelWriter(output_path5+'% Held by Insiders Ranking.xlsx', engine='xlsxwriter')
insiders_temp_df2.to_excel(writer)
writer.save()

print('% Held by Insiders - Ranking for all tickers is complete and saved as excel')
        
# Combining Price Return Rank with Z-Score Combined Rank
analysis_p5_df = pd.read_excel(output_path4+r'Z-Score and Price Return Ranking.xlsx', sheet_name=0)
#analysis_p5_df.drop(columns = ['Combined Score (with Price Return)','Combined Ranking (with Price Return)'], inplace = True, axis = 1)
#analysis_p5_df.rename(columns={'Combined Rank':'Combined Z-Scores - Ranking'}, inplace=True)

analysis_p5_df.set_index('Tickers',inplace = True, drop = True)
analysis_p5_df = pd.concat([analysis_p5_df, insiders_temp_df2], axis=1)

# 70% weightage to Price momentum, 30% weightage to Fundamental Z-Score
wt_price_return = 0.30
wt_zscores = 0.55
wt_insiders = 0.15
analysis_p5_df['Combined Score (with Price Return & Insiders Holding)'] = (wt_price_return * analysis_p5_df['Price Return - Ranking']) + (wt_zscores * analysis_p5_df['Combined Rank']) + (wt_insiders * analysis_p5_df['% Held by Insiders - Ranking'])
analysis_p5_df['Combined Ranking (with Price Return & Insiders Holding)'] = analysis_p5_df['Combined Score (with Price Return & Insiders Holding)'].rank(ascending = True, method = 'average', na_option='bottom')

# Saving Updated Analysis Dataframe (Part 3) to excel
writer = pd.ExcelWriter(output_path5+'Z-Score, Price Return, and Insiders Ranking.xlsx', engine='xlsxwriter')
analysis_p5_df.to_excel(writer)
writer.save()

print('Z-Score, Price Return, and Insiders Ranking for all Tickers is complete and saved as excel')
