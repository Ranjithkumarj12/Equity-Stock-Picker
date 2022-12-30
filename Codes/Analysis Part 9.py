# -*- coding: utf-8 -*-
"""
Created on Mon Dec  5 10:33:29 2022

@author: jrkumar
"""

import pandas as pd
import numpy as np

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
input_path2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Shareholding Pattern (Screener)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path8 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 8 (Yahoo)\\')
output_path9 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 9 (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

# Reading Annual Z-Score Ranking
annual_df = pd.read_excel(output_path3+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
annual_df.set_index('Tickers',inplace = True, drop = True)
# Reading Quarterly Z-Score Ranking
qtr_df = pd.read_excel(output_path8+r'Overall Quarterly Score and Ranking.xlsx', sheet_name=0)
qtr_df.set_index('Tickers',inplace = True, drop = True)
qtr_df.rename(columns={'MCap Category':'MCap Category_Q','Industry-Group':'Industry-Group_Q','Industry-Sub Group':'Industry-Sub Group_Q'}, inplace=True)

# Collecting Tickers which don't have Quarterly data
#na_qtr = []
#for t in annual_df.index:
#    if t not in qtr_df.index.dropna().values.tolist():
#        na_qtr.append(t)

# Removing common columns
#cols_to_use = qtr_df.columns.difference(annual_df.columns)

# Merging Annual and Quarterly data
fundamental_df = pd.merge(annual_df, qtr_df, how = 'outer', on = 'Tickers')
#pd.merge(annual_df, qtr_df[cols_to_use], left_index=True, right_index=True, how='outer')

# Initial weights for Annual and Quarterly
wt_annual = 0.8
wt_qtr = 0.2

# Calculating Score and Rank for Annual and Quarterly Data
fundamental_df['Total Combined Score - A and Q'] = (wt_annual * fundamental_df['Combined Rank']) + (wt_qtr * fundamental_df['Overall Quarterly Ranking'])
fundamental_df['Total Combined Ranking - A and Q'] = fundamental_df['Total Combined Score - A and Q'].rank(ascending = True, method = 'average', na_option='bottom')

# Saving Total Combined Score and Ranking - A and Q, to excel
writer = pd.ExcelWriter(output_path9+'Total Combined Score and Ranking - A and Q.xlsx', engine='xlsxwriter')
fundamental_df.to_excel(writer)
writer.save()

print('Total Combined Score and Ranking - A and Q is complete and saved as excel')

# List of shareholders under consideration
shareholders_list = ['Promoters','FIIs','DIIs','Public']

# Calculating 1Q Change (Difference) and 2Q Change (Difference)
for s in range(len(shareholders_list)):
    fundamental_df['{} 1Q Change'.format(shareholders_list[s])] = np.NaN
    fundamental_df['{} 2Q Change'.format(shareholders_list[s])] = np.NaN

# Collecting Tickers which don't have Shareholder pattern data
sp_na = []

for t in fundamental_df.index:
    try:    
        sp_df = pd.read_excel(input_path2+r'{} - Shareholding Pattern Data.xlsx'.format(t), sheet_name=0)
        sp_df.rename(columns={'Unnamed: 0':'Shareholders'}, inplace=True)
        sp_df.set_index('Shareholders',inplace = True, drop = True)
        for m in sp_df.index:
            if m not in shareholders_list:
                sp_df = sp_df.drop(m)
        shareholders_list2 = sp_df.index.dropna().values.tolist()

        sp_df['1Q Change'] = sp_df.iloc[:,1] - sp_df.iloc[:,0]
        sp_df['2Q Change'] = sp_df.iloc[:,2] - sp_df.iloc[:,1]
        for s in range(len(shareholders_list2)):
            fundamental_df['{} 1Q Change'.format(shareholders_list2[s])].loc[t] = sp_df['1Q Change'].loc[shareholders_list2[s]]
            fundamental_df['{} 2Q Change'.format(shareholders_list2[s])].loc[t] = sp_df['2Q Change'].loc[shareholders_list2[s]]
    except:
        sp_na.append(t)

# Calculating Score and Rank for Individual Shareholding Pattern Data
# 2Q is the recent quarter, 1Q is the quarter before recent quarter
wt_1Q_change = 0.25
wt_2Q_change = 0.75

# Ranking the 1Q Change and 2Q Change
for s in range(len(shareholders_list)):
    fundamental_df['{} 1Q Change Ranking'.format(shareholders_list[s])] = fundamental_df['{} 1Q Change'.format(shareholders_list[s])].rank(ascending = False, method = 'average', na_option='bottom')
    fundamental_df['{} 2Q Change Ranking'.format(shareholders_list[s])] = fundamental_df['{} 2Q Change'.format(shareholders_list[s])].rank(ascending = False, method = 'average', na_option='bottom')
    fundamental_df['{} Combined Score'.format(shareholders_list[s])] = (wt_1Q_change * fundamental_df['{} 1Q Change Ranking'.format(shareholders_list[s])]) + (wt_2Q_change * fundamental_df['{} 2Q Change Ranking'.format(shareholders_list[s])])
    fundamental_df['{} Combined Ranking'.format(shareholders_list[s])] = fundamental_df['{} Combined Score'.format(shareholders_list[s])].rank(ascending = True, method = 'average', na_option='bottom')

# Calculating Score and Rank for Total Shareholding Pattern Data
wt_promoters = 0.2 
wt_FIIs = 0.35
wt_DIIs = 0.35
wt_public = 0.1

fundamental_df['Shareholder Combined Score'] = (wt_promoters * fundamental_df['Promoters Combined Ranking']) + (wt_FIIs * fundamental_df['FIIs Combined Ranking']) + (wt_DIIs * fundamental_df['DIIs Combined Ranking']) + (wt_public * fundamental_df['Public Combined Ranking'])
fundamental_df['Shareholder Combined Ranking'] = fundamental_df['Shareholder Combined Score'].rank(ascending = True, method = 'average', na_option='bottom')

# Calculating Score and Rank for Annual, Quarterly and Shareholding Pattern Data 
wt_annual = 0.7
wt_qtr = 0.2
wt_shareholder = 0.1

fundamental_df['Total Combined Score - A, Q and SP'] = (wt_annual * fundamental_df['Combined Rank']) + (wt_qtr * fundamental_df['Overall Quarterly Ranking']) + (wt_shareholder * fundamental_df['Shareholder Combined Ranking'])
fundamental_df['Total Combined Ranking - A, Q and SP'] = fundamental_df['Total Combined Score - A, Q and SP'].rank(ascending = True, method = 'average', na_option='bottom')

# Saving Total Combined Score and Ranking - A, Q and SP to excel
writer = pd.ExcelWriter(output_path9+'Total Combined Score and Ranking - A, Q and SP.xlsx', engine='xlsxwriter')
fundamental_df.to_excel(writer)
writer.save()

print('Total Combined Score and Ranking - A, Q and SP is complete and saved as excel')

# Done :)


