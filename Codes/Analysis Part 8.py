# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 19:05:04 2022

@author: jrkumar
"""

import pandas as pd
import numpy as np

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path6 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 6 (Yahoo)\\')
output_path7 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\\')
output_path7_rolling = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\Rolling Quarters\\')
output_path7_qoq = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\Q-o-Q\\')
output_path8 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 8 (Yahoo)\\')
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

# Ratios used
ratios_list = ['Gross Profit Margin','Operating Profit Margin','Net Profit Margin']

# Reading final list of tickers that have 3 consistent quarterly results data
qtr_cagr_df = pd.read_excel(output_path7+r'All Quarterly CAGR and Trend.xlsx', sheet_name=0)
qtr_cagr_df.rename(columns={'Unnamed: 0':'Tickers'}, inplace=True)
qtr_cagr_df.set_index('Tickers', inplace = True) 
for j in range(len(ratios_list)):
    qtr_cagr_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])] = np.NaN
    qtr_cagr_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = np.NaN
    qtr_cagr_df['{} - Ranking Last 3 Quarter CAGR'.format(ratios_list[j])] = np.NaN
    qtr_cagr_df['{} - Ranking Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = np.NaN    

# Small Cap Companies
# Getting a unique list of Industry-Sub Groups
sm_main_df =  qtr_cagr_df.loc[qtr_cagr_df['MCap Category'] == 'Small Cap Stock']  
sm_ind_list = sm_main_df['Industry-Sub Group'].dropna().values.tolist()
sm_ind_list = list(set(sm_ind_list))

# Calculating Z-Score
for ind in sm_ind_list:
    qtr_cagr_sm_ind_temp_df = qtr_cagr_df.loc[((qtr_cagr_df['Industry-Sub Group'] == ind) & ((qtr_cagr_df['MCap Category'] == 'Small Cap Stock') | (qtr_cagr_df['MCap Category'] == 'Mid Cap Stock')))]  
    for j in range(len(ratios_list)):
        qtr_cagr_sm_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])] = np.NaN
        qtr_cagr_sm_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = np.NaN
    
    column_list = []
    for i in range(len(ratios_list)):
        column_list.append(ratios_list[i] + ' - Last 3 Quarter CAGR')
        column_list.append(ratios_list[i] + ' - Q-o-Q CAGR')
    qtr_cagr_sm_stats = pd.DataFrame(index = ['Mean','Standard Deviation'], columns = column_list)
    
    for j in range(len(ratios_list)):
        qtr_cagr_sm_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_sm_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_sm_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_sm_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].std(axis = 0)   
        qtr_cagr_sm_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_sm_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_sm_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_sm_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].std(axis = 0)
        
        for tick in range(len(qtr_cagr_sm_ind_temp_df.index)):
            qtr_cagr_sm_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_sm_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_sm_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_sm_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation']
            qtr_cagr_sm_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_sm_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_sm_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_sm_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation']
    
        for sm_tick in qtr_cagr_sm_ind_temp_df.index:
            qtr_cagr_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[sm_tick] = qtr_cagr_sm_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[sm_tick]
            qtr_cagr_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[sm_tick] = qtr_cagr_sm_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[sm_tick]  

# Mid Cap Companies
# Getting a unique list of Industry-Sub Groups
md_main_df =  qtr_cagr_df.loc[qtr_cagr_df['MCap Category'] == 'Mid Cap Stock']  
md_ind_list = md_main_df['Industry-Sub Group'].dropna().values.tolist()
md_ind_list = list(set(md_ind_list))

# Calculating Z-Score
for ind in md_ind_list:
    qtr_cagr_md_ind_temp_df = qtr_cagr_df.loc[((qtr_cagr_df['Industry-Sub Group'] == ind) & ((qtr_cagr_df['MCap Category'] == 'Small Cap Stock') | (qtr_cagr_df['MCap Category'] == 'Mid Cap Stock') | (qtr_cagr_df['MCap Category'] == 'Large Cap Stock')))]  
    for j in range(len(ratios_list)):
        qtr_cagr_md_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])] = np.NaN
        qtr_cagr_md_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = np.NaN
    
    column_list = []
    for i in range(len(ratios_list)):
        column_list.append(ratios_list[i] + ' - Last 3 Quarter CAGR')
        column_list.append(ratios_list[i] + ' - Q-o-Q CAGR')
    qtr_cagr_md_stats = pd.DataFrame(index = ['Mean','Standard Deviation'], columns = column_list)
    
    for j in range(len(ratios_list)):
        qtr_cagr_md_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_md_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_md_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_md_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].std(axis = 0)   
        qtr_cagr_md_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_md_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_md_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_md_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].std(axis = 0)
        
        for tick in range(len(qtr_cagr_md_ind_temp_df.index)):
            qtr_cagr_md_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_md_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_md_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_md_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation']
            qtr_cagr_md_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_md_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_md_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_md_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation']
    
        for md_tick in qtr_cagr_md_ind_temp_df.index:
            qtr_cagr_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[md_tick] = qtr_cagr_md_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[md_tick]
            qtr_cagr_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[md_tick] = qtr_cagr_md_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[md_tick]  
    
# Large Cap Companies
# Getting a unique list of Industry-Sub Groups
lg_main_df =  qtr_cagr_df.loc[qtr_cagr_df['MCap Category'] == 'Large Cap Stock']  
lg_ind_list = lg_main_df['Industry-Sub Group'].dropna().values.tolist()
lg_ind_list = list(set(lg_ind_list))

# Calculating Z-Score
for ind in lg_ind_list:
    qtr_cagr_lg_ind_temp_df = qtr_cagr_df.loc[((qtr_cagr_df['Industry-Sub Group'] == ind) & ((qtr_cagr_df['MCap Category'] == 'Mid Cap Stock') | (qtr_cagr_df['MCap Category'] == 'Large Cap Stock')))]  
    for j in range(len(ratios_list)):
        qtr_cagr_lg_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])] = np.NaN
        qtr_cagr_lg_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = np.NaN
    
    column_list = []
    for i in range(len(ratios_list)):
        column_list.append(ratios_list[i] + ' - Last 3 Quarter CAGR')
        column_list.append(ratios_list[i] + ' - Q-o-Q CAGR')
    qtr_cagr_lg_stats = pd.DataFrame(index = ['Mean','Standard Deviation'], columns = column_list)
    
    for j in range(len(ratios_list)):
        qtr_cagr_lg_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_lg_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_lg_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_lg_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].std(axis = 0)   
        qtr_cagr_lg_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'] = qtr_cagr_lg_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].mean(axis = 0)
        qtr_cagr_lg_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation'] = qtr_cagr_lg_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].std(axis = 0)
        
        for tick in range(len(qtr_cagr_lg_ind_temp_df.index)):
            qtr_cagr_lg_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_lg_ind_temp_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_lg_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_lg_stats['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc['Standard Deviation']
            qtr_cagr_lg_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] = (qtr_cagr_lg_ind_temp_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].iloc[tick] - qtr_cagr_lg_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Mean'])/qtr_cagr_lg_stats['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc['Standard Deviation']
    
        for lg_tick in qtr_cagr_lg_ind_temp_df.index:
            qtr_cagr_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[lg_tick] = qtr_cagr_lg_ind_temp_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].loc[lg_tick]
            qtr_cagr_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[lg_tick] = qtr_cagr_lg_ind_temp_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].loc[lg_tick]  

for j in range(len(ratios_list)):
    qtr_cagr_df['{} - Ranking Last 3 Quarter CAGR'.format(ratios_list[j])] = qtr_cagr_df['{} - Z-Score Last 3 Quarter CAGR'.format(ratios_list[j])].rank(ascending = False, method = 'average', na_option='bottom')
    qtr_cagr_df['{} - Ranking Z-Score Q-o-Q CAGR'.format(ratios_list[j])] = qtr_cagr_df['{} - Z-Score Q-o-Q CAGR'.format(ratios_list[j])].rank(ascending = False, method = 'average', na_option='bottom')
    
# Calculating Combined Score and Rank for L3Q and QoQ
qtr_cagr_df['Combined Score - Z-Score Last 3 Quarter CAGR'] = 0.25*qtr_cagr_df['Gross Profit Margin - Ranking Last 3 Quarter CAGR'] + 0.25*qtr_cagr_df['Operating Profit Margin - Ranking Last 3 Quarter CAGR'] + 0.5*qtr_cagr_df['Net Profit Margin - Ranking Last 3 Quarter CAGR']
qtr_cagr_df['Combined Score - Z-Score Q-o-Q CAGR'] = 0.25*qtr_cagr_df['Gross Profit Margin - Ranking Z-Score Q-o-Q CAGR'] + 0.25*qtr_cagr_df['Operating Profit Margin - Ranking Z-Score Q-o-Q CAGR'] + 0.5*qtr_cagr_df['Net Profit Margin - Ranking Z-Score Q-o-Q CAGR']
qtr_cagr_df['Last 3 Quarter CAGR Ranking'] = qtr_cagr_df['Combined Score - Z-Score Last 3 Quarter CAGR'].rank(ascending = True, method = 'average', na_option='bottom')
qtr_cagr_df['Q-o-Q CAGR Ranking'] = qtr_cagr_df['Combined Score - Z-Score Q-o-Q CAGR'].rank(ascending = True, method = 'average', na_option='bottom')

wt_last3qtr = 0.8
wt_qoq = 0.2

qtr_cagr_df['Overall Quarterly Score'] = ((wt_last3qtr * qtr_cagr_df['Last 3 Quarter CAGR Ranking']) + (wt_qoq * qtr_cagr_df['Q-o-Q CAGR Ranking']))
qtr_cagr_df['Overall Quarterly Ranking'] = qtr_cagr_df['Overall Quarterly Score'].rank(ascending = True, method = 'average', na_option='bottom')

# Saving Overall Quarterly Score and Ranking for all Tickers
writer = pd.ExcelWriter(output_path8+'Overall Quarterly Score and Ranking.xlsx', engine='xlsxwriter')
qtr_cagr_df.to_excel(writer)
writer.save()

print('Overall Quarterly Score and Ranking for all Tickers is complete and saved as excel')

# Now move to Analysis 9




   