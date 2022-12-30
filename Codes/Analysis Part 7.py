# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 19:05:04 2022

@author: jrkumar
"""

import pandas as pd
from datetime import date
import numpy as np

# Stored Path
input_path1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
output_path3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path6 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 6 (Yahoo)\\')
output_path7 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\\')
output_path7_rolling = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\Rolling Quarters\\')
output_path7_qoq = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 7 (Yahoo)\Q-o-Q\\')
                
ind_stats_output = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Industry Stats\\')

todays_date = date.today()

if (todays_date.month > 3) & (todays_date.month <= 6):
    qtr_month = 3
    qtr_year = todays_date.year
    qtr_minus1_month = 12
    qtr_minus1_year = todays_date.year-1
    qtr_minus2_month = 9
    qtr_minus2_year = todays_date.year-1
    
elif (todays_date.month > 6) & (todays_date.month <= 9):
    qtr_month= 6
    qtr_year = todays_date.year
    qtr_minus1_month = 3
    qtr_minus1_year = todays_date.year
    qtr_minus2_month = 12
    qtr_minus2_year = todays_date.year-1
    
elif (todays_date.month > 9) & (todays_date.month <= 12):
    qtr_month = 9
    qtr_year = todays_date.year
    qtr_minus1_month = 6
    qtr_minus1_year = todays_date.year
    qtr_minus2_month = 3
    qtr_minus2_year = todays_date.year
else:
    qtr_month = 12
    qtr_year = todays_date.year-1
    qtr_minus1_month = 9
    qtr_minus1_year = todays_date.year-1
    qtr_minus2_month = 6
    qtr_minus2_year = todays_date.year-1

ratios_list = ['Gross Profit Margin','Operating Profit Margin','Net Profit Margin']

# For Rolling Quarters
reqd_qtrs_full = 3
reqd_qtrs_partial = 2

# Reading final list of tickers that have 3 consistent quarterly results data
qtr_tickers_full_df = pd.read_excel(output_path6+r'Full data.xlsx', sheet_name=0)
qtr_tickers_full_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
qtr_tickers_full_list = qtr_tickers_full_df['Tickers with Full data'].dropna().values.tolist()

for t in qtr_tickers_full_list:
    inc_statement = pd.read_excel(input_path1+r'{} - Quarter data (Income Statement).xlsx'.format(t), sheet_name=0)
    inc_statement.rename(columns={'Unnamed: 0':'Rep_Date'}, inplace=True)
    inc_statement['Rep_date New'] = inc_statement['Rep_Date']
    
    inc_statement_full = inc_statement.copy() 
    
    inc_statement_full = inc_statement_full[1:reqd_qtrs_full+1]
    
    inc_statement_full['Gross Profit Margin'] = np.NaN
    inc_statement_full['Operating Profit Margin'] = np.NaN
    inc_statement_full['Net Profit Margin'] = np.NaN
    
    for i in range(len(inc_statement_full['Rep_date New'])):
        # Gross Profit Margin     
        try:
            inc_statement_full['Gross Profit Margin'].iloc[i] = inc_statement_full['Gross Profit'].iloc[i]/inc_statement_full['Total Revenue'].iloc[i]
            num = inc_statement_full['Gross Profit'].iloc[i]
            den = inc_statement_full['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_full['Gross Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_full['Gross Profit Margin'].iloc[i] = (inc_statement_full['Total Revenue'].iloc[i] - inc_statement_full['Cost of Revenue'].iloc[i])/inc_statement_full['Total Revenue'].iloc[i]
                num = (inc_statement_full['Total Revenue'].iloc[i] - inc_statement_full['Cost of Revenue'].iloc[i])
                den = inc_statement_full['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_full['Gross Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_full['Gross Profit Margin'].iloc[i] = np.NaN
         
        # Operating Profit Margin
        try:
            inc_statement_full['Operating Profit Margin'].iloc[i] = inc_statement_full['EBIT'].iloc[i]/inc_statement_full['Total Revenue'].iloc[i]
            num = inc_statement_full['EBIT'].iloc[i]
            den = inc_statement_full['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_full['Operating Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_full['Operating Profit Margin'].iloc[i] = inc_statement_full['Operating Income'].iloc[i]/inc_statement_full['Total Revenue'].iloc[i]
                num = inc_statement_full['Operating Income'].iloc[i]
                den = inc_statement_full['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_full['Operating Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_full['Operating Profit Margin'].iloc[i] = np.NaN
        
        # Net Profit Margin
        inc_statement_full['Net Profit Margin'].iloc[i] = inc_statement_full['Net Income'].iloc[i]/inc_statement_full['Total Revenue'].iloc[i]
        num = inc_statement_full['Net Income'].iloc[i]
        den = inc_statement_full['Total Revenue'].iloc[i]
        if ((num < 0) & (den < 0)):
            inc_statement_full['Net Profit Margin'].iloc[i] = np.NaN
            
    # Replacing infinite values with NaN
    inc_statement_full.replace([np.inf, -np.inf], np.nan, inplace=True)
    
    # Saving qtr_tickers_full_list to excel
    writer = pd.ExcelWriter(output_path7_rolling+'{} - Analysis Part 7.xlsx'.format(t), engine='xlsxwriter')
    inc_statement_full.to_excel(writer)
    writer.save()
        
    print('Ticker: {} is complete and saved as excel'.format(t))
    
# Reading final list of tickers which have either the recent two quarters of data or have q-2 and q-1 data without having the most recent quarters data
qtr_tickers_partial_df = pd.read_excel(output_path6+r'Partial data.xlsx', sheet_name=0)
qtr_tickers_partial_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
qtr_tickers_partial_list = qtr_tickers_partial_df['Tickers with Partial data'].dropna().values.tolist()

for t in qtr_tickers_partial_list:
    inc_statement = pd.read_excel(input_path1+r'{} - Quarter data (Income Statement).xlsx'.format(t), sheet_name=0)
    inc_statement.rename(columns={'Unnamed: 0':'Rep_Date'}, inplace=True)
    inc_statement['Rep_date New'] = inc_statement['Rep_Date']
    
    inc_statement_partial = inc_statement.copy() 
    inc_statement_partial = inc_statement_partial[1:reqd_qtrs_partial+1]

    inc_statement_partial['Gross Profit Margin'] = np.NaN
    inc_statement_partial['Operating Profit Margin'] = np.NaN
    inc_statement_partial['Net Profit Margin'] = np.NaN
    
    for i in range(len(inc_statement_partial['Rep_date New'])):
        # Gross Profit Margin     
        try:
            inc_statement_partial['Gross Profit Margin'].iloc[i] = inc_statement_partial['Gross Profit'].iloc[i]/inc_statement_partial['Total Revenue'].iloc[i]
            num = inc_statement_partial['Gross Profit'].iloc[i]
            den = inc_statement_partial['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_partial['Gross Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_partial['Gross Profit Margin'].iloc[i] = (inc_statement_partial['Total Revenue'].iloc[i] - inc_statement_partial['Cost of Revenue'].iloc[i])/inc_statement_partial['Total Revenue'].iloc[i]
                num = (inc_statement_partial['Total Revenue'].iloc[i] - inc_statement_partial['Cost of Revenue'].iloc[i])
                den = inc_statement_partial['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_partial['Gross Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_partial['Gross Profit Margin'].iloc[i] = np.NaN
         
        # Operating Profit Margin
        try:
            inc_statement_partial['Operating Profit Margin'].iloc[i] = inc_statement_partial['EBIT'].iloc[i]/inc_statement_partial['Total Revenue'].iloc[i]
            num = inc_statement_partial['EBIT'].iloc[i]
            den = inc_statement_partial['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_partial['Operating Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_partial['Operating Profit Margin'].iloc[i] = inc_statement_partial['Operating Income'].iloc[i]/inc_statement_partial['Total Revenue'].iloc[i]
                num = inc_statement_partial['Operating Income'].iloc[i]
                den = inc_statement_partial['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_partial['Operating Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_partial['Operating Profit Margin'].iloc[i] = np.NaN
        
        # Net Profit Margin
        inc_statement_partial['Net Profit Margin'].iloc[i] = inc_statement_partial['Net Income'].iloc[i]/inc_statement_partial['Total Revenue'].iloc[i]
        num = inc_statement_partial['Net Income'].iloc[i]
        den = inc_statement_partial['Total Revenue'].iloc[i]
        if ((num < 0) & (den < 0)):
            inc_statement_partial['Net Profit Margin'].iloc[i] = np.NaN
        
        
    # Replacing infinite values with NaN
    inc_statement_partial.replace([np.inf, -np.inf], np.nan, inplace=True)
    
    # Saving qtr_tickers_partial_list to excel
    writer = pd.ExcelWriter(output_path7_rolling+'{} - Analysis Part 7.xlsx'.format(t), engine='xlsxwriter')
    inc_statement_partial.to_excel(writer)
    writer.save()
    
    print('Ticker: {} is complete and saved as excel'.format(t))
    
# For Q-o-Q
# Reading final list of tickers which have data for both the recent quarter as well as for the same quarter in previous year
qtr_tickers_qoq_df = pd.read_excel(output_path6+r'QoQ Full data.xlsx', sheet_name=0)
qtr_tickers_qoq_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
qtr_tickers_qoq_list = qtr_tickers_qoq_df['Tickers with QoQ Full data'].dropna().values.tolist()

for t in qtr_tickers_qoq_list:
    inc_statement = pd.read_excel(input_path1+r'{} - Quarter data (Income Statement).xlsx'.format(t), sheet_name=0)
    inc_statement.rename(columns={'Unnamed: 0':'Rep_Date'}, inplace=True)
    inc_statement['Rep_date New'] = inc_statement['Rep_Date']
    
    inc_statement_qoq = inc_statement.copy() 
    
    inc_statement_qoq['Gross Profit Margin'] = np.NaN
    inc_statement_qoq['Operating Profit Margin'] = np.NaN
    inc_statement_qoq['Net Profit Margin'] = np.NaN    
    
    # Creating a new dataframe to store the two quarter rows
    inc_statement_qoq_new = pd.DataFrame(index = [0,1], columns = inc_statement_qoq.columns)
    
    # First row
    if len(inc_statement_qoq['Rep_date New'].iloc[1]) == 9:
        first_row_month_qoq = int(inc_statement_qoq['Rep_date New'].iloc[1][0])
        first_row_year_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[1][-4:])
    elif len(inc_statement_qoq['Rep_date New'].iloc[1]) == 10:
        first_row_month_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[1][0:1])
        first_row_year_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[1][-4:])
    
    #Take first row
    inc_statement_qoq_new.iloc[0] = inc_statement_qoq.iloc[1]
    
    # Take last row
    for row in (1,len(inc_statement_qoq['Rep_date New'])-1,1):
        if len(inc_statement_qoq['Rep_date New'].iloc[row]) == 9:
            row_month_qoq = int(inc_statement_qoq['Rep_date New'].iloc[row][0])
            row_year_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[row][-4:])
        elif len(inc_statement_qoq['Rep_date New'].iloc[1]) == 10:
            row_month_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[row][0:1])
            row_year_qoq =  int(inc_statement_qoq['Rep_date New'].iloc[row][-4:])
        if ((row_month_qoq == first_row_month_qoq) & (row_year_qoq == first_row_year_qoq-1)):
            inc_statement_qoq_new.iloc[1] = inc_statement_qoq.iloc[row]
        else:
            pass
    
    # Saving qtr_tickers_partial_list to excel
    writer = pd.ExcelWriter(output_path7_qoq+'{} - Analysis Part 7.xlsx'.format(t), engine='xlsxwriter')
    inc_statement_qoq_new.to_excel(writer)
    writer.save()
    
    inc_statement_qoq_new2 = pd.read_excel(output_path7_qoq+'{} - Analysis Part 7.xlsx'.format(t), sheet_name=0)
    inc_statement_qoq_new2.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
        
    for i in range(len(inc_statement_qoq_new2['Rep_date New'])):
        # Gross Profit Margin     
        try:
            inc_statement_qoq_new2['Gross Profit Margin'].iloc[i] = inc_statement_qoq_new2['Gross Profit'].iloc[i]/inc_statement_qoq_new2['Total Revenue'].iloc[i]
            num = inc_statement_qoq_new2['Gross Profit'].iloc[i]
            den = inc_statement_qoq_new2['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_qoq_new2['Gross Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_qoq_new2['Gross Profit Margin'].iloc[i] = (inc_statement_qoq_new2['Total Revenue'].iloc[i] - inc_statement_qoq_new2['Cost of Revenue'].iloc[i])/inc_statement_qoq_new2['Total Revenue'].iloc[i]
                num = (inc_statement_qoq_new2['Total Revenue'].iloc[i] - inc_statement_qoq_new2['Cost of Revenue'].iloc[i])
                den = inc_statement_qoq_new2['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_qoq_new2['Gross Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_qoq_new2['Gross Profit Margin'].iloc[i] = np.NaN
         
        # Operating Profit Margin
        try:
            inc_statement_qoq_new2['Operating Profit Margin'].iloc[i] = inc_statement_qoq_new2['EBIT'].iloc[i]/inc_statement_qoq_new2['Total Revenue'].iloc[i]
            num = inc_statement_qoq_new2['EBIT'].iloc[i]
            den = inc_statement_qoq_new2['Total Revenue'].iloc[i]
            if ((num < 0) & (den < 0)):
                inc_statement_qoq_new2['Operating Profit Margin'].iloc[i] = np.NaN
        except:
            try:
                inc_statement_qoq_new2['Operating Profit Margin'].iloc[i] = inc_statement_qoq_new2['Operating Income'].iloc[i]/inc_statement_qoq_new2['Total Revenue'].iloc[i]
                num = inc_statement_qoq_new2['Operating Income'].iloc[i]
                den = inc_statement_qoq_new2['Total Revenue'].iloc[i]
                if ((num < 0) & (den < 0)):
                    inc_statement_qoq_new2['Operating Profit Margin'].iloc[i] = np.NaN
            except:
                inc_statement_qoq_new2['Operating Profit Margin'].iloc[i] = np.NaN
        
        # Net Profit Margin
        inc_statement_qoq_new2['Net Profit Margin'].iloc[i] = inc_statement_qoq_new2['Net Income'].iloc[i]/inc_statement_qoq_new2['Total Revenue'].iloc[i]
        num = inc_statement_qoq_new2['Net Income'].iloc[i]
        den = inc_statement_qoq_new2['Total Revenue'].iloc[i]
        if ((num < 0) & (den < 0)):
            inc_statement_qoq_new2['Net Profit Margin'].iloc[i] = np.NaN
            
    # Replacing infinite values with NaN
    inc_statement_qoq_new2.replace([np.inf, -np.inf], np.nan, inplace=True)
    
    # Saving qtr_tickers_partial_list to excel
    writer = pd.ExcelWriter(output_path7_qoq+'{} - Analysis Part 7.xlsx'.format(t), engine='xlsxwriter')
    inc_statement_qoq_new2.to_excel(writer)
    writer.save()
    
    print('Ticker: {} is complete and saved as excel'.format(t))

# Combining full and partial lists for Rolling Quarters
qtr_tickers_list = qtr_tickers_full_list + qtr_tickers_partial_list

# Reading industry stats dataframe
industry_df = pd.read_excel(ind_stats_output+r'Industry Stats.xlsx', sheet_name=0)
industry_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)

column_list = []
for i in range(len(ratios_list)):
    column_list.append(ratios_list[i] + ' - Last 3 Quarter CAGR')
    column_list.append(ratios_list[i] + ' - Last 3 Quarter (Trend)')
    column_list.append(ratios_list[i] + ' - Q-o-Q CAGR')
    column_list.append(ratios_list[i] + ' - Q-o-Q (Trend)')
column_list.append('MCap Category')
column_list.append('Industry-Group')
column_list.append('Industry-Sub Group')
final_qtr_df = pd.DataFrame(index = qtr_tickers_list, columns = column_list)

for t in qtr_tickers_list:
    industry_temp_df = industry_df.loc[(industry_df['Tickers'] == t)]
    industry_temp_df.reset_index(drop = True, inplace = True)
    industry_temp_df = industry_temp_df.iloc[-1]
    industry_group = industry_temp_df['Industry-Group']
    industry_subgroup = industry_temp_df['Industry-Sub Group']
    mcap_category = industry_temp_df['MCap Category']
    
    final_qtr_df['Industry-Group'].loc[t] = industry_group
    final_qtr_df['Industry-Sub Group'].loc[t] = industry_subgroup
    final_qtr_df['MCap Category'].loc[t] = mcap_category
    
    # For Last 3 Quarters
    # Reading analysis7_df which has the 3 income sheet ratios

    analysis7_df = pd.read_excel(output_path7_rolling+'{} - Analysis Part 7.xlsx'.format(t), sheet_name=0)
    analysis7_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
    qtr_count = len(analysis7_df['Rep_Date'])-1
    
    # CAGR Computation
    for j in range(len(ratios_list)):    
        numerator1 = analysis7_df['{}'.format(ratios_list[j])].iloc[0]
        denominator1 = analysis7_df['{}'.format(ratios_list[j])].iloc[-1]   
        if (numerator1 == 0) | (denominator1 == 0):
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = np.NaN
        elif (np.isnan(numerator1)) | (np.isnan(denominator1)):
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = np.NaN
        elif (numerator1>0) & (denominator1>0):
            div = numerator1/denominator1
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = ((div)**(1/qtr_count)) - 1 
        elif (numerator1<0) & (denominator1<0):
            div = abs(numerator1)/abs(denominator1)
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = (-1)*(((div)**(1/qtr_count)) - 1)
        elif (numerator1>0) & (denominator1<0):
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = (((numerator1 + 2*abs(denominator1))/abs(denominator1))**(1/qtr_count)) - 1
        elif (numerator1<0) & (denominator1>0):
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = (-1)*((((abs(numerator1) + 2*denominator1)/denominator1)**(1/qtr_count)) - 1)
        else:
            final_qtr_df['{} - Last 3 Quarter CAGR'.format(ratios_list[j])].loc[t] = 0
    
    # Trend Identification
    for j in range(len(ratios_list)):   
        if qtr_count == 2:
            if (analysis7_df['{}'.format(ratios_list[j])].iloc[2] > analysis7_df['{}'.format(ratios_list[j])].iloc[1]) & (analysis7_df['{}'.format(ratios_list[j])].iloc[1] > analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Decreasing in last 2 leaps (3qtr data)'
            elif (analysis7_df['{}'.format(ratios_list[j])].iloc[2] < analysis7_df['{}'.format(ratios_list[j])].iloc[1]) & (analysis7_df['{}'.format(ratios_list[j])].iloc[1] < analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Increasing in last 2 leaps (3qtr data)'
            else:
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Indeterminate/Fluctuating (3qtr data)'
        
        elif qtr_count == 1:
            if (analysis7_df['{}'.format(ratios_list[j])].iloc[1] > analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Decreasing in last 1 leap (2qtr data)'
            elif (analysis7_df['{}'.format(ratios_list[j])].iloc[1] < analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Increasing in last 1 leap (2qtr data)'
            else:
                final_qtr_df['{} - Last 3 Quarter (Trend)'.format(ratios_list[j])].loc[t] = 'Indeterminate/Fluctuating (2qtr data)'
    
    print('Ticker: {} is complete and saved as excel'.format(t))
    
for t in qtr_tickers_qoq_list:
    industry_temp_df = industry_df.loc[(industry_df['Tickers'] == t)]
    industry_temp_df.reset_index(drop = True, inplace = True)
    industry_temp_df = industry_temp_df.iloc[-1]
    industry_group = industry_temp_df['Industry-Group']
    industry_subgroup = industry_temp_df['Industry-Sub Group']
    mcap_category = industry_temp_df['MCap Category']
    
    final_qtr_df['Industry-Group'].loc[t] = industry_group
    final_qtr_df['Industry-Sub Group'].loc[t] = industry_subgroup
    final_qtr_df['MCap Category'].loc[t] = mcap_category
    
    # For Q-o-Q
    # Reading analysis7_df which has the Q-o-Q income sheet ratios

    analysis7_df = pd.read_excel(output_path7_qoq+'{} - Analysis Part 7.xlsx'.format(t), sheet_name=0)
    analysis7_df.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
    qtr_count = len(analysis7_df['Rep_Date'])-1
    
    # CAGR Computation
    for j in range(len(ratios_list)):    
        numerator1 = analysis7_df['{}'.format(ratios_list[j])].iloc[0]
        denominator1 = analysis7_df['{}'.format(ratios_list[j])].iloc[-1]   
        if (numerator1 == 0) | (denominator1 == 0):
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = np.NaN
        elif (np.isnan(numerator1)) | (np.isnan(denominator1)):
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = np.NaN
        elif (numerator1>0) & (denominator1>0):
            div = numerator1/denominator1
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = ((div)**(1/qtr_count)) - 1 
        elif (numerator1<0) & (denominator1<0):
            div = abs(numerator1)/abs(denominator1)
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = (-1)*(((div)**(1/qtr_count)) - 1)
        elif (numerator1>0) & (denominator1<0):
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = (((numerator1 + 2*abs(denominator1))/abs(denominator1))**(1/qtr_count)) - 1
        elif (numerator1<0) & (denominator1>0):
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = (-1)*((((abs(numerator1) + 2*denominator1)/denominator1)**(1/qtr_count)) - 1)
        else:
            final_qtr_df['{} - Q-o-Q CAGR'.format(ratios_list[j])].loc[t] = 0
    
    # Trend Identification
    for j in range(len(ratios_list)):   
        if (analysis7_df['{}'.format(ratios_list[j])].iloc[1] > analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
            final_qtr_df['{} - Q-o-Q (Trend)'.format(ratios_list[j])].loc[t] = 'Decreasing in last 1 leap (Q-o-Q data)'
        elif (analysis7_df['{}'.format(ratios_list[j])].iloc[1] < analysis7_df['{}'.format(ratios_list[j])].iloc[0]):
            final_qtr_df['{} - Q-o-Q (Trend)'.format(ratios_list[j])].loc[t] = 'Increasing in last 1 leap (Q-o-Q data)'
        else:
            final_qtr_df['{} - Q-o-Q (Trend)'.format(ratios_list[j])].loc[t] = 'Indeterminate/Fluctuating (Q-o-Q data)'
    
    print('Ticker: {} is complete and saved as excel'.format(t))
    
# Saving qtr_tickers_partial_list to excel
writer = pd.ExcelWriter(output_path7+'All Quarterly CAGR and Trend.xlsx', engine='xlsxwriter')
final_qtr_df.to_excel(writer)
writer.save()

# Now move to Analysis 8
    
    