# -*- coding: utf-8 -*-
"""
Created on Mon Oct 17 21:41:16 2022

@author: jrkumar
"""
import pandas as pd
import os
import fnmatch

#Path to store extracted data
path_a = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')
input_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

yf_a = []
yf_s = []
#Present in Annual, but missing in Summary
missing_tickers1 = []
#Present in Summary, but missing in Annual
missing_tickers2 = []
#Combined missing tickers
missing_tickers_total = []

#Renaming the file
for file1 in os.listdir(path_a):
    f_name1 = file1.replace(".NS - Annual data.xlsx",".NS")
    yf_a.append(f_name1)
    
for file2 in os.listdir(path_s1):
    if fnmatch.fnmatch(file2, '*.NS - Stats & Price data.xlsx'):
        f_name2 = file2.replace(".NS - Stats & Price data.xlsx",".NS")
        yf_s.append(f_name2)

#for file3 in os.listdir(path_s):
#    if fnmatch.fnmatch(file3, '*.NS - Sector & Industry data.xlsx'):
#        f_name3 = file3.replace(".NS - Sector & Industry data.xlsx",".NS")
#        yf_s.append(f_name3)
        
#We have to ensure we have stats & price data for all tickers whose annual data is available    
for i in yf_s:
    if i not in yf_a:
        missing_tickers1.append(i)
for j in yf_a:
    if j not in yf_s:
        missing_tickers2.append(j)

missing_tickers_total = missing_tickers1.copy() + missing_tickers2.copy()

#ticker_list_final_a = pd.read_excel(input_path+r'Total Ticker List - Annual.xlsx', sheet_name=0)
#temp = []
#for i in ticker_list_final_a['Total Ticker List - Annual']:
#    if i not in yf_a:
#        temp.append(i)
#temp_list_df = pd.DataFrame(temp, columns = ['Tickers not ending with ".NS" or ".BO"'])
#ticker_list_final_a = pd.concat([ticker_list_final_a,temp_list_df],axis=1)
        


#Adding Columns to Annual data
ticker_list_final_a = pd.read_excel(input_path+r'Total Ticker List - Annual.xlsx', sheet_name=0)
ticker_list_final_a .drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_a1 = pd.DataFrame(missing_tickers2, columns = ['Missing Stats&Price data, Present Annual data'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_a1],axis=1)
ticker_list_a2 = pd.DataFrame(missing_tickers1, columns = ['Missing Annual data, Present Stats&Price data'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_a2],axis=1)
ticker_list_a3 = pd.DataFrame(missing_tickers_total, columns = ['Missing Annual data/Stats&Price data'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_a3],axis=1)
misc_tickers_a = pd.Series(ticker_list_final_a['Tickers not ending with ".NS" or ".BO"'].dropna().values.tolist() + ticker_list_final_a['Missing Annual data/Stats&Price data'].dropna().values.tolist()).unique()
ticker_list_a4 = pd.DataFrame(misc_tickers_a, columns = ['Miscellaneous Tickers - Annual'])
ticker_list_a4.dropna(axis = 0,inplace = True) 
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_a4],axis=1)

#Adding Columns to Summary data
ticker_list_final_s = pd.read_excel(input_path+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list_final_s .drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_s1 = pd.DataFrame(missing_tickers2, columns = ['Missing Stats&Price data, Present Annual data'])
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_s1],axis=1)
ticker_list_s2 = pd.DataFrame(missing_tickers1, columns = ['Missing Annual data, Present Stats&Price data'])
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_s2],axis=1)
ticker_list_s3 = pd.DataFrame(missing_tickers_total, columns = ['Missing Annual data/Stats&Price data'])
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_s3],axis=1)
misc_tickers_s = pd.Series(ticker_list_final_s['Tickers not ending with ".NS" or ".BO" (No Stats/Price)'].dropna().values.tolist() + ticker_list_final_s['Missing Annual data/Stats&Price data'].dropna().values.tolist()).unique()
ticker_list_s4 = pd.DataFrame(misc_tickers_s, columns = ['Miscellaneous Tickers - Summary'])
ticker_list_s4.dropna(axis = 0,inplace = True) 
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_s4],axis=1)

#Adding Misc total to both Annual data and Summary data
misc_total1 = pd.Series(ticker_list_final_a['Miscellaneous Tickers - Annual'].dropna().values.tolist() + ticker_list_final_s['Miscellaneous Tickers - Summary'].dropna().values.tolist()).unique()   
ticker_list_misc = pd.DataFrame(misc_total1, columns = ['Miscellaneous Tickers - Total'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_misc],axis=1)
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_misc],axis=1)

#Fleshing out the Final Consideration list of Tickers
final_cons_list = []
final_list_chk = ticker_list_final_a['Total Ticker List - Annual'].dropna().values.tolist()
for i in final_list_chk:
    if i not in ticker_list_final_a['Miscellaneous Tickers - Total'].dropna().values.tolist():
        final_cons_list.append(i)        
        
ticker_list_conslist = pd.DataFrame(final_cons_list, columns = ['Final Consideration List'])
ticker_list_final_a = pd.concat([ticker_list_final_a,ticker_list_conslist],axis=1)
ticker_list_final_s = pd.concat([ticker_list_final_s,ticker_list_conslist],axis=1)

#Saving to Annual data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Annual.xlsx', engine='xlsxwriter')
ticker_list_final_a.to_excel(writer)
writer.save()
        
#Saving to Summary data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_final_s.to_excel(writer)
writer.save()

# Now move to Daily Prices Data
    