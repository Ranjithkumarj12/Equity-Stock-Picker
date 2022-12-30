# -*- coding: utf-8 -*-
"""
Created on Mon Oct 17 21:41:16 2022

@author: jrkumar
"""
import pandas as pd
import os

#Path to store extracted data
path_a = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
input_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

yf_a = []
yf_s = []
yf_final1 = []
yf_final2 = []
yf_final3 = []

#Identifying tickers with Annual Data
for file1 in os.listdir(path_a):
    f_name1 = file1.replace(" - Annual data.xlsx","")
    yf_a.append(f_name1)

#Identifying tickers with Industry Data
yf_secind1 = pd.read_excel(input_path+r'Final Ticker Industry Mapping.xlsx', sheet_name=0)
yf_secind2 = yf_secind1['Industry-Sub Group'].dropna().values.tolist()

#Identifying tickers with Stats & Price Data
for file3 in os.listdir(path_s1):
    f_name3 = file3.replace(" - Stats & Price data.xlsx","")
    yf_s.append(f_name3)
        
# Identifying commonly available tickers
for i in range(len(yf_a)):
    if yf_a[i] in yf_secind2:
        yf_final1.append(yf_a[i])

for i in range(len(yf_s)):
    if yf_s[i] in yf_secind2:
        yf_final2.append(yf_s[i])

for i in range(len(yf_final1)):
    if yf_final1[i] in yf_final2:
        yf_final3.append(yf_final1[i])

# Appending Final list of tickers to Annual Ticker Input
ticker_list_a = pd.read_excel(input_path+r'Total Ticker List - Annual.xlsx', sheet_name=0)
ticker_list_a.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_final_a = pd.DataFrame(yf_final3, columns = ['Final Consideration List'])
ticker_list_a2 = pd.concat([ticker_list_a,ticker_list_final_a],axis=1)

#Saving to Annual data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Annual.xlsx', engine='xlsxwriter')
ticker_list_a2.to_excel(writer)
writer.save()
       
# Appending Final list of tickers to Summary Ticker Input
ticker_list_s = pd.read_excel(input_path+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list_s.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_final_s = pd.DataFrame(yf_final3, columns = ['Final Consideration List'])
ticker_list_s2 = pd.concat([ticker_list_s,ticker_list_final_s],axis=1)

#Saving to Summary data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_s2.to_excel(writer)
writer.save()

# Now move to Daily Prices Data
    