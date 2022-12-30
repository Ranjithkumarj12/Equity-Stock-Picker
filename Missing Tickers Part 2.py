# -*- coding: utf-8 -*-
"""
Created on Mon Oct 17 21:41:16 2022

@author: jrkumar
"""
import pandas as pd
import os

#Path to store extracted data
path_p = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')
input_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

yf_fin = []
yf_p = []
yf_final4 = []

#Identifying tickers with Annual Data
for file1 in os.listdir(path_p):
    f_name1 = file1.replace(" - Daily Price data.xlsx","")
    yf_p.append(f_name1)

#Identifying tickers with Industry Data
yf_secind1 = pd.read_excel(input_path+r'Total Ticker List - Annual.xlsx', sheet_name=0)
yf_secind1.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
yf_secind2 = yf_secind1['Final Consideration List'].dropna().values.tolist()


# Identifying commonly available tickers
for i in range(len(yf_p)):
    if yf_p[i] in yf_secind2:
        yf_final4.append(yf_p[i])

# Appending Final list of tickers to Annual Ticker Input
ticker_list_final_a = pd.DataFrame(yf_final4, columns = ['Final Consideration List - 2'])
yf_secind1 = pd.concat([yf_secind1,ticker_list_final_a],axis=1)

#Saving to Annual data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Annual.xlsx', engine='xlsxwriter')
yf_secind1.to_excel(writer)
writer.save()
       
# Appending Final list of tickers to Summary Ticker Input
ticker_list_s = pd.read_excel(input_path+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list_s.drop(columns = ['Unnamed: 0'], inplace = True, axis = 1)
ticker_list_final_s = pd.DataFrame(yf_final4, columns = ['Final Consideration List - 2'])
ticker_list_s2 = pd.concat([ticker_list_s,ticker_list_final_s],axis=1)

#Saving to Summary data excel
writer = pd.ExcelWriter(input_path+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_s2.to_excel(writer)
writer.save()

# Now move to Daily Prices Data
    