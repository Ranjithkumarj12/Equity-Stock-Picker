# -*- coding: utf-8 -*-
"""
Created on Wed Oct 12 12:59:45 2022

@author: jrkumar
"""

#Importing Libraries
import pandas as pd
import requests as rq
from bs4 import BeautifulSoup as bs
import fnmatch
import os
import numpy as np
#import re

input_path_s = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

ticker_list1 = pd.read_excel(input_path_s+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list1.dropna(axis = 0,inplace = True)
ticker_list1 = ticker_list1[ticker_list1['Total Ticker List - Summary'] !=  'NULL']
tickers1 = ticker_list1['Total Ticker List - Summary'].tolist()

#Input Tickers
tickers_yahoo_s1 = tickers1[:36]
tickers_na_yahoo_s1 = []
tickers_noprice_nostats1 = []
tickers_noprofile1 = []

#Measure Start time
#t_start = time.time()   

#Initiating Combined Dictionary to store data of all tickers
combined_dict2 = {}

#Path to store extracted data
output_path_s1 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Stats & Price data\\')
output_path_s2 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Sector & Industry data\\')
output_path_s3 = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Summary (Yahoo)\Total data\\')

# Part 1 - Web Scrapping Process - Using Selenium to interact, and Beautiful Soup to parse
for t in tickers_yahoo_s1:
#Round 1 - Key Stats & Price    
    try:
        temp_dir2 = {}
        
        url_stats = 'https://finance.yahoo.com/quote/{}/key-statistics?p={}'.format(t,t)
        
        headers = {"User-Agent" : "Chrome/100.0.4896.75"}
        page_stats = rq.get(url_stats, headers = headers)
        page_content_stats = page_stats.content
        soup_stats = bs(page_content_stats,'html.parser')
        table_data_stats = soup_stats.find_all('table', attrs = {'class':'W(100%) Bdcl(c) '})
        for i in table_data_stats:
            rows_stats = i.find_all('tr')
            for row_stats in rows_stats:      
                x = row_stats.get_text(separator = '|').split('|')[-1].replace(",","")
                x = x.replace("-","")
                x = x.replace("%","E-02")
                x = x.replace("k","E+03")
                x = x.replace("M","E+06")
                x = x.replace("B","E+09")
                x = x.replace("T","E+12")
                y = row_stats.get_text(separator = '|').split('|')[0]
                temp_dir2[y] = x
                
        print('Key Stats for Ticker:{} is complete'.format(t))
        
        # If Real time Prices are needed (not working btw)
        #price_data = soup.body.findAll(text=re.compile('^e3b14781$'))
        #price = price_data.get_text(separator = '|').split('|')
        
        # Price data (delayed) Scrapping
        price_data = soup_stats.find('fin-streamer', attrs = {'class':'Fw(b) Fz(36px) Mb(-4px) D(ib)'})
        price = price_data.get_text(separator = '|').split('|')[0].replace(",","")
        temp_dir2['Price'] = price
        
        #Converting to Dataframe
        temp_df2 = pd.DataFrame([temp_dir2], columns = temp_dir2.keys())
        temp_df2 ['Year'] = 'Stats Data'
        temp_df2 .set_index('Year', inplace = True)
        
        #Saving to Excel
        writer = pd.ExcelWriter(output_path_s1+'{} - Stats & Price data.xlsx'.format(t), engine='xlsxwriter')
        temp_df2.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Stats & Price data for Ticker:{} saved as excel file'.format(t))
        
    except:
        tickers_noprice_nostats1.append(t)
    
# Round 1 - Sector & Industry
    try:    
        temp_dir3 = {}
        
        url_secind = 'https://finance.yahoo.com/quote/{}/profile?p={}'.format(t,t)
        
        headers = {"User-Agent" : "Chrome/100.0.4896.75"}
        page_secind = rq.get(url_secind, headers = headers)
        page_content_secind = page_secind.content
        soup_secind = bs(page_content_secind,'html.parser')
        table_data_secind = soup_secind.find('p', attrs = {'class':'D(ib) Va(t)'})
        sector_x = table_data_secind.get_text(separator = '|').split('|')[0]
        sector_x = sector_x[:-3]
        sector_y = table_data_secind.get_text(separator = '|').split('|')[2]
        industry_x = table_data_secind.get_text(separator = '|').split('|')[3]
        industry_y = table_data_secind.get_text(separator = '|').split('|')[5]
        temp_dir3[sector_x] = sector_y
        temp_dir3[industry_x] = industry_y
                
        print('Sector & Industry for Ticker:{} is complete'.format(t))
        
        #Converting to Dataframe
        temp_df3 = pd.DataFrame([temp_dir3], columns = temp_dir3.keys())
        temp_df3 ['Year'] = 'Sector/Industry Data'
        temp_df3 .set_index('Year', inplace = True)
        
        #Saving to Excel - Sector & Industry
        writer = pd.ExcelWriter(output_path_s2+'{} - Sector & Industry data.xlsx'.format(t), engine='xlsxwriter')
        temp_df3.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Sector & Industry data for Ticker:{} saved as excel file'.format(t))
        
        temp_dir2.update(temp_dir3)
        temp_dir4 = temp_dir2.copy()
        combined_dict2[t] = temp_dir4
        
        #Converting to Dataframe
        temp_df4 = pd.DataFrame([temp_dir4], columns = temp_dir4.keys())
        temp_df4['Year'] = 'Summary Data'
        temp_df4.set_index('Year', inplace = True)
        
        #Saving to Excel - Total
        writer = pd.ExcelWriter(output_path_s3+'{} - Total data.xlsx'.format(t), engine='xlsxwriter')
        temp_df4.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Total data for Ticker:{} saved as excel file'.format(t))
    
    except:
        tickers_noprofile1.append(t)

#Combining noprice_nostats with noprofile
tickers_na_yahoo_s1 = tickers_noprice_nostats1.copy()
for i in tickers_noprofile1:
    if i not in tickers_noprice_nostats1:
        tickers_na_yahoo_s1.append(i)        

#Collecting Tickers which got Processed
tickers_yahoo_s1done = []
for i in tickers_yahoo_s1:
    if i not in tickers_na_yahoo_s1:
        tickers_yahoo_s1done.append(i)

#Creating new columns to ticker_list        
ticker_list2 = pd.DataFrame(tickers_yahoo_s1done, columns = ['Tickers ending with ".NS"'])
ticker_list_final = pd.concat([ticker_list1,ticker_list2],axis=1)
ticker_list_noprice_nostats1 = pd.DataFrame(tickers_noprice_nostats1, columns = ['No Stats/Price ending with ".NS"'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list_noprice_nostats1],axis=1)
ticker_list_noprofile1 = pd.DataFrame(tickers_noprofile1, columns = ['No Sector/Industry ending with ".NS"'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list_noprofile1],axis=1)
ticker_list3 = pd.DataFrame(tickers_na_yahoo_s1, columns = ['Tickers not ending with ".NS"'])
#Converting Tickers which don't end with ".NS" to ".BO"
ticker_list3['Tickers converted to end with ".BO"'] = np.NaN
for i in range(len(ticker_list3['Tickers not ending with ".NS"'])):
    ticker_list3['Tickers converted to end with ".BO"'][i] = ticker_list3['Tickers not ending with ".NS"'][i].replace(".NS",".BO")
ticker_bo_list = pd.Series(ticker_list3['Tickers converted to end with ".BO"'].dropna().values.tolist(),name='Tickers converted to end with ".BO"')
ticker_list_final = pd.concat([ticker_list_final,ticker_bo_list],axis=1)

#Saving to excel
writer = pd.ExcelWriter(input_path_s+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_final.to_excel(writer)
writer.save()

ticker_list_final = pd.read_excel(input_path_s+r'Total Ticker List - Summary.xlsx', sheet_name=0)

tickers2 = ticker_list_final['Tickers converted to end with ".BO"'].dropna().tolist()

#Tickers converted to end with ".BO"
tickers_yahoo_s2 = tickers2
#For tickers which don't end with ".NS" or ".BO" (Miscellaneous Tickers)
tickers_na_yahoo_s2 = []
#For tickers which don't have either Stats & Price or Sector & Industry
tickers_noprice_nostats2 = []
tickers_noprofile2 = []

# Part 2 - Web Scrapping Process - Using Selenium to interact, and Beautiful Soup to parse
#Round 2 - Key Stats & Price

for t in tickers_yahoo_s2:
    try:
        temp_dir5 = {}
        
        url_stats = 'https://finance.yahoo.com/quote/{}/key-statistics?p={}'.format(t,t)
        
        headers = {"User-Agent" : "Chrome/100.0.4896.75"}
        page_stats = rq.get(url_stats, headers = headers)
        page_content_stats = page_stats.content
        soup_stats = bs(page_content_stats,'html.parser')
        table_data_stats = soup_stats.find_all('table', attrs = {'class':'W(100%) Bdcl(c) '})
        for i in table_data_stats:
            rows_stats = i.find_all('tr')
            for row_stats in rows_stats:      
                x = row_stats.get_text(separator = '|').split('|')[-1].replace(",","")
                x = x.replace("-","")
                x = x.replace("%","E-02")
                x = x.replace("k","E+03")
                x = x.replace("M","E+06")
                x = x.replace("B","E+09")
                x = x.replace("T","E+12")
                y = row_stats.get_text(separator = '|').split('|')[0]
                temp_dir5[y] = x
                
        print('Key Stats for Ticker:{} is complete'.format(t))
        
        # If Real time Prices are needed (not working btw)
        #price_data = soup.body.findAll(text=re.compile('^e3b14781$'))
        #price = price_data.get_text(separator = '|').split('|')
        
        # Price data (delayed) Scrapping
        price_data = soup_stats.find('fin-streamer', attrs = {'class':'Fw(b) Fz(36px) Mb(-4px) D(ib)'})
        price = price_data.get_text(separator = '|').split('|')[0]
        temp_dir5['Price'] = price
        
        #Converting to Dataframe
        temp_df5 = pd.DataFrame([temp_dir5], columns = temp_dir5.keys())
        temp_df5 ['Year'] = 'Stats Data'
        temp_df5 .set_index('Year', inplace = True)
        
        #Saving to Excel
        writer = pd.ExcelWriter(output_path_s1+'{} - Stats & Price data.xlsx'.format(t), engine='xlsxwriter')
        temp_df5.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Stats & Price data for Ticker:{} saved as excel file'.format(t))
        
    except:
        tickers_noprice_nostats2.append(t)
    
    # Round 1 - Sector & Industry
    try:    
        temp_dir6 = {}
        
        url_secind = 'https://finance.yahoo.com/quote/{}/profile?p={}'.format(t,t)
        
        headers = {"User-Agent" : "Chrome/100.0.4896.75"}
        page_secind = rq.get(url_secind, headers = headers)
        page_content_secind = page_secind.content
        soup_secind = bs(page_content_secind,'html.parser')
        table_data_secind = soup_secind.find('p', attrs = {'class':'D(ib) Va(t)'})
        sector_x = table_data_secind.get_text(separator = '|').split('|')[0]
        sector_x = sector_x[:-3]
        sector_y = table_data_secind.get_text(separator = '|').split('|')[2]
        industry_x = table_data_secind.get_text(separator = '|').split('|')[3]
        industry_y = table_data_secind.get_text(separator = '|').split('|')[5]
        temp_dir6[sector_x] = sector_y
        temp_dir6[industry_x] = industry_y
                
        print('Sector & Industry for Ticker:{} is complete'.format(t))
        
        #Converting to Dataframe
        temp_df6 = pd.DataFrame([temp_dir6], columns = temp_dir6.keys())
        temp_df6 ['Year'] = 'Sector/Industry Data'
        temp_df6 .set_index('Year', inplace = True)
        
        writer = pd.ExcelWriter(output_path_s2+'{} - Sector & Industry data.xlsx'.format(t), engine='xlsxwriter')
        temp_df6.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Sector & Industry data for Ticker:{} saved as excel file'.format(t))
        
        temp_dir5.update(temp_dir6)
        temp_dir7 = temp_dir5.copy()
        combined_dict2[t] = temp_dir7
        
        #Converting to Dataframe
        temp_df7 = pd.DataFrame([temp_dir7], columns = temp_dir7.keys())
        temp_df7['Year'] = 'Summary Data'
        temp_df7.set_index('Year', inplace = True)
        
        #Saving to Excel - Total
        writer = pd.ExcelWriter(output_path_s3+'{} - Total data.xlsx'.format(t), engine='xlsxwriter')
        temp_df7.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        print('Total data for Ticker:{} saved as excel file'.format(t))
    
    except:
        tickers_noprofile2.append(t)
  
#Combining noprice_nostats with noprofile
tickers_na_yahoo_s2 = tickers_noprice_nostats2.copy()
for i in tickers_noprofile2:
    if i not in tickers_noprice_nostats2:
        tickers_na_yahoo_s2.append(i)     
    
#Collecting Tickers which got Processed
tickers_yahoo_s2done = []
for i in tickers_yahoo_s2:
    if i not in tickers_na_yahoo_s2:
        tickers_yahoo_s2done.append(i)

#Replacing back ".BO" with ".NS"
for i in range(len(tickers_yahoo_s2done)):
    tickers_yahoo_s2done[i] = tickers_yahoo_s2done[i].replace(".BO",".NS")
for i in range(len(tickers_noprice_nostats2)):
    tickers_noprice_nostats2[i] = tickers_noprice_nostats2[i].replace(".BO",".NS")
for i in range(len(tickers_noprofile2)):
    tickers_noprofile2[i] = tickers_noprofile2[i].replace(".BO",".NS")

#Creating new columns to ticker_list
ticker_list4 = pd.DataFrame(tickers_yahoo_s2done, columns = ['Tickers ending with ".BO"'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list4],axis=1)
ticker_list_noprice_nostats2 = pd.DataFrame(tickers_noprice_nostats2, columns = ['No Stats/Price ending with ".NS" or ".BO"'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list_noprice_nostats2],axis=1)
ticker_list_noprofile2 = pd.DataFrame(tickers_noprofile2, columns = ['No Sector/Industry ending with ".NS" or ".BO"'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list_noprofile2],axis=1)
# Tickers not ending with .NS or .BO will be same as No Stats/Price ending with ".NS" or ".BO".
#Here the ticker is not considered if it doesn't have Stats/Price data, irrespective of whethere it has Sector/Industry data
ticker_list5 = pd.DataFrame(ticker_list_noprice_nostats2, columns = ['Tickers not ending with ".NS" or ".BO" (No Stats/Price)'])
ticker_list_final = pd.concat([ticker_list_final,ticker_list5],axis=1)

#Saving to excel
writer = pd.ExcelWriter(input_path_s+'Total Ticker List - Summary.xlsx', engine='xlsxwriter')
ticker_list_final.to_excel(writer)
writer.save()

#Miscellaneous tickers in tickers_na_yahoo_a2 can either be left out or could be researched further for their correct ticker names

# Part - 3:
#File Name Changer (from ".BO" to ".NS")
#Stats & Price data
NS_with_BO_1 = []
for file in os.listdir(output_path_s1):
    if fnmatch.fnmatch(file, '*.BO - Stats & Price data.xlsx'):
        f_name1 = file.replace(".BO - Stats & Price data.xlsx",".BO")
        f_name2 = file.replace(".BO - Stats & Price data.xlsx",".NS")
        NS_with_BO_1.append(f_name1)        
        
        sec_list = pd.read_excel(output_path_s1+r'{}'.format(file), sheet_name = f_name1)
        
        writer = pd.ExcelWriter(output_path_s1+'{} - Stats & Price data.xlsx'.format(f_name2), engine='xlsxwriter')
        sec_list.to_excel(writer, sheet_name = str(f_name2))
        writer.save()
        
        os.remove(output_path_s1+file)
        print("Renamed {} and Removed Old file for {}".format(f_name1,f_name2))

NS_with_BO_2 = []
#Sector & Industry data
for file in os.listdir(output_path_s2):
    if fnmatch.fnmatch(file, '*.BO - Sector & Industry data.xlsx'):
        f_name1 = file.replace(".BO - Sector & Industry data.xlsx",".BO")
        f_name2 = file.replace(".BO - Sector & Industry data.xlsx",".NS")
        NS_with_BO_2.append(f_name1)        
        
        sec_list = pd.read_excel(output_path_s2+r'{}'.format(file), sheet_name = f_name1)
        
        writer = pd.ExcelWriter(output_path_s2+'{} - Sector & Industry data data.xlsx'.format(f_name2), engine='xlsxwriter')
        sec_list.to_excel(writer, sheet_name = str(f_name2))
        writer.save()
        
        os.remove(output_path_s2+file)
        print("Renamed {} and Removed Old file for {}".format(f_name1,f_name2))

NS_with_BO_3 = []
#Total data
for file in os.listdir(output_path_s3):
    if fnmatch.fnmatch(file, '*.BO - Total data.xlsx'):
        f_name1 = file.replace(".BO - Total data.xlsx",".BO")
        f_name2 = file.replace(".BO - Total data.xlsx",".NS")
        NS_with_BO_3.append(f_name1)        
        
        sec_list = pd.read_excel(output_path_s3+r'{}'.format(file), sheet_name = f_name1)
        
        writer = pd.ExcelWriter(output_path_s3+'{} - Total data.xlsx'.format(f_name2), engine='xlsxwriter')
        sec_list.to_excel(writer, sheet_name = str(f_name2))
        writer.save()
        
        os.remove(output_path_s3+file)
        print("Renamed {} and Removed Old file for {}".format(f_name1,f_name2))        
        
# NS_with_BO should match with tickers_yahoo_a2done (Just a formality check)
if len(NS_with_BO_3) == len(tickers_yahoo_s2done):
    print("All Done :)")
else:
    print("Check NS_with_BO and tickers_yahoo_a2done")

# Now move to Missing Tickers
    
#t_end1 = time.time()
#t_data = t_end1 - t_start
#print(t_data)

