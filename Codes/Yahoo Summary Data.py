# -*- coding: utf-8 -*-
"""
Created on Wed Oct 12 12:59:45 2022

@author: jrkumar
"""

#Importing Libraries
import pandas as pd
import requests as rq
from bs4 import BeautifulSoup as bs
#import re

input_path_s = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

ticker_list1 = pd.read_excel(input_path_s+r'Total Ticker List - Summary.xlsx', sheet_name=0)
ticker_list1.dropna(axis = 0,inplace = True)
ticker_list1 = ticker_list1[ticker_list1['Total Ticker List - Summary'] !=  'NULL']
tickers1 = ticker_list1['Total Ticker List - Summary'].tolist()

#Input Tickers
tickers_yahoo_s1 = tickers1
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

ticker_count = len(tickers_yahoo_s1)
count = 1

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

    print("Remaining Tickers: {}".format(str(ticker_count - count)))
    count = count + 1
    
# Now move to Missing Tickers
    
#t_end1 = time.time()
#t_data = t_end1 - t_start
#print(t_data)

