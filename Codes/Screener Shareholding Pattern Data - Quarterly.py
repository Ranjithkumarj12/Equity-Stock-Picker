# -*- coding: utf-8 -*-
"""
Created on Sun Dec  4 16:14:51 2022

@author: jrkumar
"""

#Importing Libraries
import pandas as pd
import requests as rq
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import os
#import re

#Measure Start time
#t_start = time.time()   

#Path to store extracted data
output_path_sp = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Shareholding Pattern (Screener)\\')
ticker_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
input_path_q = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')
dwd_path_temp = (r'C:\Users\jrkumar\Downloads\Ticker Data - Quarterly Shareholding Pattern (Screener)\\')

# Getting list of tickerS
ticker_df = pd.read_excel(ticker_path+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
ticker_list = ticker_df['Tickers_New'].dropna().values.tolist()
for i in range(len(ticker_list)):
   ticker_list[i] = ticker_list[i].replace(".NS", "")
   ticker_list[i] = ticker_list[i].replace(".BO", "")

#Adjust rest time to ensure scapping of complete data (Currenly set for Optimum tunrover time)
rest = 1

tickers_yahoo_sp1 = ticker_list

#tickers_yahoo_sp1 = []
#tickers_yahoo_sp2 = []
#for file1 in os.listdir(output_path_sp):
#    f_name1 = file1.replace(" - Shareholding Pattern Data.xlsx","")
#    tickers_yahoo_sp2.append(f_name1)
    
#for i in ticker_list:
#    if i not in tickers_yahoo_sp2:
#        tickers_yahoo_sp1.append(i)

#For tickers which got processed
tickers_yahoo_done1 = []
tickers_yahoo_done1_df = pd.DataFrame()

#For tickers which didnt get through
tickers_na_yahoo_sp1 = [] 
tickers_na_yahoo_sp1_df = pd.DataFrame()

# Web Scrapping Process - Using Beautiful Soup to parse
tickers_na_yahoo_sp1 = tickers_yahoo_sp1
while  tickers_na_yahoo_sp1 != []:
    ticker_count = len(tickers_yahoo_sp1)
    count = 1
    for t in tickers_yahoo_sp1:   
        try:
            url_sp = 'https://www.screener.in/company/{}/consolidated/'.format(t)
            
            options = Options()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            driver = webdriver.Chrome(options=options)
            driver.get(url_sp)
            time.sleep(rest)
            page_content_sp = driver.page_source
            driver.quit()
            
            #headers = {"User-Agent" : "Chrome/100.0.4896.75"}
            #page_sp = rq.get(url_sp, headers = headers)
            #page_content_sp = page_sp.content
            soup_sp = bs(page_content_sp,'html.parser')
            table_data_sp = soup_sp.find_all('table', attrs = {'class':'data-table'})
            table_data_sp2 = table_data_sp[6]
            
            headings_sp = table_data_sp2.find_all('button', attrs = {'class':'button-plain'})
            headings = []
            for heading in headings_sp:
                y = heading.get_text(separator = '|').split('|')[0].replace(" ", "")
                y = y.replace("\n","")
                y = y.replace(" ", "")
                y = y.replace("\xa0","")
                headings.append(y)
            cnt_row = len(headings)
            
            # We take only the last 3 Quarters of Shareholding Pattern data
            headings_sp2 = table_data_sp2.find_all('th')
            headings2 = []
            for heading2 in headings_sp2:
                z = heading2.get_text(separator = '|').split('|')[0]
                headings2.append(z)
            headings2 = headings2[-3:]
            
            rows_sp2 = table_data_sp2.find_all('tr')
            rows = []
            for row in range(len(rows_sp2)):
                row_data = []
                for i in rows_sp2[row]:
                    x = i.get_text(separator = '|').split('|')[0]
                    row_data.append(x)
                while "\n" in row_data:
                    row_data.remove("\n")
                row_data = row_data[-3:]
                rows.append(row_data)
                rows = rows[-cnt_row:]
            
            sp_df = pd.DataFrame(rows, index = headings, columns = headings2)
            
            #Saving to Excel
            writer = pd.ExcelWriter(output_path_sp+'{} - Shareholding Pattern Data.xlsx'.format(t), engine='xlsxwriter')
            sp_df.to_excel(writer, sheet_name = str(t))
            writer.save()
            
            print('Shareholding Pattern for Ticker:{} is complete and saved as excel'.format(t))
            
            tickers_yahoo_done1.append(t)
            
            if t in tickers_na_yahoo_sp1:
                tickers_na_yahoo_sp1.remove(t)
                
        except:
            try:
                url_sp = 'https://www.screener.in/company/{}/consolidated/'.format(t)
                
                options = Options()
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
                driver = webdriver.Chrome(options=options)
                driver.get(url_sp)
                time.sleep(rest)
                page_content_sp = driver.page_source
                driver.quit()
                
                #headers = {"User-Agent" : "Chrome/100.0.4896.75"}
                #page_sp = rq.get(url_sp, headers = headers)
                #page_content_sp = page_sp.content
                soup_sp = bs(page_content_sp,'html.parser')
                table_data_sp = soup_sp.find_all('table', attrs = {'class':'data-table'})
                table_data_sp2 = table_data_sp[5]
                
                headings_sp = table_data_sp2.find_all('button', attrs = {'class':'button-plain'})
                headings = []
                for heading in headings_sp:
                    y = heading.get_text(separator = '|').split('|')[0].replace(" ", "")
                    y = y.replace("\n","")
                    y = y.replace(" ", "")
                    y = y.replace("\xa0","")
                    headings.append(y)
                cnt_row = len(headings)
                
                # We take only the last 3 Quarters of Shareholding Pattern data
                headings_sp2 = table_data_sp2.find_all('th')
                headings2 = []
                for heading2 in headings_sp2:
                    z = heading2.get_text(separator = '|').split('|')[0]
                    headings2.append(z)
                headings2 = headings2[-3:]
                
                rows_sp2 = table_data_sp2.find_all('tr')
                rows = []
                for row in range(len(rows_sp2)):
                    row_data = []
                    for i in rows_sp2[row]:
                        x = i.get_text(separator = '|').split('|')[0]
                        row_data.append(x)
                    while "\n" in row_data:
                        row_data.remove("\n")
                    row_data = row_data[-3:]
                    rows.append(row_data)
                    rows = rows[-cnt_row:]
                
                sp_df = pd.DataFrame(rows, index = headings, columns = headings2)
                
                #Saving to Excel
                writer = pd.ExcelWriter(output_path_sp+'{} - Shareholding Pattern Data.xlsx'.format(t), engine='xlsxwriter')
                sp_df.to_excel(writer, sheet_name = str(t))
                writer.save()
                
                print('Shareholding Pattern for Ticker:{} is complete and saved as excel'.format(t))
                
                tickers_yahoo_done1.append(t)
                
                if t in tickers_na_yahoo_sp1:
                    tickers_na_yahoo_sp1.remove(t)
            except:
                if t not in tickers_na_yahoo_sp1:
                    tickers_na_yahoo_sp1.append(t)
      
        print("Remaining Tickers: {}".format(str(ticker_count - count)))
        count = count + 1
    
    tickers_yahoo_sp1 = tickers_na_yahoo_sp1

# Change file name back to original
#ticker_df.reset_index(inplace = True, drop = True)
temp = []
for file in os.listdir(dwd_path_temp):
    try:
        file = file.replace(" - Shareholding Pattern Data.xlsx","")
        for tick in range(len(ticker_df['Tickers_New'])):
            if ticker_df['Tickers_New'].iloc[tick] == file:
                file_new = ticker_df['Tickers'].iloc[tick]
                file_new = file_new + ' - Shareholding Pattern Data.xlsx'
        os.rename(dwd_path_temp + '{} - Shareholding Pattern Data.xlsx'.format(file), dwd_path_temp + file_new)
        
        # Removing Analysis Part 1 for these tickers
        print("Renamed Ticker from {} to {}".format(file+' - Shareholding Pattern Data.xlsx',file_new))   
    except:
        temp.append(file)

# Now move to Analysis 6
