# -*- coding: utf-8 -*-
"""
Created on Sun Oct 30 17:53:43 2022

@author: jrkumar
"""

#Importing Libraries
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import fnmatch
import time

#Stored Path
ticker_input = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Inputs - Generic\\')

#Path to store extracted data
path_dwds = (r'C:\Users\jrkumar\Desktop\Algo Trading Engine\Ticker Data - Daily Prices (Yahoo)\\')

#Input Tickers
tickers_df = pd.read_excel(ticker_input+r'Total Ticker List - Summary.xlsx', sheet_name=0)
#Final Consideration List
final_tickers1 = tickers_df['Final Consideration List'].dropna().values.tolist()
final_tickers1 = final_tickers1[:33]
#Adjust rest time between scrapping
rest1 = 1

#Creating empty lists to collect tickers
tickers_yahoo_p1 = []
tickers_na_yahoo_p1 = []
#Downloading Excel files of Financials from Yahoo Finance
# Round 1 - For tickers which end with ".NS"
for t in final_tickers1:
    
    try:
        #Financial Statements Data Downloading
        url = 'https://finance.yahoo.com/quote/{}/history?p={}'.format(t,t)
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        
        params = {'behavior':'allow', 'downloadPath': path_dwds}
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd('Page.setDownloadBehavior', params)
        driver.get(url)
        time.sleep(rest1)
        
        # Click on Time Period
        driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div/div/span').click()
        time.sleep(rest1)
        # Click on 5-Yrs
        driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[3]/button/span').click()
        time.sleep(rest1)
        # Click on download
        driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[2]/span[2]/a/span').click()
        time.sleep(rest1)
        
        page_content = driver.page_source
        driver.quit()
        time.sleep(rest1)
        
        print('Downloaded Price data for Ticker: {}'.format(t))
        
        #Renaming the file (Also converting it from .csv to .xlsx)
        for file in os.listdir(path_dwds):
            if fnmatch.fnmatch(file, '{}.csv'.format(t)):
                f_name = file
        
        sec_list = pd.read_csv(path_dwds+r'{}'.format(f_name))
        f_name = f_name.replace(".csv","")
        
        # Writing to folder
        writer = pd.ExcelWriter(path_dwds+'{} - Daily Price data.xlsx'.format(f_name), engine='xlsxwriter')
        sec_list.to_excel(writer, sheet_name = str(f_name))
        writer.save()
        
        # Removing the Original File
        os.remove(path_dwds+f_name+".csv")
        
        # Appending to done tickers list
        tickers_yahoo_p1.append(f_name)
        
    except:
        tickers_na_yahoo_p1.append(t)

# Round 2 - for tickers which end with ".BO"
# Replacing back ".BO" with ".NS"
for i in range(len(tickers_na_yahoo_p1)):
    tickers_na_yahoo_p1[i] = tickers_na_yahoo_p1[i].replace(".NS",".BO")

tickers_yahoo_p2 = []
tickers_na_yahoo_p2 = []
   
for t in tickers_na_yahoo_p1:
    
    try:
    #Financial Statements Data Downloading
        url = 'https://finance.yahoo.com/quote/{}/history?p={}'.format(t,t)
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        
        params = {'behavior':'allow', 'downloadPath': path_dwds}
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd('Page.setDownloadBehavior', params)
        driver.get(url)
        time.sleep(rest1)
        
        # Click on Time Period
        driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div/div/span').click()
        time.sleep(rest1)
        # Click on 5-Yrs
        driver.find_element_by_xpath('//*[@id="dropdown-menu"]/div/ul[2]/li[3]/button/span').click()
        time.sleep(rest1)
        # Click on download
        driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[2]/span[2]/a/span').click()
        time.sleep(rest1)
        
        page_content = driver.page_source
        driver.quit()
        time.sleep(rest1)
        
        print('Downloaded Price data for Ticker: {}'.format(t))    
        
        #Renaming the file (Also converting it from .csv to .xlsx)
        for file in os.listdir(path_dwds):
            if fnmatch.fnmatch(file, '{}.csv'.format(t)):
                f_name = file
        
        sec_list = pd.read_excel(path_dwds+r'{}'.format(f_name))
        f_name = f_name.replace(".BO.csv",".NS")
        
        # Writing to folder
        writer = pd.ExcelWriter(path_dwds+'{} - Daily Price data.xlsx'.format(f_name), engine='xlsxwriter')
        sec_list.to_excel(writer, sheet_name = str(f_name))
        writer.save()
    
        # Removing the Original File
        f_name = f_name.replace(".NS",".BO")
        os.remove(path_dwds+f_name+".csv")
        
        # Appending to done tickers list
        f_name = f_name.replace(".BO",".NS")
        tickers_yahoo_p2.append(f_name)
        
    except:
        tickers_na_yahoo_p2.append(t)    

# Now move to Analysis Part 1 - A
    
    
    
