# -*- coding: utf-8 -*-
"""
Created on Fri Nov 25 17:02:56 2022

@author: jrkumar

"""
#Importing Libraries
import pandas as pd
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
#from datetime import date

ticker_path = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Analysis Part 3 (Yahoo)\\')
output_path_q = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
input_path_q = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

# Getting list of tickerS
ticker_df = pd.read_excel(ticker_path+r'All Ticker Z-Score CAGR and Ranking.xlsx', sheet_name=0)
ticker_list = ticker_df['Tickers'].dropna().values.tolist()
tickers_yahoo_q1 = ticker_list

#For tickers which got processed
tickers_yahoo_done1 = []
tickers_yahoo_done1_df = pd.DataFrame()

#For tickers which didnt get through
tickers_na_yahoo_q1 = [] 
tickers_na_yahoo_q1_df = pd.DataFrame()

#For tickers which don't have recent quarter results
#tickers_na_rctqtr_yahoo_q1 = [] 
#tickers_na_rctqtr_yahoo_q1_df = pd.DataFrame()

#For tickers which don't have quarter results of y-1
#tickers_na_yminus1qtr_yahoo_q1 = [] 
#tickers_na_yminus1qtr_yahoo_q1_df = pd.DataFrame()

#For tickers which don't have uniform years across all 3 FS
#tickers_non_uniform_a1 = []
#tickers_non_uniform_a1_df = pd.DataFrame()

#Adjust rest time to ensure scapping of complete data (Currenly set for Optimum tunrover time)
rest = 20
rest_popup = 1
rest_qtr = 2

#Measure Start time
#t_start = time.time()   

#Initiating Combined Dictionary to store data of all tickers
combined_dict1 = {}

#Path to store extracted data
output_path_a = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Quarterly Financials (Yahoo)\\')
ticker_count = len(tickers_yahoo_q1)
count = 1

date_collector_df = pd.DataFrame()
# Part - 1:
#Round 1 - Annual Data Web Scrapping Process - Using Selenium to interact, and Beautiful Soup to parse
for t in tickers_yahoo_q1:
    date_collector = []    
    try:
        #Creating empty dictionary to store data from all 3 FS
        temp_dir = {}
        
        #Driver - Income Statement
        url_inc = 'https://finance.yahoo.com/quote/{}/financials?p={}'.format(t,t)
        
        table_title_inc = {}
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(options=options)
        driver.get(url_inc)
        time.sleep(rest)
        try:
            driver.find_element("xpath",'//*[@id="myLightboxContainer"]/section/button[1]').click() 
            time.sleep(rest_popup)
        except:
            pass
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()
        time.sleep(rest_qtr)
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
           
        page_content_inc = driver.page_source
        driver.quit()
        
        soup_inc = bs(page_content_inc,'html.parser')
        table_data_inc = soup_inc.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_inc:
            heading_inc = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_inc in heading_inc:
                table_title_inc[top_row_inc.get_text(separator = '|').split('|')[0]] = top_row_inc.get_text(separator = '|').split('|')[1:]
        
        inc_header = 'Income Statement'
        date_collector.append(inc_header)
        inc_list = table_title_inc[list(table_title_inc.keys())[0]]
        for i in range(len(inc_list)):
            date_collector.append(inc_list[i])
        print('Income Statment for Ticker:{} is complete'.format(t)) 
        
        #Income Statement Scrapping 
        for i in table_data_inc:
            rows_inc = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
            for row_inc in rows_inc:
                x = row_inc.get_text(separator = '|').split('|')[1:]
                for m in range(len(x)):
                    x[m] = x[m].replace(",","")
                    if '-' in x[m]:
                        if len(x[m]) == 1:
                            x[m] = x[m].replace("-","")
                    x[m] = x[m].replace("k","E+03")
                y = row_inc.get_text(separator = '|').split('|')[0]
                temp_dir[y] = x
        
        print('Income Statment for Ticker:{} is complete'.format(t)) 
        
        #Converting to Dataframe
        temp_df = pd.DataFrame(temp_dir, index = table_title_inc[list(table_title_inc.keys())[0]])

        #Saving to excel
        writer = pd.ExcelWriter(output_path_q+'{} - Quarter data (Income Statement).xlsx'.format(t), engine='xlsxwriter')
        temp_df.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        #Driver - Balance Sheet
        temp_dir = {}
        url_bs = 'https://finance.yahoo.com/quote/{}/balance-sheet?p={}'.format(t,t)
        
        table_title_bs = {}
            
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(options=options)
        driver.get(url_bs)
        time.sleep(rest)
        driver.find_element("xpath",'//*[@id="myLightboxContainer"]/section/button[1]').click() 
        time.sleep(rest_popup)
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()
        time.sleep(rest_qtr)
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
        page_content_bs = driver.page_source
        driver.quit()
        
        soup = bs(page_content_bs,'html.parser')
        table_data_bs = soup.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_bs:
            heading_bs = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_bs in heading_bs:
                table_title_bs[top_row_bs.get_text(separator = '|').split('|')[0]] = top_row_bs.get_text(separator = '|').split('|')[1:]
        
        bs_header = 'Balance Sheet'
        date_collector.append(bs_header)
        bs_list = table_title_bs[list(table_title_bs.keys())[0]]
        for i in range(len(bs_list)):
            date_collector.append(bs_list[i])
        print('Driver - Balance Sheet for Ticker:{} is complete'.format(t))
        
        #Balance Sheet Scrapping            
        for i in table_data_bs:
            rows_bs = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
            for row_bs in rows_bs:    
                x = row_bs.get_text(separator = '|').split('|')[1:]
                for m in range(len(x)):
                    x[m] = x[m].replace(",","")
                    if '-' in x[m]:
                        if len(x[m]) == 1:
                            x[m] = x[m].replace("-","")
                    x[m] = x[m].replace("k","E+03")
                y = row_bs.get_text(separator = '|').split('|')[0]
                temp_dir[y] = x
        
        print('Balance Sheet for Ticker:{} is complete'.format(t))
        
        #Converting to Dataframe
        temp_df = pd.DataFrame(temp_dir, index = table_title_bs[list(table_title_bs.keys())[0]])

        #Saving to excel
        writer = pd.ExcelWriter(output_path_q+'{} - Quarter data (Balance Sheet).xlsx'.format(t), engine='xlsxwriter')
        temp_df.to_excel(writer, sheet_name = str(t))
        writer.save()
        
        #Driver - Cash flow Statement
        temp_dir = {}
        url_cf = 'https://finance.yahoo.com/quote/{}/cash-flow?p={}'.format(t,t)
        
        table_title_cf = {}
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(options=options)
        driver.get(url_cf)
        time.sleep(rest)
        driver.find_element("xpath",'//*[@id="myLightboxContainer"]/section/button[1]').click() 
        time.sleep(rest_popup)
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()
        time.sleep(rest_qtr)
        driver.find_element("xpath",'//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
        page_content_cf = driver.page_source
        driver.quit()
            
        soup = bs(page_content_cf,'html.parser')
        table_data_cf = soup.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_cf:
            heading_cf = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_cf in heading_cf:
                table_title_cf[top_row_cf.get_text(separator = '|').split('|')[0]] = top_row_cf.get_text(separator = '|').split('|')[1:]
        
        cf_header = 'Cash Flow Statement'
        date_collector.append(cf_header)
        cf_list = table_title_cf[list(table_title_cf.keys())[0]]
        for i in range(len(cf_list)):
            date_collector.append(cf_list[i])
        print('Driver - Cash flow Statment for Ticker:{} is complete'.format(t))
        
        #Cashflow Statement Scrapping
        for i in table_data_cf:
            rows_cf = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
            for row_cf in rows_cf:    
                x = row_cf.get_text(separator = '|').split('|')[1:]
                for m in range(len(x)):
                    x[m] = x[m].replace(",","")
                    if '-' in x[m]:
                        if len(x[m]) == 1:
                            x[m] = x[m].replace("-","")
                    x[m] = x[m].replace("k","E+03")
                y = row_cf.get_text(separator = '|').split('|')[0]
                temp_dir[y] = x
        
        print('Cashflow Statement for Ticker:{} is complete'.format(t))
        
        #Converting to Dataframe
        temp_df = pd.DataFrame(temp_dir, index = table_title_cf[list(table_title_cf.keys())[0]])

        #Saving to excel
        writer = pd.ExcelWriter(output_path_q+'{} - Quarter data (Cash Flow Statement).xlsx'.format(t), engine='xlsxwriter')
        temp_df.to_excel(writer, sheet_name = str(t))
        writer.save()
                 
        
        print('Ticker:{} is complete and Dataframe saved as excel file'.format(t))
        
        # Date collector dataframe
        date_collector_temp_df = pd.DataFrame(date_collector, columns = [t])
        date_collector_df = pd.concat([date_collector_df,date_collector_temp_df],axis=1)
        
        #Saving to excel
        writer = pd.ExcelWriter(input_path_q+'Date Collector.xlsx', engine='xlsxwriter')
        date_collector_df.to_excel(writer)
        writer.save()
        
        #Appending tickers that are done
        tickers_yahoo_done1.append(t)
        tickers_yahoo_done1_tempdf = pd.DataFrame(tickers_yahoo_done1, columns = ['Ticker Name'])
        tickers_yahoo_done1_df = pd.concat([tickers_yahoo_done1_df,tickers_yahoo_done1_tempdf],axis=1)
        tickers_yahoo_done1_df = tickers_yahoo_done1_df.loc[:,~tickers_yahoo_done1_df.columns.duplicated(keep = 'last')]
        
        #Saving to excel
        writer = pd.ExcelWriter(input_path_q+'Quarter_Done.xlsx', engine='xlsxwriter')
        tickers_yahoo_done1_df.to_excel(writer)
        writer.save()
        
        print("Remaining Tickers: {}".format(str(ticker_count - count)))
        count = count + 1
        
    #Appending Tickers which didn't get through    
    except:
        tickers_na_yahoo_q1.append(t)
        tickers_na_yahoo_q1_tempdf = pd.DataFrame(tickers_na_yahoo_q1, columns = ['Ticker Name'])
        tickers_na_yahoo_q1_df = pd.concat([tickers_na_yahoo_q1_df,tickers_na_yahoo_q1_tempdf],axis=1)
        tickers_na_yahoo_q1_df = tickers_na_yahoo_q1_df.loc[:,~tickers_na_yahoo_q1_df.columns.duplicated(keep = 'last')]
        
        #Saving to excel
        writer = pd.ExcelWriter(input_path_q+'Quarter_Not_Done.xlsx', engine='xlsxwriter')
        tickers_na_yahoo_q1_df.to_excel(writer)
        writer.save()
        
        print("Remaining Tickers: {}".format(str(ticker_count - count)))
        count = count + 1

# Now move to Yahoo Summary Data




    
        
        
        
    
    
    
    
    