# -*- coding: utf-8 -*-
"""
Created on Mon Apr 11 14:50:25 2022

@author: jrkumar
"""
#Importing Libraries
import pandas as pd
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

input_path_a = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Inputs - Generic\\')

ticker_list1 = pd.read_excel(input_path_a+r'Total Ticker List - Annual.xlsx', sheet_name=0)
ticker_list1.dropna(axis = 0,inplace = True)
ticker_list1 = ticker_list1[ticker_list1['Total Ticker List - Annual'] !=  'NULL']
tickers1 = ticker_list1['Total Ticker List - Annual'].tolist()

#Total Ticker List
tickers_yahoo_a1 = tickers1
#For tickers which got processed
tickers_yahoo_done1 = []
tickers_yahoo_done1_df = pd.DataFrame()
#For tickers which don't end with ".NS"
tickers_na_yahoo_a1 = []
tickers_na_yahoo_a1_df = pd.DataFrame()
#For tickers which don't have uniform years across all 3 FS
tickers_non_uniform_a1 = []
tickers_non_uniform_a1_df = pd.DataFrame()

#Adjust rest time to ensure scapping of complete data (Currenly set for Optimum tunrover time)
rest = 20
rest_popup = 1

#Measure Start time
#t_start = time.time()   

#Initiating Combined Dictionary to store data of all tickers
combined_dict1 = {}

#Path to store extracted data
output_path_a = (r'C:\Users\jrkumar\OneDrive - Indxx\Desktop\Algo Trading Engine\Ticker Data - Financials (Yahoo)\\')
ticker_count = len(tickers_yahoo_a1)
count = 1

# Part - 1:
#Round 1 - Annual Data Web Scrapping Process - Using Selenium to interact, and Beautiful Soup to parse
for t in tickers_yahoo_a1:
    
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
            driver.find_element_by_xpath('//*[@id="myLightboxContainer"]/section/button[1]').click() 
            time.sleep(rest_popup)
        except:
            pass
        driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
           
        page_content_inc = driver.page_source
        driver.quit()
        
        soup_inc = bs(page_content_inc,'html.parser')
        table_data_inc = soup_inc.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_inc:
            heading_inc = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_inc in heading_inc:
                table_title_inc[top_row_inc.get_text(separator = '|').split('|')[0]] = top_row_inc.get_text(separator = '|').split('|')[2:]
        
        print('Driver - Income Statment for Ticker:{} is complete'.format(t)) 
        
        #Driver - Balance Sheet
        url_bs = 'https://finance.yahoo.com/quote/{}/balance-sheet?p={}'.format(t,t)
        
        table_title_bs = {}
            
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(options=options)
        driver.get(url_bs)
        time.sleep(rest)
        try:
            driver.find_element_by_xpath('//*[@id="myLightboxContainer"]/section/button[1]').click() 
            time.sleep(rest_popup)
        except:
            pass
        driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
        page_content_bs = driver.page_source
        driver.quit()
        
        soup = bs(page_content_bs,'html.parser')
        table_data_bs = soup.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_bs:
            heading_bs = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_bs in heading_bs:
                table_title_bs[top_row_bs.get_text(separator = '|').split('|')[0]] = top_row_bs.get_text(separator = '|').split('|')[1:]
        
        print('Driver - Balance Sheet for Ticker:{} is complete'.format(t))
        
        #Driver - Cash flow Statement
        url_cf = 'https://finance.yahoo.com/quote/{}/cash-flow?p={}'.format(t,t)
        
        table_title_cf = {}
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(options=options)
        driver.get(url_cf)
        time.sleep(rest)
        try:
            driver.find_element_by_xpath('//*[@id="myLightboxContainer"]/section/button[1]').click() 
            time.sleep(rest_popup)
        except:
            pass
        driver.find_element_by_xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()
        page_content_cf = driver.page_source
        driver.quit()
            
        soup = bs(page_content_cf,'html.parser')
        table_data_cf = soup.find_all('div', attrs = {'class':'M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)'})
        for i in table_data_cf:
            heading_cf = i.find_all('div', attrs = {'class':'D(tbr) C($primaryColor)'})
            for top_row_cf in heading_cf:
                table_title_cf[top_row_cf.get_text(separator = '|').split('|')[0]] = top_row_cf.get_text(separator = '|').split('|')[2:]
        
        print('Driver - Cash flow Statment for Ticker:{} is complete'.format(t))
        
        #Identifying minimum no. of years (to ensure uniform no. of years across all 3 FS)
        yr_inc = len(list(table_title_inc.values())[0])
        yr_bs = len(list(table_title_bs.values())[0])
        yr_cf = len(list(table_title_cf.values())[0])
        min_yrs = min(yr_inc,yr_bs,yr_cf)
        
        if yr_inc == min_yrs:
            table_title_bs['Breakdown'] = table_title_bs['Breakdown'][0:min_yrs]
            table_title_cf['Breakdown'] = table_title_cf['Breakdown'][0:min_yrs]
            years = table_title_inc['Breakdown'].copy()
            years2 = table_title_inc['Breakdown'].copy()
            for i in range(len(years2)):
                years2[i] = years2[i][-4:]
        elif yr_bs == min_yrs:
            table_title_inc['Breakdown'] = table_title_inc['Breakdown'][0:min_yrs]
            table_title_cf['Breakdown'] = table_title_cf['Breakdown'][0:min_yrs]
            years = table_title_bs['Breakdown'].copy()
            years2 = table_title_bs['Breakdown'].copy()
            for i in range(len(years2)):
                years2[i] = years2[i][-4:]
        else:
            table_title_inc['Breakdown'] = table_title_inc['Breakdown'][0:min_yrs]
            table_title_bs['Breakdown'] = table_title_bs['Breakdown'][0:min_yrs]
            years = table_title_cf['Breakdown'].copy()
            years2 = table_title_cf['Breakdown'].copy()
            for i in range(len(years2)):
                years2[i] = years2[i][-4:]
        
        #Proceed only if years of all 3FS match
        if (table_title_inc['Breakdown'] == table_title_bs['Breakdown']) & (table_title_bs['Breakdown'] == table_title_cf['Breakdown']) & (table_title_cf['Breakdown'] == table_title_inc['Breakdown']):
        
            #Income Statement Scrapping 
            for i in table_data_inc:
                rows_inc = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
                for row_inc in rows_inc:
                    x = row_inc.get_text(separator = '|').split('|')[2:(min_yrs+2)]
                    for m in range(len(x)):
                        x[m] = x[m].replace(",","")
                        if '-' in x[m]:
                            if len(x[m]) == 1:
                                x[m] = x[m].replace("-","")
                        x[m] = x[m].replace("k","E+03")
                    y = row_inc.get_text(separator = '|').split('|')[0]
                    temp_dir[y] = x
            
            print('Income Statment for Ticker:{} is complete'.format(t))        
            
            #Balance Sheet Scrapping            
            for i in table_data_bs:
                rows_bs = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
                for row_bs in rows_bs:    
                    x = row_bs.get_text(separator = '|').split('|')[1:(min_yrs+1)]
                    for m in range(len(x)):
                        x[m] = x[m].replace(",","")
                        if '-' in x[m]:
                            if len(x[m]) == 1:
                                x[m] = x[m].replace("-","")
                        x[m] = x[m].replace("k","E+03")
                    y = row_bs.get_text(separator = '|').split('|')[0]
                    temp_dir[y] = x
            
            print('Balance Sheet for Ticker:{} is complete'.format(t))
            
            #Cashflow Statement Scrapping
            for i in table_data_cf:
                rows_cf = i.find_all('div', attrs = {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
                for row_cf in rows_cf:    
                    x = row_cf.get_text(separator = '|').split('|')[2:(min_yrs+2)]
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
            temp_df = pd.DataFrame(temp_dir, index = years2)
            temp_df['Rep_Date'] = 'temp'
            for i in range(len(temp_df)):
                temp_df['Rep_Date'][i] = years[i] 
            #Saving to excel
            writer = pd.ExcelWriter(output_path_a+'{} - Annual data.xlsx'.format(t), engine='xlsxwriter')
            temp_df.to_excel(writer, sheet_name = str(t))
            writer.save()
                
            combined_dict1[t] = temp_dir
            
            print('Ticker:{} is complete and Dataframe saved as excel file'.format(t))
            tickers_yahoo_done1.append(t)
            tickers_yahoo_done1_tempdf = pd.DataFrame(tickers_yahoo_done1, columns = ['Ticker Name'])
            tickers_yahoo_done1_df = pd.concat([tickers_yahoo_done1_df,tickers_yahoo_done1_tempdf],axis=1)
            tickers_yahoo_done1_df = tickers_yahoo_done1_df.loc[:,~tickers_yahoo_done1_df.columns.duplicated(keep = 'last')]
            
            #Saving to excel
            writer = pd.ExcelWriter(input_path_a+'Annual_Done.xlsx', engine='xlsxwriter')
            tickers_yahoo_done1_df.to_excel(writer)
            writer.save()
            
            print("Remaining Tickers: {}".format(str(ticker_count - count)))
            count = count + 1
            
        else:
            tickers_non_uniform_a1.append(t)
            tickers_non_uniform_a1_tempdf = pd.DataFrame(tickers_non_uniform_a1, columns = ['Ticker Name'])
            tickers_non_uniform_a1_df = pd.concat([tickers_non_uniform_a1,tickers_non_uniform_a1_tempdf],axis=1)
            tickers_non_uniform_a1_df = tickers_non_uniform_a1_df.loc[:,~tickers_non_uniform_a1_df.columns.duplicated(keep = 'last')]
            
            #Saving to excel
            writer = pd.ExcelWriter(input_path_a+'Annual_Non_Uniform.xlsx', engine='xlsxwriter')
            tickers_non_uniform_a1_df.to_excel(writer)
            writer.save()
            
            print("Remaining Tickers: {}".format(str(ticker_count - count)))
            count = count + 1
        
    #Appending Tickers which don't end with ".NS"    
    except:
        tickers_na_yahoo_a1.append(t)
        tickers_na_yahoo_a1_tempdf = pd.DataFrame(tickers_na_yahoo_a1, columns = ['Ticker Name'])
        tickers_na_yahoo_a1_df = pd.concat([tickers_na_yahoo_a1_df,tickers_na_yahoo_a1_tempdf],axis=1)
        tickers_na_yahoo_a1_df = tickers_na_yahoo_a1_df.loc[:,~tickers_na_yahoo_a1_df.columns.duplicated(keep = 'last')]
        
        #Saving to excel
        writer = pd.ExcelWriter(input_path_a+'Annual_Not_Done.xlsx', engine='xlsxwriter')
        tickers_na_yahoo_a1_df.to_excel(writer)
        writer.save()
        
        print("Remaining Tickers: {}".format(str(ticker_count - count)))
        count = count + 1

# Now move to Yahoo Summary Data



    
        
        
        
    
    
    
    
    