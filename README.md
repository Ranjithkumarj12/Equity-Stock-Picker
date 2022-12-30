# Equity-Stock-Picker
This is an Investing tool that I built for assisting me with my Equity portoflio investments.

It analyses around 3500+ companies that are listed on NSE & BSE, and selects top performing securities on the basis of 4 parameters.

The 4 parameters are - Recent 3-yr CAGR of Z-Scores of Annual Fundamental Ratios, Recent 3-qtr CAGR of Z-Scores of Quarterly Fundamental Ratios , YoY recent qtr CAGR of Z-Scores of Quarterly Fundamental Ratios, and recent 3-qtr change in shareholding pattern of FIIs/DIIs/Promoters/Public. These 3 parameters are weighted 70%, 16%, 4%, and 10%, respectively.

Ratios Used for Parameter 1 are:

Net Profit Margin	- 8.33%
Operating ROA	- 8.33%
Gross Profit Margin	- 4.17%
Op. Profit Margin	- 4.17%
Total Asset Turnover - 5.00%
Working Capital Turnover	- 5.00%
Current Ratio	- 2.50%
Cash Ratio - 2.50%
Debt-Equity Ratio - 8.33%
Debt-Capital Ratio - 8.33%
Interest Coverage Ratio - 4.17%
CFO Interest Coverage	- 4.17%
CFO to Net Revenue (Gross Profit) -	6.67%
CFO to Avg Total Assets	- 6.67%
CFO to Avg Total Equity	- 3.33%
Return on Equity - 3.33%
Earnings to Price	- 5.00%
Book Value to Price	- 5.00%
FCFF to Price	- 5.00%

Ratios Used for Parameter 2 & 3 are:
Gross Profit Margin - 25%
Operating Profit Margin - 25%
Net Profit Margin - 50%

The main objective of the code is not to provide you with a final list of securities to invest, but rather, to assist you in narrowing down the universe of securities to just a handful of securities for you to research on. Therefore, it is strongly recommended to conduct a business level due-dillengence of the companies which the code recommends.

The code removes all financial services based companies from the consideration list, as the financial statements of these companies are organized differently. I will later on build a seperate framework to analyze financial services based companies. 

The code is entirely based on web-scrapped data from Yahoo Finance and Screener, and therefore, would have to be modified if there are any changes to the UI of Yahoo Finances' website. The code does a similar job as that of other freely available screeners like Ticker tape, screener, etc., but, the code was designed to offer the user a more flexible way to modify their investment criteria to pick stocks.

The order for executing the codes is as follows: Annual Data > Summary Data > Missing Tickers Part 1 > Daily Prices > Missing Tickers Part 2 > Analysis Part 1 - A > Analysis Part 1 - B > Industry Stats > Analysis Part 1 - C > Analysis Part 2 > Analysis Part 3 > Screener Shareholding Pattern Data > Analysis Part 6 > Analysis Part 7 > Analysis Part 8 > Analysis Part 9

Please use the code only for educational purposes. The user of the code shall be completely responsible for any investment decision that is made, basis the output of the code. The creator of the code shall not be held responsbile for any financial loss or damage(s) that the user of the code may face.  
