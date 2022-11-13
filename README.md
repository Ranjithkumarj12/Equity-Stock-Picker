# Equity-Stock-Picker
This is an Investing tool that I built for assisting me with my Equity portoflio investments.

It analyses around 3500+ companies that are listed on NSE & BSE, and selects top performing securities on the basis of 3 parameters.

The 3 parameters are - 3-yr CAGR of Z-Scores of Fundamental Ratios, Short term Stock Return (from last annual reporting rate till present date), and % Shares held by Insiders/Promoters. These 3 parameters are weighted 55%, 30%, and 15%, respectively.

Ratios Used for Parameter 1= ['Gross Profit Margin',	'Operating Profit Margin',	'Pre-tax Margin',	'Net Profit Margin','Return on Assets (Added Gr.Interest)',	'Operating Return on Assets',	'Return on Total Capital',	'Return on Equity', 'Inventory Turnover',	'Days of Inventory on hand', 'Total Asset Turnover',	'Fixed Asset Turnover',	'Working Capital Turnover',	'Current Ratio', 'Cash Ratio',	'Debt-to-Equity',	'Debt-to-Capital Ratio',	'Debt-to-Assets',	'Financial Leverage',	'Interest Coverage (Income)',	'CFO-to-Net Revenue',	'CFO-to-Assets',	'CFO-to-Equity',	'CFO-to-Op.Income',	'CFO-to-Debt',	'Interest Coverage (CFO)', 'Earnings to Price','Sales to Price',	'CFO to Price',	'BV to Price',	'FCFF to Price']

The main objective of the code is not to provide you with a final list of securities to invest, but rather, to assist you in narrowing down the universe of securities to just a handful of securities for you to research on. Therefore, it is strongly recommended to conduct a business level due-dillengence of the companies which the code recommends. 

The code is entirely based on web-scrapped data from Yahoo Finance, and therefore, would have to be modified if there are any changes to the UI of Yahoo Finances' website.

The order for executing the codes is as follows: Annual Data > Summary Data > Missing Tickers > Daily Prices > Analysis Part 1 - A > Analysis Part 1 - B > Industry Stats > Analysis Part 1 - C > Analysis Part 2 > Analysis Part 3 > Analysis Part 4 > Analysis Part

Please use the code only for educational purposes. The user of the code shall be completely responsible for any investment decision that is made, basis the output of the code. The creator of the code shall not be held responsbile for any financial loss or damage(s) that the user of the code may face.  
