
# Stock Analysis and Recommendation

The deliverables of this project are:

- Stock Markets Sector Analysis using sentiment analysis to find emerging investable sectors for the future.
- Perform stock analysis within each sector based on fundamental analysis, and generate investment advice across portfolios.
- Use Technical analysis to monitor the stocks within a portfolio to generate buy/sell indicator for a particular stock, at a particular time.
- Create a portfolio from a basket of stocks and use optimization methods to maximize returns by optimal asset allocation.

## Data Source

- Yahoo Finance
- TickerTape
- Trendlyne
- Value Research

## Tools used:

- MS Excel
- Python
- Jupyter
- Power Automate

## Procedure:

### Part 1: Stock Research and Sector Analysis

- Create a screener on TickerTape for the universe of stocks you want to perform comparitive analysis on.
- Add filters for the data points you need for your analysis.
- There are 11 sectors defined by Global Industry Classification Standard (GICS), namely: 
    - Energy
    - Materials
    - Industrials
    - Utilities
    - Healthcare
    - Financials
    - Consumer Discretionary
    - Consumer Staples
    - Information Technology
    - Communication Services
    - Real Estate
- Download the data from TickerTape for category and store it in csv files in a folder.

- Create a screener on Trendlyne for a similar universe (doen't matter if they are exactly the same, a join operation later can solve any redundancy).
- Add required metrics for analysis.
- Download the data for the sector provided there and store it in a folder.
Note: TickerTape provides the GICS compliant sectors, whereas Trendlyne doesn't, so the Data from Trendlyne has to be categorized, based on the category provided for the same data point (stock) on TickerTape.

- Run the 'database creation.ipynb' file to merge the data from the multiple sources to create a main database for further analysis.
- Some debugging might be required if any column might be missing or an extra. Sector 'Analysis Debug.ipynb' might come in handy for that.
- Once the dataset is ready and compliant, run the 'Sector Automate.py' file which categorizes the data into different dataframes based on their sector, and stores it into an excel workbook, with a new sheet for every sector.
- The stocks under each sector are then further categorized based on their Market Cap, as stocks under different market cap levels shouldn't be compared.
- A comparitive analysis is done for each stock within its market cap, within its sector and are ranked accordingly. The best ones are showed at the top of their table, followed by the lower ranks.


### Part 2: Portfolio Management and Optimization

- Create an excel workbook. Create a sheet named Universe, which would contain important data about the stocks in the portfolio/s, i.e. Stock Name from Value Research, Industry, Ticker, Beta, Current Price. (these 4 can be obtained using Excel Stock Data).
- Calculate Bollinger Bands and SuperTrend Indicator using Moving Average for universe of stocks, or get them using Trendlyne, and concatenate it to the current universe.
- Create a Benchmark sheet which would contain information about the market benchmark names and their tickers, along with the risk-free rate of returns. These data points are useful while performing benchmarking for portfolio performance.
- Create a Client sheet, which would contain the Names of the different portfolios, along with other important information like their AUM, Portfolio returns and performance metrics.
- Create a portfolio template on excel, with the following required data points:
    - Name
    - Start-Date
    - End-Date
    - Benchmark Ticker Symbol (to extract its data)
    - Names of company stocks in the portfolio
      
- Use VLOOKUP to get the Stock industry, ticker, beta, current market price and category from the Universe Sheet.
- Now this can go two ways:

    #### If you have a stock tradebook (80% excel - 20% python):
    - Upload your tradebook to value research. Select the start and end date you have set in your excel sheet, and download the stock performance excel sheet.
    - Get the stock names from Value Research and paste them in the Company Name column in your sheet. The other required data will get pulled by VLOOKUP. 
    - You can get the number of current invested shares by dividing the Current Investment Value for a stock by its Current Market Price.
    - Copy the Invested Amount, Current Amount, Actual and Absolute Return columns from the file and paste it in your sheet. You will see majority of the calculations and analysis ready for the portfolio.
    - Save the file and close it.
    - Run the 'Sector Automate.py' file, which will update the benchmark return for your specified time frame and also run a portfolio optimization code with defined bounds and constraints.
    - THe Optimization code will suggest optimized weights for the current portfolio stock, with higher returns. lower beta and a maximized Treynor Ratio.
    - One can also look at the expected return column, calculated using the CAPM model, which can suggest if a particular stock has realized its expected gains.
 
  #### If you want to create a portfolio from scratch (80% python - 20% excel):
    -  Get the tickers of the stock you want to invest in, and the market benchmark for the portfolio.
    -  Get the stock industry, beta from ezxel stock data function, and also their investment categories.
    -  Select the period of investment, and the corresponding risk-free rate od return, and also the number of initial shares allocation for each stock.
    -  Run the 'Portfolio Optimization.ipynb' for the above data. The code downloads stock price history data from yahoofinance, and calculates stock return and volatility.
    -  Based on given and extracted data, it calculates important portfolio statistics, and prints them out in a table below the portfolio.
    -  An optimization code block optimizes the portfolio by maximing the sharpe, treynor, and alpha in different scenarios, and gives an asset allocation strategy for each case.
 
- The portfolios created by either of the two methods provide an efficient way to manage multiple portfolios at once, and is useful to track performance and investment strategies.


## Conclusion
- The two parts of the project aim to provide a complete guide to an investor about emerging sectors to invest in, how to select the best stocks amongst thousands of stocks, and select those which might suit their investment criteria and requirment.
- It also helps them to track a portfolio's performance, and update the allocation according to current market scenario, in order to maximize gains and minimize portfolio risk.
