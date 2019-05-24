# taken directly from https://medium.com/python-data/effient-frontier-in-python-34b0c3043314
# import needed modules
import quandl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from pandas_datareader import data as web
import openpyxl




# get adjusted closing prices of 5 selected companies with Quandl
NUM_TRADING_DAYS = 252
start_date = '2014-1-1'
end_date = '2019-05-18'
start_time = pd.to_datetime(start_date)
date_analysis = pd.to_datetime(pd.date_range(start=start_time
                                             + pd.DateOffset(months=12),
                                             end=pd.to_datetime(end_date)
                                             , freq="M"))
# select the tickers
#selected = ['QQQ', 'AMZN', 'PLUG', 'VTI', 'VB']
#selected = ['CNP', 'F', 'WMT', 'GE', 'TSLA']
selected = ['VTI', 'COKE', 'AMZN', 'QQQ', 'CMG', 'MMM', 'SBUX', 'CELG',
            'WMT', 'LVS', 'GE', 'HON']
data_source = input('Select Data Source: [1]=Quandl, [2]=Yahoo')
if data_source == '1':
    quandl.ApiConfig.api_key = input('Type in Quandl API key')
    
    data = quandl.get_table('WIKI/PRICES', ticker = selected,
                            qopts = { 'columns': ['date', 'ticker',
                                                  'adj_close'] },
    date = { 'gte': '2018-1-1', 'lte': '2018-12-31' }, paginate=True)
    # reorganise data pulled by setting date as index with
    # columns of tickers and their corresponding adjusted prices
    clean = data.set_index('date')
    table = clean.pivot(columns='ticker')

elif data_source == '2':
    data = pd.DataFrame()

    for i in selected:
     data[i] = web.DataReader(i, 'yahoo', start='2014-1-1'
         , end='2018-5-18')['Adj Close']
    table = data
    
min_vol_results = pd.DataFrame()
max_sharpe_results = pd.DataFrame()

# set the number of combinations for imaginary portfolios
for analysis_date in date_analysis:
    
    #table = data.loc[analysis_date-pd.DateOffset(months=12):analysis_date]
    table = data.loc[pd.to_datetime(pd.Timestamp(start_date)):analysis_date]
    
    num_assets = len(selected)
    num_portfolios = 50000

    # set the risk free rate
    rf_rate = 0



    # calculate daily and annual returns of the stocks
    returns_daily = table.pct_change()
    returns_annual = returns_daily.mean() * NUM_TRADING_DAYS

    # get daily and covariance of returns of the stock
    cov_daily = returns_daily.cov()
    cov_annual = cov_daily * NUM_TRADING_DAYS

    # empty lists to store returns, volatility and weights of imiginary portfolios
    port_returns = []
    port_volatility = []
    sharpe_ratio = []
    stock_weights = []

    # populate the empty lists with each portfolios returns,risk and weights
    for single_portfolio in range(num_portfolios):
        weights = np.random.random(num_assets)
        weights /= np.sum(weights)
        returns = np.dot(weights, returns_annual)
        volatility = np.sqrt(np.dot(weights.T, np.dot(cov_annual, weights)))
        sharpe = (returns - rf_rate) / volatility
        sharpe_ratio.append(sharpe)
        port_returns.append(returns)
        port_volatility.append(volatility)
        stock_weights.append(weights)

    # a dictionary for Returns and Risk values of each portfolio
    portfolio = {'Returns': port_returns,
                 'Volatility': port_volatility,
                 'Sharpe Ratio': sharpe_ratio}

    # extend original dictionary to accomodate each ticker and 
    # weight in the portfolio
    for counter,symbol in enumerate(selected):
        portfolio[symbol+' Weight'] = [Weight[counter] for Weight in stock_weights]
    
    # make a nice dataframe of the extended dictionary
    df = pd.DataFrame(portfolio)
    
    # get better labels for desired arrangement of columns
    column_order = (['Returns', 'Volatility', 'Sharpe Ratio']
                  + [stock+' Weight' for stock in selected])

    # reorder dataframe columns
    df = df[column_order]



    # use the min, max values to locate and create the two special portfolios
    #df=df[df['Returns']>=0.15]
    #df=df[df['Returns']<=0.16]

    # find min Volatility & max sharpe values in the dataframe (df)
    min_volatility = df['Volatility'].min()
    max_sharpe = df['Sharpe Ratio'].max()

    sharpe_portfolio = df.loc[df['Sharpe Ratio'] == max_sharpe]
    min_variance_port = df.loc[df['Volatility'] == min_volatility]

    # plot the efficient frontier with a scatter plot
    plt.style.use('seaborn')
    df.plot.scatter(x='Volatility', y='Returns', c='Sharpe Ratio',
                    cmap='RdYlGn', edgecolors='black', figsize=(10, 8), grid=True)
    plt.scatter(x=sharpe_portfolio['Volatility'], y=sharpe_portfolio['Returns'],
                c='red', marker='D', s=200)
    plt.scatter(x=min_variance_port['Volatility'], y=min_variance_port['Returns'],
                c='blue', marker='D', s=200 )
    plt.xlabel('Volatility (Std. Deviation)')
    plt.ylabel('Expected Returns')
    plt.title('Efficient Frontier')
    plt.show()
    
    min_vol_results = min_vol_results.append(min_variance_port,ignore_index=True)
    max_sharpe_results = max_sharpe_results.append(sharpe_portfolio, ignore_index=True)

    print(min_variance_port.T)
    print(sharpe_portfolio.T)
    
min_vol_results.set_index(date_analysis,inplace=True)
min_vol_results.to_excel('MinVolPortfolioResults.xlsx')

max_sharpe_results.set_index(date_analysis,inplace=True)
max_sharpe_results.to_excel('MaxSharpePortfolioResults.xlsx')