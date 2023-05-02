# Import Packages
import numpy as np
import pandas as pd
import pandas_datareader as pdr
import yfinance as yf


startDate = pd.to_datetime('2020-02-04')

#tesla_df = pdr.data.get_data_yahoo('TSLA', start=startDate)

msft = yf.Ticker('MSFT')

t = yf.download('VOO')
t2 = yf.download('GOOG', group_by='tickers')

tickers = yf.Tickers('msft aapl goog')

nasdaqSymbols = pdr.get_nasdaq_symbols()

print('Complete')
