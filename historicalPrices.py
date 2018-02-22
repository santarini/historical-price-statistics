import datetime as dt
import pandas as pd
import pandas_datareader.data as web

start = dt.datetime(2010,1,1)
end = dt.datetime(2018,1,1)

#use 'stooq' for indexes no dates necessary
#df = web.DataReader('^DJI', 'stooq')
#use 'morningstar' for stocks

df = web.DataReader('AAPL', 'morningstar', start, end)


#print(df.head())
df.to_csv('AAPL.csv')
