import csv
import os
import datetime as dt
import pandas_datareader.data as web

#create source folder if it doesnt exist yet
if not os.path.exists('stock_dfs'):
    os.makedirs('stock_dfs')

#open csv
with open("sandp500.csv") as csvfile:
    reader = csv.DictReader(csvfile)

    for row in reader:
        ticker = (row['ticker'])
        if not os.path.exists('stock_dfs/{}.csv'.format(ticker)):
            start = dt.datetime(2010,1,1)
            end = dt.datetime(2018,1,1)
            #use 'morningstar' for stocks
            df = web.DataReader('AAPL', 'morningstar', start, end)
            #use 'stooq' for indexes no dates necessary
            #df = web.DataReader('^DJI', 'stooq')
            df.to_csv('stock_dfs/{}.csv'.format(ticker))
            #you can also print these to test the program instead of going head first into csv
            #print(df.head())
        else:
            print('Already have {}'.format(ticker))
