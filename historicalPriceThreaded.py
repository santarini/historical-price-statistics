import os
import csv
import itertools
import datetime as dt
import pandas_datareader.data as web
import threading
import time

start_time = time.time()

def getHistoricalData(startTicker, endTicker):
    #create source folder if it doesnt exist yet
    if not os.path.exists('stock_dfs'):
        os.makedirs('stock_dfs')

    #open csv
    with open("SandPCopy.csv") as csvfile:
        reader = csv.DictReader(csvfile)

        #note: using j+1 in itertools.islice to make it inclusive of endTicker
        for row in itertools.islice(reader, startTicker, endTicker+1):
            ticker = (row['Ticker'])
            print("Getting " + ticker)
            if not os.path.exists('stock_dfs/{}.csv'.format(ticker)):
                start = dt.datetime(2010,1,1)
                end = dt.datetime(2018,1,1)
                df = web.DataReader(ticker, 'morningstar', start, end)
                df.to_csv('stock_dfs/{}.csv'.format(ticker))
            else:
                print('Already have {}'.format(ticker))



#dj thread that shit
downloadThreads = []
for i in range(1, 479, 50):
    time.sleep(.5)
    downloadThread = threading.Thread(target=getHistoricalData, args=(i, i + 49))
    downloadThreads.append(downloadThread)
    downloadThread.start()

elapsed_time = time.time() - start_time
