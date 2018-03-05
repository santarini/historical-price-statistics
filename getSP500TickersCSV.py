import requests
import bs4 as bs
import csv

def getSP500TickersCSV():
    resp = requests.get('http://en.wikipedia.org/wiki/List_of_S%26P_500_companies')
    soup = bs.BeautifulSoup(resp.text, 'lxml')
    table = soup.find('table', {'class': 'wikitable sortable'})
    tickers = []
    for row in table.findAll('tr')[1:]:
        ticker = row.findAll('td')[0].text
        tickers.append(ticker)

    with open('SandP500.csv', 'a') as csvfile:
        fieldnames = ['Ticker']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator = '\n')
        writer.writeheader()
        for x in tickers:
            writer.writerow({'Ticker': x})
    print(tickers)

getSP500TickersCSV()
