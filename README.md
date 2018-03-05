# Historical Price Statistics

<a href="https://github.com/santarini/pandas-datareader/blob/master/getSP500TickersCSV.py">getSP500TickersCSV.py</a> is a Python script that use Beautiful Soup to scrape S & P 500 Tickers from <a href="https://en.wikipedia.org/wiki/List_of_S%26P_500_companies">Wikipedia</a> and then creates a CSV of those tickers.

<a href="https://github.com/santarini/pandas-datareader/blob/master/historicalPrices.py">historicalPrices.py</a> is a Python script that pulls historical prices of companies from a specified period using <a href="https://pandas-datareader.readthedocs.io/en/latest/">Pandas Data Reader</a> and puts them into indvidual CSVs.

<a href="https://github.com/santarini/pandas-datareader/blob/master/Alpha.bas">Alpha.bas</a> is a Visual Basic script that populates a spreadsheet with data from external CSVs, scrubs and resamples the data, and then cacluates descriptive statistics for each.

## Table of Contents

* Alpha.bas
  * Populates a spreadsheet with data from other CSVs and run descriptive statistics on that data.
* README.md
  * Is what you're reading right meow.
* compileByPath.bas
  * The main routine of Alpha.bas
* createSummaryPage.bas
  * Function that creates a summary worksheet in an excel workbook
* formulaStructure.csv
  * Workbook used to help write all the formulas in this script
* getSP500Tickers.py
  * Python script that pulls S&P 500 tickers from wikipeida and puts them in an CSV
* historicalPrices.py
  *  Python script that pulls historical prices of companies using Pandas Data Reader and puts them into indvidual CSVs.
* manipulateData.bas
  * Resamples data to find changes between periods
* orderDataForGraphing.bas
  * Move columns around so that they can be graphed in Excel easily
* populateSummary.bas
  * Populates the Summary page with Descriptive Statistics taken from each imported CSV
* sandp500.csv
  * CSV containing all the constiuent tickers in the S & P 500
* scrubData.bas
  * Adds commas and dollar signs

## Prerequisites

Python

Microsoft Excel

Beautiful Soup

Pandas DataReader
