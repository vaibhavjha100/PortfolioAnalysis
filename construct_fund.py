"""
Module to construct a personal fund from processed tradebooks.
"""

import os
import pandas as pd
import numpy as np
import yfinance as yf
import config as cfg
from datetime import date
import xlwings as xw
import time
import matplotlib.pyplot as plt

def construct_fund(broker, start_date=None, end_date=date.today()):
    """
    Constructs a personal fund based type of broker, start date, and end date.
    By default, start date is None and end date is today.

    Personal fund is a dataframe with the following columns:
    - NAV: Net Asset Value
    - Composition: The composition of the fund
    """
    broker = broker.lower()
    pf = pd.DataFrame()
    sf = pd.DataFrame()

    if broker == 'zerodha':
        # Load Zerodha tradebook data
        tf = pd.read_csv(os.path.join(cfg.DATADIR, 'zerodha.csv'))

        # If start date is not provided, set it to the earliest date in the data
        if start_date is None:
            start_date = tf['trade_date'].min()

        # Filter data based on start and end dates
        tf = tf[(tf['trade_date'] >= start_date) & (tf['trade_date'] <= end_date)]

        # Download stock data of all the stocks in the tradebook
        # starting from the start date to the end date
        # Store in dictionary: stock_data
        stock_data = {}
        wb = xw.Book(os.path.join(cfg.EXCELDIR, 'test.xlsm'))
        sheet = wb.sheets[0]

        start_year = start_date.split("-")[0]
        start_month = start_date.split("-")[1]
        start_day = start_date.split("-")[2]
        end_year = end_date.split("-")[0]
        end_month = end_date.split("-")[1]
        end_day = end_date.split("-")[2]

        for symbol in tf['symbol'].unique():
            # Clear the sheet before writing new data
            sheet.clear()
            cell = sheet.range("A1")
            cell.value = symbol
            cell2 = sheet.range("A2")
            cell2.formula2 = f'=STOCKHISTORY(A1, DATE({start_year}, {start_month}, {start_day}), DATE({end_year}, {end_month}, {end_day}), 0)'
            # Sleep for 5 seconds to allow the formula to calculate
            time.sleep(2.5)
            # Get all the data from the sheet into a dataframe
            data = sheet.range("A2").expand().options(pd.DataFrame).value
            stock_data[symbol] =data
        wb.close()
        # Merge all the dataframes into one with the symbol as the column name
        for symbol, data in stock_data.items():
            data.columns = [symbol]
            data.index = pd.to_datetime(data.index)
            data['Date'] = data.index
            # Reindex the data to remove date as index
            data.reset_index(drop=True, inplace=True)
            if sf.empty:
                sf = data
            else:
                sf = pd.merge(pf, data, on='Date', how='outer')
        # Set the date as index
        sf.set_index('Date', inplace=True)

        # Create a new column for every column
        # in the dataframe and set the value to 0
        # column name is column with suffix _weight
        for col in sf.columns:
            pf[col + '_weight'] = 0

        # Iterate through the df (tradebook) in order to get the weights
        # for each stock on each date
        # If stock is bought then the weight is increased by the quantity
        # If stock is sold then the weight is decreased by the quantity


if __name__ == "__main__":
    # Example usage
    # fund = construct_fund('zerodha', start_date='2021-11-29', end_date='2023-09-11')
    # print(fund.head())
    # print(fund.info())
    tbd = {"date": ["2021-11-29", "2021-11-30", "2021-12-01", "2021-12-02", "2021-12-03"],
           "symbol": ["AAPL", "MSFT", "AAPL", "AAPL", "MSFT"],
           "trade_type": ["buy", "buy", "sell", "buy", "sell"],
           "quantity": [10, 5, 5, 10, 5],
           "price": [150, 200, 155, 157, 208]}
    tb = pd.DataFrame(tbd)
    tb['date'] = pd.to_datetime(tb['date'])
    tb.set_index('date', inplace=True)

    pfd = {"date": ["2021-11-29", "2021-11-30", "2021-12-01", "2021-12-02", "2021-12-03", "2021-12-04", "2021-12-05"],
           "AAPL": [150, 155, 160, 155, 145, 147, 150],
           "MSFT": [200, 205, 210, 205, 210, 220, 230],
           "AAPL_weight": [10, 10, 5, 15, 15, 15, 15],
           "MSFT_weight": [0, 5, 5, 5, 0, 0, 0]}
    pf = pd.DataFrame(pfd)
    pf['date'] = pd.to_datetime(pf['date'])
    pf.set_index('date', inplace=True)

    # Code Logic
    price_cols = pf.columns[~pf.columns.str.endswith('_weight')]
    # aum is sumproduct of price and weight for each stock from pf
    aum = []
    for date, row in pf.iterrows():
        aum_value = 0
        for cols in price_cols:
            weight_col = cols + '_weight'
            aum_value += row[cols] * row[weight_col]
        aum.append(aum_value)

    initial_nav = 100
    nav = [initial_nav]
    units = [aum[0] / initial_nav]
    result = pd.DataFrame({
        'AUM': aum
    }, index=pf.index)

    for date, row in pf.iterrows():
        # Skip the first date as it is the initial date
        if date == pf.index[0]:
            result.loc[date, 'Units'] = units[0]
            result.loc[date, 'NAV'] = result.loc[date, 'AUM']/ units[0]
            continue
        prev_units = units[-1]
        prev_nav = nav[-1]
        if date in tb.index:
            if tb.loc[date, 'trade_type'] == 'buy':
                units.append(prev_units + (tb.loc[date, 'quantity']*tb.loc[date, 'price'] / prev_nav))
            elif tb.loc[date, 'trade_type'] == 'sell':
                units.append(prev_units - (tb.loc[date, 'quantity']*tb.loc[date, 'price'] / prev_nav))
            else:
                units.append(prev_units)
        # Append AUM/units to nav
        nav.append(result.loc[date, 'AUM'] / units[-1])
        # Add NAV and units to result dataframe for this date
        result.loc[date, 'NAV'] = nav[-1]
        result.loc[date, 'Units'] = units[-1]

    print(result)
    print(result.info())
    plt.plot(result['NAV'])
    plt.title('NAV Over Time')
    plt.xlabel('Date')
    plt.ylabel('NAV')
    plt.show()
    # Need to double check the NAV calculation with excel