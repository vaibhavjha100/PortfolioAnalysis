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
import preprocessing as ppc

def get_tradebook(*broker, start_date=None, end_date=date.today()):
    """
    Function to get the tradebook for a specific broker.
    By default, start date is None and end date is today.
    """
    if not broker:
        broker = ('zerodha',)
    tb = pd.read_csv(os.path.join(cfg.DATADIR, "tradebook.csv"))
    tb.set_index('time', inplace=True)
    tb['date'] = pd.to_datetime(tb['date'])
    # If start date is not provided, set it to the earliest date in the data
    if start_date is None:
        start_date = tb['date'].min()
    start_date = pd.Timestamp(start_date)
    end_date = pd.Timestamp(end_date)
    # Filter the tradebook based on start and end dates
    tb = tb[(tb['date'] >= start_date) & (tb['date'] <= end_date)]
    # Convert start_date to string format YYYY-MM-DD
    start_date = start_date.strftime('%Y-%m-%d')

    return tb, start_date

def get_stock_data(tickers, start_date=None, end_date=date.today()):
    """
    Function to get stock data for a list of tickers.
    By default, start date is None and end date is today.
    """
    if start_date is None:
        print("Function cannot be used without start date.")
        return None
    end_date = end_date.strftime('%Y-%m-%d')
    stock_data = {}
    ineligible_tickers = []
    app = xw.App(visible=False)
    wb = app.books.open(os.path.join(cfg.EXCELDIR, 'Stockdata.xlsx'))
    sheet = wb.sheets[0]
    start_year = start_date.split("-")[0]
    start_month = start_date.split("-")[1]
    start_day = start_date.split("-")[2]
    end_year = end_date.split("-")[0]
    end_month = end_date.split("-")[1]
    end_day = end_date.split("-")[2]
    for ticker in tickers:
        # Clear the sheet before writing new data
        sheet.clear()
        cell = sheet.range("A1")
        cell.value = ticker
        cell2 = sheet.range("A2")
        cell2.formula2 = f'=STOCKHISTORY(A1, DATE({start_year}, {start_month}, {start_day}), DATE({end_year}, {end_month}, {end_day}), 0)'
        # Sleep for 5 seconds to allow the formula to calculate
        time.sleep(2.5)
        # Get all the data from the sheet into a dataframe
        data = sheet.range("A2").expand().options(pd.DataFrame).value
        if data.empty:
            print(f"No data found for ticker: {ticker}")
            ineligible_tickers.append(ticker)
            continue
        # Rename the columns to the ticker symbol
        data.rename(columns={data.columns[0]: ticker}, inplace=True)
        stock_data[ticker] = data
    wb.close()
    app.quit()

    sd = pd.DataFrame()

    # Merge all the dataframes into one with the symbol as the column name
    for symbol, data in stock_data.items():
        data.index = pd.to_datetime(data.index)
        data['Date'] = data.index
        # Reindex the data to remove date as index
        data.reset_index(drop=True, inplace=True)
        if sd.empty:
            sd = data
        else:
            sd = pd.merge(sd, data, on='Date', how='outer')
    # Set the date as index
    sd.set_index('Date', inplace=True)

    # Delete columns with no data
    sd.dropna(axis=1, how='all', inplace=True)

    # Create a new column for every column
    # in the dataframe and set the value to 0
    # column name is column with suffix _weight
    for col in sd.columns:
        sd[col + '_weight'] = np.nan

    return sd, ineligible_tickers

def calculate_nav(stock_data, tradebook, ineligible_tickers=[]):
    """
    Function to calculate the Net Asset Value (NAV) of a personal fund based on stock data and tradebook.
    """
    if stock_data.empty or tradebook.empty:
        print("Stock data or tradebook is empty. Cannot calculate NAV.")
        return None

    # Drop ineligible tickers from stock_data and tradebook
    if ineligible_tickers:
        stock_data = stock_data.drop(columns=[ticker for ticker in ineligible_tickers if ticker in stock_data.columns], errors='ignore')
        tradebook = tradebook[~tradebook['ticker'].isin(ineligible_tickers)]

    # Iterate through the df (tradebook) in order to get the weights
    # for each stock on each date
    # If stock is traded then the weight is increased by the quantity
    # If stock is not traded then the weight is not changed
    for index, row in tradebook.iterrows():
        ticker = row['ticker']
        quantity = row['quantity']
        date = row['date']
        # Add the quantity to the weight of the stock on that date if weight was not nan
        # If weight was nan, then set it to the quantity
        if pd.isna(stock_data.loc[date, ticker + '_weight']):
            stock_data.loc[date, ticker + '_weight'] = quantity
        else:
            stock_data.loc[date, ticker + '_weight'] += quantity
        # For all other tickers if there was a weight on previous date, then set it to previous weight
        # If there was no previous weight, then set it to 0
        # Use iloc to get the previous date. If its first row, then set it to 0
        for col in stock_data.columns:
            if col.endswith('_weight') and col != ticker + '_weight':
                if pd.isna(stock_data.loc[date, col]):
                    if date == stock_data.index[0]:
                        stock_data.loc[date, col] = 0
                    else:
                        stock_data.loc[date, col] = stock_data.iloc[stock_data.index.get_loc(date) - 1][col]
        # Fill forward the weights for nan values
        stock_data.fillna(method='ffill', inplace=True)

    price_cols =  stock_data.columns[~stock_data.columns.str.endswith('_weight')]

    # aum is sumproduct of price and weight for each stock from pf
    aum = []
    for date, row in stock_data.iterrows():
        aum_value = 0
        for cols in price_cols:
            weight_col = cols + '_weight'
            aum_value += row[cols] * row[weight_col]
        if aum_value == 0 and len(aum) > 0:
            aum_value = aum[-1]
        aum.append(aum_value)

    initial_nav = 100
    nav = [initial_nav]
    units = [aum[0] / initial_nav]
    result = pd.DataFrame({
        'AUM': aum
    }, index=stock_data.index)

    for date, row in stock_data.iterrows():
        # Skip the first date as it is the initial date
        if date == stock_data.index[0]:
            result.loc[date, 'Units'] = units[0]
            result.loc[date, 'NAV'] = result.loc[date, 'AUM'] / units[0]
            continue
        prev_units = units[-1]
        prev_nav = nav[-1]
        if date in tradebook['date'].values:
            # If the tradebook has a record for this date, add units

            if tradebook.loc[date, 'quantity'] != 0:
                units.append(prev_units + (tradebook.loc[date, 'quantity'] * tradebook.loc[date, 'price'] / prev_nav))
            else:
                units.append(prev_units)

        # Append AUM/units to nav
        nav.append(result.loc[date, 'AUM'] / units[-1])
        # Add NAV and units to result dataframe for this date
        result.loc[date, 'NAV'] = nav[-1]
        result.loc[date, 'Units'] = units[-1]

    # Merge the NAV with the stock data
    final_result = pd.merge(result, stock_data, left_index=True, right_index=True, how='outer')

    return final_result

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
        # Set order execution time as index
        tf.set_index('order_execution_time', inplace=True)

        # If start date is not provided, set it to the earliest date in the data
        if start_date is None:
            start_date = tf['trade_date'].min()

        # Filter data based on start and end dates
        tf = tf[(tf['trade_date'] >= start_date) & (tf['trade_date'] <= end_date)]

        # Download stock data of all the stocks in the tradebook
        # starting from the start date to the end date
        # Store in dictionary: stock_data
        stock_data = {}
        app = xw.App(visible=False)
        wb = app.books.open(os.path.join(cfg.EXCELDIR, 'Stockdata.xlsx'))
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
        app.quit()
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
                sf = pd.merge(sf, data, on='Date', how='outer')
        # Set the date as index
        sf.set_index('Date', inplace=True)

        # Create a new column for every column
        # in the dataframe and set the value to 0
        # column name is column with suffix _weight
        for col in sf.columns:
            sf[col + '_weight'] = np.nan

        # Iterate through the df (tradebook) in order to get the weights
        # for each stock on each date
        # If stock is bought then the weight is increased by the quantity
        # If stock is sold then the weight is decreased by the quantity

        for index, row in tf.iterrows():
            symbol = row['symbol']
            quantity = row['quantity']
            date = row['trade_date']
            if row['trade_type'] == 'buy':
                sf.loc[date, symbol + '_weight'] += quantity
            elif row['trade_type'] == 'sell':
                sf.loc[date, symbol + '_weight'] -= quantity
        # Fill forward the weights. All the weights are 0 if nothing is filled
        sf.fillna(method='ffill', inplace=True)
        # Fill backward to 0
        sf.fillna(0, inplace=True)

        price_cols = sf.columns[~sf.columns.str.endswith('_weight')]
        # aum is sumproduct of price and weight for each stock from pf
        aum = []
        for date, row in sf.iterrows():
            aum_value = 0
            for cols in price_cols:
                weight_col = cols + '_weight'
                aum_value += row[cols] * row[weight_col]
            if aum_value == 0 and len(aum) > 0:
                aum_value = aum[-1]
            aum.append(aum_value)

        initial_nav = 100
        nav = [initial_nav]
        units = [aum[0] / initial_nav]
        result = pd.DataFrame({
            'AUM': aum
        }, index=sf.index)

        for date, row in sf.iterrows():
            # Skip the first date as it is the initial date
            if date == sf.index[0]:
                result.loc[date, 'Units'] = units[0]
                result.loc[date, 'NAV'] = result.loc[date, 'AUM'] / units[0]
                continue
            prev_units = units[-1]
            prev_nav = nav[-1]
            if date in tf['trade_date'].values:
                if tf.loc[date, 'trade_type'] == 'buy':
                    units.append(prev_units + (tf.loc[date, 'quantity'] * tf.loc[date, 'price'] / prev_nav))
                elif tf.loc[date, 'trade_type'] == 'sell':
                    units.append(prev_units - (tf.loc[date, 'quantity'] * tf.loc[date, 'price'] / prev_nav))
                else:
                    units.append(prev_units)
            # Append AUM/units to nav
            nav.append(result.loc[date, 'AUM'] / units[-1])
            # Add NAV and units to result dataframe for this date
            result.loc[date, 'NAV'] = nav[-1]
            result.loc[date, 'Units'] = units[-1]

        # Merge the NAV with the stock data
        pf = pd.merge(result, sf, left_index=True, right_index=True, how='outer')

        return pf

if __name__ == "__main__":
    # Example usage
    tradebook, start_date = get_tradebook('zerodha')
    stock_data, iticks = get_stock_data(tradebook['ticker'].unique(), start_date=start_date)
    nav_data = calculate_nav(stock_data, tradebook, iticks)
    print(nav_data.head())
    print(nav_data.info())

