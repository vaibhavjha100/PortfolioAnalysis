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

    if broker == 'zerodha':
        # Load Zerodha tradebook data
        df = pd.read_csv(os.path.join(cfg.DATADIR, 'zerodha.csv'))

        # If start date is not provided, set it to the earliest date in the data
        if start_date is None:
            start_date = df['trade_date'].min()

        # Filter data based on start and end dates
        df = df[(df['trade_date'] >= start_date) & (df['trade_date'] <= end_date)]

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

        for symbol in df['symbol'].unique():
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
            if pf.empty:
                pf = data
            else:
                pf = pd.merge(pf, data, on='Date', how='outer')
        # Set the date as index
        pf.set_index('Date', inplace=True)

        # Create a new column for every column
        # in the dataframe and set the value to 0
        # column name is column with suffix _weight
        for col in pf.columns:
            pf[col + '_weight'] = 0

        # Iterate through the df (tradebook) in order to get the weights
        # for each stock on each date
        # If stock is bought then the weight is increased by the quantity
        # If stock is sold then the weight is decreased by the quantity

if __name__ == "__main__":
    # Example usage
    fund = construct_fund('zerodha', start_date='2021-11-29', end_date='2023-09-11')
    print(fund.head())
    print(fund.info())