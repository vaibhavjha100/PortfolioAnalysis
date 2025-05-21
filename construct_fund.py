"""
Module to construct a personal fund from processed tradebooks.
"""

import os
import pandas as pd
import numpy as np
import yfinance as yf
import config as cfg
from datetime import date

def construct_fund(type, start_date=None, end_date=date.today()):
    """
    Constructs a personal fund based type of broker, start date, and end date.
    By default, start date is None and end date is today.

    Personal fund is a dataframe with the following columns:
    - NAV: Net Asset Value
    - Composition: The composition of the fund
    """
    type = type.lower()
    pf = pd.DataFrame()

    if type == 'zerodha':
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
        for symbol in df['symbol'].unique():
            data = yf.download(symbol, start=start_date, end=end_date, multi_level_index=False)
            stock_data[symbol] =data
        return stock_data
        # Stuck on problem on how to handle data for stocks that are merged

if __name__ == "__main__":
    # Example usage
    fund = construct_fund('zerodha', start_date='2021-11-29', end_date='2023-09-11')
    print(fund)