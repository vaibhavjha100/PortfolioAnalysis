"""
Module for preprocessing processed tradebooks.
"""

import pandas as pd
import numpy as np
import yfinance as yf
import os
import config as cfg


def eligible_securities(type):
    """
    Function to get eligible securities based on the type of tradebook.
    An eligible security is one whose data is available through yfinance.
    Function returns the tradebook which only contains eligible securities.
    """
    type = type.lower()
    if type == "zerodha":
        tb = pd.read_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))
        # Add .NS suffix to all tickers
        tb['symbol'] = tb['symbol'].apply(lambda x: x + ".NS")
        # Filter for eligible securities using check_yf_availability function
        tickers = tb['symbol'].unique()
        eligible_tickers, ineligible_tickers = check_yf_availability(tickers)
        tb = tb[tb['symbol'].isin(eligible_tickers)]
        print(f"The following tickers are excluded from the analysis: {ineligible_tickers}")
        tb.to_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))

def check_yf_availability(tickers):
    """
    Function to check which tickers are available through yfinance.
    Available tickers are returned as a list.
    Non-available tickers are printed to the console.
    """
    available_tickers = []
    unavailable_tickers = []
    for ticker in tickers:
        try:
            yf.Ticker(ticker).info
            available_tickers.append(ticker)
        except Exception as e:
            unavailable_tickers.append(ticker)
            print(f"Ticker {ticker} is not available: {e}")
    return available_tickers, unavailable_tickers

def reindex_tradebooks(type):
    """
    Fucntion to reindex tradebooks based on order execution time.
    """
    type = type.lower()
    if type == "zerodha":
        tb = pd.read_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))
        tb['order_execution_time'] = pd.to_datetime(tb['order_execution_time'])
        tb.set_index('order_execution_time', inplace=True)
        tb.sort_index(inplace=True)
        tb.to_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))

def preprocess_tradebooks(type):
    """
    Function to preprocess tradebooks based on the type of tradebook.
    This function will call the eligible_securities and reindex_tradebooks functions.
    """
    eligible_securities(type)
    reindex_tradebooks(type)

if __name__ == "__main__":
    pass
    #preprocess_tradebooks("zerodha")