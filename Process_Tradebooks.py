"""
Module reads all tradebooks from the RAWDATADIR directory, categorizes them, processes them, and saves the processed
data to the DATADIR directory.
"""

import os
import pandas as pd
import config as cfg

def load_tradebooks():
    """
    Load tradebooks from the RAWDATADIR directory.
    """
    tradebooks = []
    for filename in os.listdir(cfg.RAWDATADIR):
        if filename.endswith(".csv"):
            filepath = os.path.join(cfg.RAWDATADIR, filename)
            tradebook = pd.read_csv(filepath)
            tradebooks.append(tradebook)
    return tradebooks

def categorize_tradebooks():
    """
    Categorize tradebooks based on the broker. Check the columns and assign a category.
    Stock broker formats are defined in config.py.
    """
    tradebooks = load_tradebooks()
    tradebooks_brokers = []
    for i in tradebooks:
        if set(i.columns) == set(cfg.ZERODHA_FORMAT):
            tradebooks_brokers.append([i, "Zerodha"])
    return tradebooks_brokers

def process_tradebooks():
    """
    Process the tradebooks by creating different csv file formats for different stock brokers.
    Then modify the tradebooks to have date column as index and save them to the DATADIR directory.
    """
    tradebooks = categorize_tradebooks()
    zerodha = pd.DataFrame(columns=cfg.ZERODHA_FORMAT)
    for tradebook, broker in tradebooks:
        if broker == "Zerodha":
            zerodha = pd.concat([zerodha, tradebook])

    # For Zerodha, make the trade_date column the index and sort the DataFrame
    zerodha['trade_date'] = pd.to_datetime(zerodha['trade_date'])
    zerodha.set_index('trade_date', inplace=True)
    zerodha.sort_index(inplace=True)
    # Save the processed DataFrame to the DATADIR directory
    zerodha.to_csv(os.path.join(cfg.DATADIR, "zerodha.csv"), index=True)

if __name__ == "__main__":
    pass
    # process_tradebooks()

