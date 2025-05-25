"""
Configuration file for the project.
This file contains directory paths and formats for different stock brokers.
"""

import os

# Defining directory paths
ROOTDIR = os.path.join(os.path.expanduser("~"), "PycharmProjects", "PortfolioAnalysis")
RAWDATADIR = os.path.join(ROOTDIR, "Tradebooks", "Raw")
DATADIR = os.path.join(ROOTDIR, "Tradebooks", "Processed")
EXCELDIR = os.path.join(ROOTDIR, "Excel")

# Defining formats for different stock brokers
STANDARD_FORMAT = ['time', 'date', 'ticker', 'exchange', 'quantity', 'price']
ZERODHA_FORMAT = ['symbol', 'isin', 'trade_date', 'exchange', 'segment', 'series',
       'trade_type', 'auction', 'quantity', 'price', 'trade_id', 'order_id',
       'order_execution_time']