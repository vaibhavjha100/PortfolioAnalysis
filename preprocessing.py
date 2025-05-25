"""
Module for preprocessing processed tradebooks.
"""

import pandas as pd
import numpy as np
import yfinance as yf
import os
import config as cfg
import xlwings as xw
import time
import Process_Tradebooks as pt

def standardize_tradebook_format(*brokers):
    """
    Function to standardize the tradebook format for different brokers.
    """
    pt.process_tradebooks()
    if not brokers:
        brokers = ("zerodha",)
    standard_tb = pd.DataFrame(columns=cfg.STANDARD_FORMAT)
    for broker in brokers:
        broker = broker.lower()
        if broker == "zerodha":
            tb = pd.read_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))
            # Check if standard dataframe is empty
            if standard_tb.empty:
                '''standard_tb['time'] = tb['order_execution_time']
                standard_tb.set_index('time', inplace=True)
                standard_tb['date'] = tb['trade_date']
                standard_tb['ticker'] = tb['symbol']
                standard_tb['exchange'] = tb['exchange']
                standard_tb['quantity'] = tb['quantity']
                standard_tb['price'] = tb['price']'''
                standard_tb = tb.copy()
                # Filter for needed columns
                standard_tb = standard_tb[['order_execution_time', 'trade_date', 'symbol', 'exchange', 'quantity', 'price']]
                # Rename columns to standard format
                standard_tb.rename(columns={
                    'order_execution_time': 'time',
                    'trade_date': 'date',
                    'symbol': 'ticker',
                    'exchange': 'exchange',
                    'quantity': 'quantity',
                    'price': 'price'
                }, inplace=True)
                standard_tb.set_index('time', inplace=True)
            else:
                # Append to the existing standard dataframe
                temp_tb = pd.DataFrame({
                    'time': tb['order_execution_time'],
                    'date': tb['trade_date'],
                    'ticker': tb['symbol'],
                    'exchange': tb['exchange'],
                    'quantity': tb['quantity'],
                    'price': tb['price']
                })
                temp_tb.set_index('time', inplace=True)
                standard_tb = pd.concat([standard_tb, temp_tb])
    # Save the standardized tradebook to a CSV file
    standard_tb.to_csv(os.path.join(cfg.DATADIR, "tradebook.csv"))


def eligible_securities(start_date=None, end_date=None):
    """
    Function to get eligible securities on the standard tradebook.
    An eligible security is one whose data is available through yfinance or excel.
    Function returns the tradebook which only contains eligible securities.
    """
    tb = pd.read_csv(os.path.join(cfg.DATADIR, "tradebook.csv"))
    tb.index = pd.to_datetime(tb['time'])
    tb['date'] = pd.to_datetime(tb['date'])
    # If start date is not provided, set it to the earliest date in the data
    if start_date is None:
        start_date = tb['date'].min().strftime('%Y-%m-%d')
    # If end date is not provided, set it to today's date
    if end_date is None:
        end_date = pd.to_datetime('today').strftime('%Y-%m-%d')
    # Filter for eligible securities using check_yf_availability function
    tickers = tb['ticker'].unique()
    eligible_tickers, ineligible_tickers = check_stockhistory_availability(tickers, start_date, end_date)
    tb = tb[tb['ticker'].isin(eligible_tickers)]
    print(f"The following tickers are excluded from the analysis: {ineligible_tickers}")
    tb.to_csv(os.path.join(cfg.DATADIR, "tradebook.csv"))

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

def check_stockhistory_availability(tickers, start_date=None, end_date=None):
    """
    Function to check if stock hisotry is there for the tickers through excel.
    """
    vba_injection()
    app = xw.App(visible=False)
    wb = app.books.open(os.path.join(cfg.EXCELDIR, 'test.xlsm'))
    valid_tickers = []
    invalid_tickers = []
    sheet= wb.sheets[0]

    for ticker in tickers:
        cell = sheet.range("A1")
        cell.value= ticker
        cell2 = sheet.range("A2")

        start_year = start_date.split("-")[0]
        start_month = start_date.split("-")[1]
        start_day = start_date.split("-")[2]
        end_year = end_date.split("-")[0]
        end_month = end_date.split("-")[1]
        end_day = end_date.split("-")[2]

        # Call stockhistory function from 2020 till date
        cell2.formula2 = f'=STOCKHISTORY(A1, DATE({start_year}, {start_month}, {start_day}), DATE({end_year}, {end_month}, {end_day}), 0)'
        # Sleep for 5 seconds to allow the formula to calculate
        time.sleep(5)
        # Check if C2 cell is empty or na
        if sheet.range("B3").value is None or pd.isna(sheet.range("B3").value):
            invalid_tickers.append(ticker)
        else:
            valid_tickers.append(ticker)
    wb.close()
    app.quit()

    return valid_tickers, invalid_tickers

def vba_injection():
    """
    Function to inject vba code into the excel file.
    VBA code is:
    Sub vba_test()
        Dim rng As Range
        Set rng = Range("A1") ' Update this to your target cell

        On Error Resume Next
        rng.Value = rng.Value ' Re-set value to trigger recognition
        rng.NumberFormat = "General" ' Avoid formatting issues
        rng.TextToColumns Destination:=rng ' Force Excel to reevaluate
        rng.ConvertToLinkedDataType xlLinkedDataTypeSourceAutomatic, "Stocks"
        On Error GoTo 0
    End Sub
    """
    app = xw.App(visible=False)
    wb = app.books.open(os.path.join(cfg.EXCELDIR, 'test.xlsm'))
    sheet = wb.sheets[0]
    vba_code = """Sub vba_test()
        Dim rng As Range
        Set rng = Range("A1") ' Update this to your target cell
    
        On Error Resume Next
        rng.Value = rng.Value ' Re-set value to trigger recognition
        rng.NumberFormat = "General" ' Avoid formatting issues
        rng.TextToColumns Destination:=rng ' Force Excel to reevaluate
        rng.ConvertToLinkedDataType xlLinkedDataTypeSourceAutomatic, "Stocks"
        On Error GoTo 0
    End Sub"""
    try:
        code_module = wb.api.VBProject.VBComponents("Sheet1").CodeModule
        code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(vba_code)
        wb.save()
        print("VBA code injected successfully.")
    except Exception as e:
        print(f"Error injecting VBA code: {e}")
    finally:
        wb.close()
        app.quit()

def preprocess_tradebooks(*brokers, start_date=None, end_date=None):
    """
    Function to preprocess tradebooks based on the type of tradebook.
    This function will call the eligible_securities and reindex_tradebooks functions.
    """
    if not brokers:
        brokers = ["zerodha"]
    # Convert brokers to list and lowercase
    brokers = [broker.lower() for broker in brokers]
    # Standardize the tradebook format
    standardize_tradebook_format(*brokers)
    # Get eligible securities
    eligible_securities(start_date, end_date)

if __name__ == "__main__":
    pass
    # preprocess_tradebooks("zerodha")