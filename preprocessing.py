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


def eligible_securities(type, start_date=None, end_date=None):
    """
    Function to get eligible securities based on the type of tradebook.
    An eligible security is one whose data is available through yfinance.
    Function returns the tradebook which only contains eligible securities.
    """
    type = type.lower()
    if type == "zerodha":
        tb = pd.read_csv(os.path.join(cfg.DATADIR, "zerodha.csv"))
        # Filter for eligible securities using check_yf_availability function
        tickers = tb['symbol'].unique()
        eligible_tickers, ineligible_tickers = check_stockhistory_availability(tickers, start_date, end_date)
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

def preprocess_tradebooks(type, start_date=None, end_date=None):
    """
    Function to preprocess tradebooks based on the type of tradebook.
    This function will call the eligible_securities and reindex_tradebooks functions.
    """
    eligible_securities(type, start_date, end_date)
    reindex_tradebooks(type)

if __name__ == "__main__":
    pass
    #preprocess_tradebooks("zerodha")