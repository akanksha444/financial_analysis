import pandas as pd
import os
import matplotlib.pyplot as plt
import glob
import yfinance as yf
import numpy as np


def get_ticker_price(ticker_symbol):
    ticker = yf.Ticker(ticker_symbol)
    ticker_info = ticker.history(period='1d')
    if not ticker_info.empty:
        return ticker_info['Close'][0]
    else:
        return None


company_ticker= yf.Ticker("wfc")

# get income statement
income_statement = company_ticker.income_stmt

# get balance sheet
balance_sheet = company_ticker.balance_sheet

# get cash flow
cash_flow = company_ticker.cashflow


income_statement_data = {
        'Total Revenue':  income_statement[income_statement.index.values == 'Total Revenue'].squeeze().tolist()[1:],
       # 'Non Interest Expense': income_statement[income_statement.index.values == 'Non Interest Expense'].squeeze().tolist()[1:],
        'Special Income Charges': income_statement[income_statement.index.values == 'Special Income Charges'].squeeze().tolist()[1:],
        'Pretax Income': income_statement[income_statement.index.values == 'Pretax Income'].squeeze().tolist()[1:],
        'Tax Provision': income_statement[income_statement.index.values == 'Tax Provision'].squeeze().tolist()[1:],
        'Net Income Common Stockholders': income_statement[income_statement.index.values == 'Net Income Common Stockholders'].squeeze().tolist()[1:],
        'Diluted NI Available to Com Stockholders': income_statement[income_statement.index.values == 'Diluted NI Availto Com Stockholders'].squeeze().tolist()[1:],
        'Basic EPS': income_statement[income_statement.index.values == 'Basic EPS'].squeeze().tolist()[1:],
        'Diluted EPS': income_statement[income_statement.index.values == 'Diluted EPS'].squeeze().tolist()[1:],
        'Basic Average Shares': income_statement[income_statement.index.values == 'Basic Average Shares'].squeeze().tolist()[1:],
        'Diluted Average Shares': income_statement[income_statement.index.values == 'Diluted Average Shares'].squeeze().tolist()[1:],
        'Net Interest Income': income_statement[income_statement.index.values == 'Net Interest Income'].squeeze().tolist()[1:],
        'Net Income from Continuing & Discontinued operations': income_statement[income_statement.index.values == 'Net Income From Continuing And Discontinued Operation'].squeeze().tolist()[1:],
        'Normalized Income': income_statement[income_statement.index.values == 'Normalized Income'].squeeze().tolist()[1:],
        'Reconciled Depreciation': income_statement[income_statement.index.values == 'Reconciled Depreciation'].squeeze().tolist()[1:],
        'Net Income from Continuing Operation net minority interest': income_statement[income_statement.index.values == 'Net Income From Continuing Operation Net Minority Interest'].squeeze().tolist()[1:],
        'Total Unusual Items Excluding goodwill': income_statement[income_statement.index.values == 'Total Unusual Items Excluding Goodwill'].squeeze().tolist()[1:],
        'Total Unusual Items': income_statement[income_statement.index.values == 'Total Unusual Items'].squeeze().tolist()[1:],
        'Tax Rate for Calcs': income_statement[income_statement.index.values == 'Tax Rate For Calcs'].squeeze().tolist()[1:],
        'Tax Effect of Unusual Items': income_statement[income_statement.index.values == 'Tax Effect Of Unusual Items'].squeeze().tolist()[1:]
    }



# df for ratio
df_result = pd.DataFrame()

df_income_statement = pd.DataFrame(income_statement_data)

# ratios to analyse


print("")
"""