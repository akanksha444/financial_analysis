import pandas as pd
import os
import matplotlib.pyplot as plt
import glob
from functions.ranking import ranking
from functions.comparision_analysis import comparison_analysis

# Specify the pattern to match filenames (e.g., all .xlsx files)
pattern = 'sample/*.xlsx'
filenames = glob.glob(pattern)
filenames = [filename.strip() for filename in filenames]
print(filenames)


for file_path in filenames:
    income_statement = pd.read_excel(file_path, sheet_name='income_statement')
    balance_sheet = pd.read_excel(file_path, sheet_name='balance_sheet')
    cash_flow = pd.read_excel(file_path, sheet_name='cash_flow')

    print(file_path)
    # Define the income statement data
    income_statement_data = {
        'Total Revenue':  income_statement[income_statement['Breakdown'] == 'Total Revenue'].squeeze().tolist()[1:],
        'Credit Losses Provision':  income_statement[income_statement['Breakdown'] == 'Credit Losses Provision'].squeeze().tolist()[1:],
        'Non Interest Expense': income_statement[income_statement['Breakdown'] == 'Non Interest Expense'].squeeze().tolist()[1:],
        'Special Income Charges': income_statement[income_statement['Breakdown'] == 'Special Income Charges'].squeeze().tolist()[1:],
        'Pretax Income': income_statement[income_statement['Breakdown'] == 'Pretax Income'].squeeze().tolist()[1:],
        'Tax Provision': income_statement[income_statement['Breakdown'] == 'Tax Provision'].squeeze().tolist()[1:],
        'Net Income Common Stockholders': income_statement[income_statement['Breakdown'] == 'Net Income Common Stockholders'].squeeze().tolist()[1:],
        'Diluted NI Available to Com Stockholders': income_statement[income_statement['Breakdown'] == 'Diluted NI Available to Com Stockholders'].squeeze().tolist()[1:],
        'Basic EPS': income_statement[income_statement['Breakdown'] == 'Basic EPS'].squeeze().tolist()[1:],
        'Diluted EPS': income_statement[income_statement['Breakdown'] == 'Diluted EPS'].squeeze().tolist()[1:],
        'Basic Average Shares': income_statement[income_statement['Breakdown'] == 'Basic Average Shares'].squeeze().tolist()[1:],
        'Diluted Average Shares': income_statement[income_statement['Breakdown'] == 'Diluted Average Shares'].squeeze().tolist()[1:],
        'Interest Income after Provision for loan loss': income_statement[income_statement['Breakdown'] == 'Interest Income after Provision for Loan Loss'].squeeze().tolist()[1:],
        'Net Income from Continuing & Discontinued operations': income_statement[income_statement['Breakdown'] == 'Net Income from Continuing & Discontinued Operation'].squeeze().tolist()[1:],
        'Normalized Income': income_statement[income_statement['Breakdown'] == 'Normalized Income'].squeeze().tolist()[1:],
        'Reconciled Depreciation': income_statement[income_statement['Breakdown'] == 'Reconciled Depreciation'].squeeze().tolist()[1:],
        'Net Income from Continuing Operation net minority interest': income_statement[income_statement['Breakdown'] == 'Net Income from Continuing Operation Net Minority Interest'].squeeze().tolist()[1:],
        'Total Unusual Items Excluding goodwill': income_statement[income_statement['Breakdown'] == 'Total Unusual Items Excluding Goodwill'].squeeze().tolist()[1:],
        'Total Unusual Items': income_statement[income_statement['Breakdown'] == 'Total Unusual Items'].squeeze().tolist()[1:],
        'Tax Rate for Calcs': income_statement[income_statement['Breakdown'] == 'Tax Rate for Calcs'].squeeze().tolist()[1:],
        'Tax Effect of Unusual Items': income_statement[income_statement['Breakdown'] == 'Tax Effect of Unusual Items'].squeeze().tolist()[1:]
    }



    # Convert data to DataFrame
    df_income_statement = pd.DataFrame(income_statement_data)
    df_result = pd.DataFrame()

    # Calculate financial ratios
    df_result['Gross Profit Margin'] = (df_income_statement['Total Revenue'] - df_income_statement['Credit Losses Provision']) / df_income_statement['Total Revenue']
    df_result['Operating Profit Margin'] = (df_income_statement['Total Revenue'] - df_income_statement['Non Interest Expense'] - df_income_statement['Special Income Charges']) / df_income_statement['Total Revenue']
    df_result['Net Profit Margin'] = df_income_statement['Net Income Common Stockholders'] / df_income_statement['Total Revenue']
    df_result['Earnings per Share (EPS)'] = df_income_statement['Net Income Common Stockholders'] / df_income_statement['Basic Average Shares']
    df_result['Interest Coverage Ratio'] = df_income_statement['Pretax Income'] / df_income_statement['Tax Provision']
    df_result['Effective Tax Rate'] = df_income_statement['Tax Provision'] / df_income_statement['Pretax Income']
    df_result['Leverage Ratio'] = df_income_statement['Total Unusual Items Excluding goodwill'] / df_income_statement['Total Unusual Items']
    df_result['Earnings Before Interest and Taxes (EBIT)'] = df_income_statement['Pretax Income'] + df_income_statement[
        'Interest Income after Provision for loan loss']
    df_result['Operating Income'] = df_income_statement['Pretax Income'] + df_income_statement['Interest Income after Provision for loan loss'] - df_income_statement[
        'Tax Provision']
    df_result['Net Operating Loss (NOL)'] = df_income_statement['Pretax Income'] - df_income_statement['Tax Provision']
    df_result['Adjusted Earnings'] = df_income_statement['Net Income Common Stockholders'] + df_income_statement['Reconciled Depreciation'] + df_income_statement['Special Income Charges']
    df_result['Normalized Earnings'] = df_result['Adjusted Earnings'] + df_income_statement['Total Unusual Items Excluding goodwill']
    df_result['Net Income from Continuing Operations'] = df_income_statement['Net Income from Continuing & Discontinued operations'] - df_income_statement[
        'Total Unusual Items']
    df_result['Net Income Including Non-Controlling Interests'] = df_income_statement['Net Income Common Stockholders']  # Assuming no non-controlling interests
    df_result['Net Income Continuous Operations'] = df_income_statement['Net Income from Continuing Operation net minority interest'] - df_income_statement[
        'Net Income from Continuing & Discontinued operations']
    df_result['Total Comprehensive Income'] = df_income_statement['Net Income Common Stockholders']  # Assuming no other comprehensive income

    # Print the DataFrame with calculated metrics
    print("Income Statement Metrics:")
    print(df_result)

    # Define the balance sheet data
    balance_sheet_data = {
        'Total Assets': balance_sheet[balance_sheet['Breakdown'] == 'Total Assets'].squeeze().tolist()[1:],
        'Total Liabilities Net Minority Interest': balance_sheet[balance_sheet['Breakdown'] == 'Total Liabilities Net Minority Interest'].squeeze().tolist()[1:],
        'Total Equity Gross Minority Interest': balance_sheet[balance_sheet['Breakdown'] == 'Total Equity Gross Minority Interest'].squeeze().tolist()[1:],
        'Total Capitalization': balance_sheet[balance_sheet['Breakdown'] == 'Total Capitalization'].squeeze().tolist()[1:],
        'Net Tangible Assets': balance_sheet[balance_sheet['Breakdown'] == 'Net Tangible Assets'].squeeze().tolist()[1:],
        'Invested Capital': balance_sheet[balance_sheet['Breakdown'] == 'Invested Capital'].squeeze().tolist()[1:],
        'Tangible Book Value': balance_sheet[balance_sheet['Breakdown'] == 'Tangible Book Value'].squeeze().tolist()[1:],
        'Total Debt': balance_sheet[balance_sheet['Breakdown'] == 'Total Debt'].squeeze().tolist()[1:],
        'Share Issued': balance_sheet[balance_sheet['Breakdown'] == 'Share Issued'].squeeze().tolist()[1:],
        'Ordinary Shares Number': balance_sheet[balance_sheet['Breakdown'] == 'Ordinary Shares Number'].squeeze().tolist()[1:],
        'Treasury Shares Number': balance_sheet[balance_sheet['Breakdown'] == 'Treasury Shares Number'].squeeze().tolist()[1:]
    }

    # Convert data to DataFrame
    df_balance_sheet = pd.DataFrame(balance_sheet_data)

    # Calculate financial metrics
    df_result['Current Ratio'] = df_balance_sheet['Total Assets'] / df_balance_sheet['Total Liabilities Net Minority Interest']
    df_result['Debt-to-Equity Ratio'] = df_balance_sheet['Total Debt'] / df_balance_sheet['Total Equity Gross Minority Interest']
    df_result['Leverage Ratio'] = df_balance_sheet['Total Debt'] / df_balance_sheet['Total Assets']
    df_result['Book Value per Share'] = df_balance_sheet['Total Equity Gross Minority Interest'] / df_balance_sheet['Ordinary Shares Number']
    df_result['Tangible Book Value per Share'] = df_balance_sheet['Tangible Book Value'] / df_balance_sheet['Ordinary Shares Number']
    df_result['Market Value per Share'] = df_balance_sheet['Total Capitalization'] / df_balance_sheet['Ordinary Shares Number']
    df_result['Market-to-Book Ratio'] = df_result['Market Value per Share'] / df_result['Book Value per Share']
    df_result['Asset Turnover'] = df_income_statement['Total Revenue']/ df_balance_sheet['Total Assets']
    df_result['Return on Assets (ROA)'] = df_income_statement['Net Income Common Stockholders'] / df_balance_sheet['Total Assets']
    df_result['Return on Equity (ROE)'] = df_income_statement['Net Income Common Stockholders'] / df_balance_sheet['Total Equity Gross Minority Interest']
    df_result['Price to Earning Ratio'] = df_result['Market Value per Share']/ df_result['Earnings per Share (EPS)']
    df_result['Price to Book Ratio'] = df_result['Market Value per Share']/ df_result['Book Value per Share']
    # Print the DataFrame with calculated metrics
    print("Balance Sheet Metrics:")
    print(df_result)

    # Define the cash flow data
    cash_flow_data = {
        'Operating Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Operating Cash Flow'].squeeze().tolist()[1:],
        'Investing Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Investing Cash Flow'].squeeze().tolist()[1:],
        'Financing Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Financing Cash Flow'].squeeze().tolist()[1:],
        'End Cash Position': cash_flow[cash_flow['Breakdown'] == 'End Cash Position'].squeeze().tolist()[1:],
        'Income Tax Paid Supplemental Data': cash_flow[cash_flow['Breakdown'] == 'Income Tax Paid Supplemental Data'].squeeze().tolist()[1:],
        'Interest Paid Supplemental Data': cash_flow[cash_flow['Breakdown'] == 'Interest Paid Supplemental Data'].squeeze().tolist()[1:],
        'Issuance of Debt': cash_flow[cash_flow['Breakdown'] == 'Issuance of Debt'].squeeze().tolist()[1:],
        'Repayment of Debt': cash_flow[cash_flow['Breakdown'] == 'Repayment of Debt'].squeeze().tolist()[1:],
        'Repurchase of Capital Stock': cash_flow[cash_flow['Breakdown'] == 'Repurchase of Capital Stock'].squeeze().tolist()[1:]
    }

    # Convert data to DataFrame
    df_cash_flow = pd.DataFrame(cash_flow_data)

    # Calculate financial metrics
    df_result['Free Cash Flow'] = df_cash_flow['Operating Cash Flow'] + df_cash_flow['Investing Cash Flow']
    df_result['Net Debt Issued'] = df_cash_flow['Issuance of Debt'] + df_cash_flow['Repayment of Debt']
    df_result['Net Capital Stock Repurchased'] = df_cash_flow['Repurchase of Capital Stock']

    # Print the DataFrame with calculated metrics
    print("Cash Flow Metrics:")
    print(df_result)

    df_result.index = income_statement.columns[1:].tolist()
    file_name_without_extension = os.path.splitext(os.path.basename(file_path))[0]

    output_folder = f'analysis_report/{file_name_without_extension}/'

    if not os.path.exists(output_folder):
        # Create the folder if it does not exist
        os.makedirs(output_folder)

    company_report_file = f'{output_folder}{file_name_without_extension}.xlsx'

    if not os.path.isfile(company_report_file):
        # If the file doesn't exist, create it by saving the DataFrame to an Excel file
        df_result.to_excel(company_report_file)

    # Set index=False if you don't want to include row numbers
    df_result.to_excel(company_report_file)

    plotting_data = ['Gross Profit Margin', 'Net Profit Margin', 'Return on Assets (ROA)', 'Return on Equity (ROE)',
                     'Asset Turnover', 'Current Ratio', 'Price to Earning Ratio', 'Price to Book Ratio']

    plt.figure(figsize=(12, 9))
    years = income_statement.columns[1:].tolist()

    for i, term in enumerate(plotting_data):
        x = 2
        y = int(len(plotting_data)/2)
        plt.subplot(x, y, i + 1)  # (rows, columns, plot_number)
        if term in plotting_data:
            #plt.plot([1,2,3,4], [1,2,3,4], color='blue')
            plt.plot(years, df_result[term], color='blue')
            plt.xticks(years)
    plt.tight_layout()  # Adjust layout to prevent overlapping
    plt.savefig(f'{output_folder}time_series_trend.png')


if not os.path.exists('comparison_analysis/'):
    # Create the folder if it does not exist
    os.makedirs('comparison_analysis/')
comparison_analysis()
ranking()
