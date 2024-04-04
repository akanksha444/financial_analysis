import pandas as pd
import os
import matplotlib.pyplot as plt
import glob
import numpy as np

# Specify the pattern to match filenames (e.g., all .xlsx files)
pattern = 'sample/*.xlsx'
filenames = glob.glob(pattern)
print(filenames)

filenames = ['sample/bnp_paribas.xlsx']

for file_path in filenames:
    income_statement = pd.read_excel(file_path, sheet_name='income_statement')
    balance_sheet = pd.read_excel(file_path, sheet_name='balance_sheet')
    cash_flow = pd.read_excel(file_path, sheet_name='cash_flow')

    data = {
        'Year': income_statement.columns.tolist()[1:],
        'Total Revenue': income_statement[income_statement['Breakdown'] == 'Total Revenue'].squeeze().tolist()[1:],
        'Net Interest Income': income_statement[income_statement['Breakdown'] == 'Net Interest Income'].squeeze().tolist()[1:],
        'Non-Interest Income': income_statement[income_statement['Breakdown'] == 'Non Interest Income'].squeeze().tolist()[1:],
        'Fees and Commissions': income_statement[income_statement['Breakdown'] == 'Fees And Commissions'].squeeze().tolist()[1:],
        'Pretax Income': income_statement[income_statement['Breakdown'] == 'Pretax Income'].squeeze().tolist()[1:],
        'Tax Provision': income_statement[income_statement['Breakdown'] == 'Tax Provision'].squeeze().tolist()[1:],
        'Total Assets': balance_sheet[balance_sheet['Breakdown'] == 'Total Assets'].squeeze().tolist()[1:],
        'Total Liabilities': balance_sheet[balance_sheet['Breakdown'] == 'Total Liabilities Net Minority Interest'].squeeze().tolist()[1:],
        'Total Equity': balance_sheet[balance_sheet['Breakdown'] == 'Total Equity Gross Minority Interest'].squeeze().tolist()[1:],
        'Total Loans': balance_sheet[balance_sheet['Breakdown'] == 'Net Loan'].squeeze().tolist()[1:],
        'Total Deposits': balance_sheet[balance_sheet['Breakdown'] == 'Total Deposits'].squeeze().tolist()[1:],
        'Operating Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Operating Cash Flow'].squeeze().tolist()[1:],
        'Investing Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Investing Cash Flow'].squeeze().tolist()[1:],
        'Financing Cash Flow': cash_flow[cash_flow['Breakdown'] == 'Financing Cash Flow'].squeeze().tolist()[1:]
    }

    df = pd.DataFrame(data)

    # Compute ratios
    df['Current Ratio'] = df['Total Assets'] / df['Total Liabilities']
    df['Quick Ratio'] = (df['Total Assets'] - df['Total Loans']) / df['Total Liabilities']
    df['Debt-to-Equity Ratio'] = df['Total Liabilities'] / df['Total Equity']
    df['Debt Ratio'] = df['Total Liabilities'] / df['Total Assets']
    df['Equity Ratio'] = df['Total Equity'] / df['Total Assets']
    df['Net Interest Margin'] = (df['Net Interest Income'] / df['Total Assets']) * 100
    df['ROA'] = df['Pretax Income'] / df['Total Assets']
    df['ROE'] = df['Pretax Income'] / df['Total Equity']
    df['Fee Income Ratio'] = (df['Non-Interest Income'] / df['Total Revenue']) * 100
    df['Asset Turnover Ratio'] = df['Total Revenue'] / df['Total Assets']
    df['Loan-to-Deposit Ratio'] = df['Total Loans'] / df['Total Deposits']
    # Operating Cash Flow Analysis
    df['Operating Cash Flow Margin'] = (df['Operating Cash Flow'] / df['Operating Cash Flow'].sum()) * 100
    # Investing Cash Flow Analysis
    total_investments = df['Investing Cash Flow'].sum()
    capital_expenditures = abs(df['Investing Cash Flow']).sum()
    df['Total Investments'] = total_investments
    df['Capital Expenditures'] = capital_expenditures
    # Financing Cash Flow Analysis
    debt_issuance_repayment = df['Financing Cash Flow'].sum()
    df['Debt Issuance or Repayment'] = debt_issuance_repayment
    # Free Cash Flow Analysis
    df['Free Cash Flow'] = df['Operating Cash Flow'] - capital_expenditures
    # Display computed ratios
    print(df)

    file_name_without_extension = os.path.splitext(os.path.basename(file_path))[0]

    output_folder = f'analysis_report/{file_name_without_extension}/'

    if not os.path.exists(output_folder):
        # Create the folder if it does not exist
        os.makedirs(output_folder)

    company_report_file = f'{output_folder}company_report.xlsx'

    if not os.path.isfile(company_report_file):
        # If the file doesn't exist, create it by saving the DataFrame to an Excel file
        df.to_excel(company_report_file, index=False)

    df.to_excel(company_report_file, index=False)  # Set index=False if you don't want to include row numbers

    firm_ratios = df.columns.tolist()[15:]

    plt.figure(figsize=(12, 6))

    for i, term in enumerate(firm_ratios):
        plt.subplot(4, 4, i+1)  # (rows, columns, plot_number)
        if term in firm_ratios:
            plt.plot(df['Year'], df[term], color='blue')
        plt.title(f'Plot {i+1}: {term}')

    plt.tight_layout()  # Adjust layout to prevent overlapping
    plt.savefig(f'{output_folder}time_series_trend.png')
    plt.show()
    print(f'{output_folder}time_series_trend.png')