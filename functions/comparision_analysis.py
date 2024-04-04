import glob
import os
import pandas as pd
import matplotlib.pyplot as plt


def comparison_analysis():
    pattern = 'analysis_report/**/*.xlsx'
    filenames = glob.glob(pattern)
    print(filenames)

    plotting_data = ['Gross Profit Margin', 'Net Profit Margin', 'Return on Assets (ROA)', 'Return on Equity (ROE)',
                     'Asset Turnover', 'Current Ratio', 'Price to Earning Ratio', 'Price to Book Ratio']


    # Plotting





    # Plot the ratio values for each company
    df_cons = pd.DataFrame()
    data = {}
    years = []
    for term in plotting_data:
        plt.figure(figsize=(10, 6))
        for i, filename in enumerate(filenames):
            df = pd.read_excel(filename, index_col=0)
            company_name = os.path.splitext(os.path.basename(filename))[0]
            years = df.index.tolist()
            plt.plot(df.index.tolist(), df[term].tolist(), label=f'{company_name}')
            plt.xticks(df.index.tolist())
            df_cons[f'{term}_{company_name}'] = df[term].tolist()
            for year in years:
                data[(term, company_name)] = df[term].tolist()

        plt.xlabel('Year')
        plt.ylabel(f'{term}')
        plt.title(f'Comparison of {term} for Five Companies')
        plt.legend()
        plt.grid(True)
        plt.savefig(f'comparison_analysis/{term}.png')
    df_cons = df_cons.transpose()
    df_cons.columns = years
    df_cons.to_excel(f'comparison_analysis/comparison_sheet.xlsx')
    df_dict = pd.DataFrame(data)
    df_dict.index = years
    df_dict.to_excel(f'comparison_analysis/comparison_sheet2.xlsx')
    df_dict_transpose = df_dict.transpose()
    df_dict_transpose.to_excel(f'comparison_analysis/comparison_sheet3.xlsx')
