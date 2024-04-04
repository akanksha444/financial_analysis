import glob
import os
import pandas as pd
import matplotlib.pyplot as plt


def ranking():
    pattern = 'analysis_report/**/*.xlsx'
    filenames = glob.glob(pattern)
    print(filenames)

    plotting_data = ['Gross Profit Margin', 'Net Profit Margin', 'Return on Assets (ROA)', 'Return on Equity (ROE)',
                     'Asset Turnover', 'Current Ratio', 'Price to Earning Ratio', 'Price to Book Ratio']

    years = []
    companies = [os.path.splitext(os.path.basename(filename))[0] for filename in filenames]
    # Plotting
    plt.figure(figsize=(10, 6))

    for term in plotting_data:
        for i, filename in enumerate(filenames):
            df = pd.read_excel(filename, index_col=0)
            company_name = os.path.splitext(os.path.basename(filename))[0]
            years = df.index.tolist()
            plt.plot(df.index.tolist(), df[term].tolist(), label=f'{company_name}')
            plt.xticks(df.index.tolist())
        plt.xlabel('Year')
        plt.ylabel(f'{term}')
        plt.title(f'Comparison of {term} for Five Companies')
        plt.legend()
        plt.grid(True)
        plt.savefig(f'comparison_analysis/{term}.png')

    ratios = plotting_data

    # Create a DataFrame to store the data
    data = pd.DataFrame(index=years, columns=pd.MultiIndex.from_product([companies, ratios]))

    # Generate random ratio values for each company over 4 years
    for filename in filenames:
        for ratio in ratios:
            company = os.path.splitext(os.path.basename(filename))[0]
            df = pd.read_excel(filename, index_col=0)
            data[(company, ratio)] = df[ratio].tolist()

    # Create a DataFrame to store the rankings
    rankings = pd.DataFrame(index=years, columns=pd.MultiIndex.from_product([companies, ratios]))

    # Calculate rankings for each ratio and each year
    for year in years:
        for ratio in ratios:
            # Calculate ranking for each company based on the ratio values for the current year
            print(year, ratio)
            rankings.loc[year, (slice(None), ratio)] = data.loc[year, (slice(None), ratio)].rank(ascending=False)

    # Display the ranking table
    print("Ranking Table:")
    rankings.to_excel(f'comparison_analysis/ranking.xlsx')

