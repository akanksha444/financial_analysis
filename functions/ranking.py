import glob
import os
import pandas as pd
import matplotlib.pyplot as plt


def ranking():
    report_folder = "analysis_report/"
    pattern = report_folder + '/**/*.xlsx'
    filenames = glob.glob(pattern, recursive=True)
    print(filenames)

    plotting_data = ['Gross Profit Margin', 'Net Profit Margin', 'Return on Assets (ROA)', 'Return on Equity (ROE)',
                     'Asset Turnover', 'Current Ratio', 'Price to Earning Ratio', 'Price to Book Ratio']

    years = []
    companies = [os.path.splitext(os.path.basename(filename))[0] for filename in filenames]


    for term in plotting_data:
        # Plotting
        plt.figure(figsize=(10, 6))
        for i, filename in enumerate(filenames):
            df_plot = pd.read_excel(filename, index_col=0)
            company_name = os.path.splitext(os.path.basename(filename))[0]
            years = df_plot.index.tolist()
            plt.plot(df_plot.index.tolist(), df_plot[term].tolist(), label=f'{company_name}')
            plt.xticks(df_plot.index.tolist())
        plt.xlabel('Year')
        plt.ylabel(f'{term}')
        plt.title(f'Comparison of {term} for Four Companies')
        plt.legend()
        plt.grid(True)
        fig_filename = f'comparison_analysis/{term}.png'
        plt.savefig(fig_filename)


    ratios = plotting_data

    # Create a DataFrame to store the data
    data = pd.DataFrame(index=years, columns=pd.MultiIndex.from_product([companies, ratios]))

    # Generate random ratio values for each company over 4 years
    for filename in filenames:
        for ratio in ratios:
            company = os.path.splitext(os.path.basename(filename))[0]
            df_ranking = pd.read_excel(filename, index_col=0)
            data[(company, ratio)] = df_ranking[ratio].tolist()

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

