import pandas as pd
import itertools
from openpyxl import Workbook

# Load the IPO data from a CSV file
def load_data(file_path):
    return pd.read_csv(file_path)

# Check for overlapping dates
def has_no_overlap(portfolio):
    dates = [(row['IPO Open Date'], row['IPO Close Date'], row['Listing Date']) for index, row in portfolio.iterrows()]
    for i in range(len(dates)):
        for j in range(i + 1, len(dates)):
            if (dates[i][2] >= dates[j][0] and dates[i][0] <= dates[j][2]) or (dates[j][2] >= dates[i][0] and dates[j][0] <= dates[i][2]):
                return False
    return True

# Generate portfolios
def generate_portfolios(ipo_data):
    portfolios = []
    for r in range(1, len(ipo_data) + 1):
        for combination in itertools.combinations(ipo_data.index, r):
            portfolio = ipo_data.loc[list(combination)]
            if has_no_overlap(portfolio):
                portfolios.append(portfolio)
    return portfolios

# Analyze portfolio
def analyze_portfolio(portfolio):
    analysis = {
        'Number of IPOs': len(portfolio),
        'Total Market Cap': portfolio['Market Cap'].sum() if 'Market Cap' in portfolio.columns else 'N/A'
    }
    return analysis

# Save portfolios to Excel
def save_to_excel(portfolios, file_name):
    workbook = Workbook()
    for i, portfolio in enumerate(portfolios):
        sheet = workbook.create_sheet(title=f'Portfolio {i + 1}')
        for r in dataframe_to_rows(portfolio, index=False, header=True):
            sheet.append(r)
        analysis = analyze_portfolio(portfolio)
        sheet.append(['Analysis'])
        for key, value in analysis.items():
            sheet.append([key, value])
    workbook.save(file_name)

# Main function
def main():
    file_path = 'ipo_data.csv'  # Path to your CSV file
    ipo_data = load_data(file_path)
    
    # Ensure unique IPOs
    ipo_data = ipo_data.drop_duplicates(subset=['IPO Identifier'])  # Replace with actual identifier column
    
    portfolios = generate_portfolios(ipo_data)
    save_to_excel(portfolios, 'ipo_portfolios.xlsx')

if __name__ == "__main__":
    main()