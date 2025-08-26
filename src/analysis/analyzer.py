import pandas as pd
from itertools import combinations
import openpyxl

# Load the IPO data from a CSV file
def load_data(file_path):
    return pd.read_csv(file_path)

# Check for overlapping dates
def has_no_overlap(portfolio):
    for i in range(len(portfolio)):
        for j in range(i + 1, len(portfolio)):
            if (portfolio.iloc[i]['listing_date'] <= portfolio.iloc[j]['close_date'] and
                portfolio.iloc[j]['listing_date'] <= portfolio.iloc[i]['close_date']):
                return False
    return True

# Generate portfolios
def generate_portfolios(ipo_data):
    unique_ipos = ipo_data.drop_duplicates(subset=['ipo_id'])  # Assuming 'ipo_id' is the unique identifier
    portfolios = []
    
    for r in range(1, len(unique_ipos) + 1):
        for combo in combinations(unique_ipos.index, r):
            portfolio = unique_ipos.loc[list(combo)]
            if has_no_overlap(portfolio):
                portfolios.append(portfolio)
    
    return portfolios

# Analyze portfolios
def analyze_portfolio(portfolio):
    analysis = {
        'number_of_ipos': len(portfolio),
        'total_market_cap': portfolio['market_cap'].sum()  # Assuming 'market_cap' is a column in the DataFrame
    }
    return analysis

# Write portfolios to Excel
def write_to_excel(portfolios, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, portfolio in enumerate(portfolios):
            portfolio_name = f'Portfolio_{i + 1}'
            portfolio.to_excel(writer, sheet_name=portfolio_name, index=False)
            analysis = analyze_portfolio(portfolio)
            analysis_df = pd.DataFrame([analysis])
            analysis_df.to_excel(writer, sheet_name=f'{portfolio_name}_Analysis', index=False)

# Main function
def main(input_file, output_file):
    ipo_data = load_data(input_file)
    portfolios = generate_portfolios(ipo_data)
    write_to_excel(portfolios, output_file)

# Example usage
if __name__ == "__main__":
    input_file = 'ipos.csv'  # Path to your CSV file
    output_file = 'ipo_portfolios.xlsx'  # Output Excel file
    main(input_file, output_file)