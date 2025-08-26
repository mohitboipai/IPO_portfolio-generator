### Step-by-Step Guide

1. **Read the CSV File**: Load the IPO data from the CSV file into a pandas DataFrame.
2. **Filter Unique IPOs**: Ensure that the IPOs are unique based on their identifiers (e.g., ticker symbol).
3. **Generate Portfolios**: Create combinations of IPOs ensuring that their listing dates do not overlap.
4. **Output to Excel**: Write each portfolio to a separate sheet in an Excel file.
5. **Analysis**: Provide a basic analysis of each portfolio, such as the number of IPOs, total market cap, etc.

### Sample Code

```python
import pandas as pd
from itertools import combinations
import openpyxl

def load_data(file_path):
    # Load the CSV file into a DataFrame
    return pd.read_csv(file_path)

def filter_unique_ipos(df):
    # Assuming 'ticker' is the unique identifier for IPOs
    return df.drop_duplicates(subset='ticker')

def generate_portfolios(df):
    portfolios = []
    n = len(df)
    
    # Generate all combinations of IPOs
    for r in range(1, n + 1):
        for combo in combinations(df.iterrows(), r):
            # Extract the IPOs from the combination
            selected_ipos = [row[1] for row in combo]
            if is_valid_portfolio(selected_ipos):
                portfolios.append(selected_ipos)
    
    return portfolios

def is_valid_portfolio(ipos):
    # Check for overlapping listing dates
    listing_dates = [ipo['listing_date'] for ipo in ipos]
    return len(listing_dates) == len(set(listing_dates))

def analyze_portfolio(portfolio):
    # Perform analysis on the portfolio
    analysis = {
        'num_ipos': len(portfolio),
        'total_market_cap': sum(ipo['market_cap'] for ipo in portfolio)  # Assuming 'market_cap' exists
    }
    return analysis

def save_to_excel(portfolios, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, portfolio in enumerate(portfolios):
            # Create a DataFrame for the portfolio
            portfolio_df = pd.DataFrame(portfolio)
            # Analyze the portfolio
            analysis = analyze_portfolio(portfolio)
            # Add analysis as a new row
            analysis_df = pd.DataFrame([analysis])
            portfolio_df = pd.concat([portfolio_df, analysis_df], ignore_index=True)
            # Write to a new sheet
            portfolio_df.to_excel(writer, sheet_name=f'Portfolio_{i+1}', index=False)

def main(file_path, output_file):
    df = load_data(file_path)
    unique_ipos = filter_unique_ipos(df)
    portfolios = generate_portfolios(unique_ipos)
    save_to_excel(portfolios, output_file)

# Example usage
if __name__ == "__main__":
    input_file = 'ipos.csv'  # Path to your CSV file
    output_file = 'ipo_portfolios.xlsx'  # Output Excel file
    main(input_file, output_file)
```

### Explanation of the Code

1. **Loading Data**: The `load_data` function reads the CSV file into a DataFrame.
2. **Filtering Unique IPOs**: The `filter_unique_ipos` function removes duplicate IPOs based on the ticker.
3. **Generating Portfolios**: The `generate_portfolios` function creates combinations of IPOs and checks for valid portfolios using the `is_valid_portfolio` function.
4. **Analyzing Portfolios**: The `analyze_portfolio` function calculates the number of IPOs and total market cap for each portfolio.
5. **Saving to Excel**: The `save_to_excel` function writes each portfolio and its analysis to separate sheets in an Excel file.

### Note
- Ensure that the CSV file contains the necessary columns such as `ticker`, `listing_date`, and `market_cap`.
- You may need to adjust the column names in the code based on your actual CSV file structure.
- This code assumes that the listing dates are in a format that can be compared directly (e.g., YYYY-MM-DD). If they are in a different format, you may need to convert them to datetime objects.