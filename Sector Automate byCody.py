import numpy as np
import pandas as pd
import os
import re
from openpyxl import load_workbook, Workbook
from typing import Dict, List
from sklearn.impute import KNNImputer

# Path to the Trendlyne folder
TRENDLYNE_FOLDER_PATH = "data/Trendlyne"
TICKERTAPE_FOLDER_PATH = "data/Tickertape"

# Dictionary to map subsectors to sectors
SUBSECTOR_TO_SECTOR = {
    'Consumer Discretionary': ['Auto Parts', 'Tires & Rubber', 'Four Wheelers', 'Three Wheelers', 'Two Wheelers', 'Cycles', 'Education Services', 'Wellness Services', 'Hotels, Resorts & Cruise Lines', 'Restaurants & Cafes', 'Theme Parks & Gaming', 'Tour & Travel Services', 'Home Electronics & Appliances', 'Home Furnishing', 'Housewares', 'Retail - Apparel', 'Retail - Department Stores', 'Retail - Online', 'Retail - Speciality', 'Apparel & Accessories', 'Footwear', 'Precious Metals, Jewellery & Watches', 'Textiles', 'Animation'],
    'Communication Services': ['Advertising', 'Cable & D2H', 'Movies & TV Serials', 'Publishing', 'Radio', 'Theatres', 'TV Channels & Broadcasters', 'Online Services', 'Telecom Equipments', 'Telecom Infrastructure', 'Telecom Services'],
    'Consumer Staples': ['Alcoholic Beverages', 'Soft Drinks', 'Tea & Coffee', 'Agro Products', 'FMCG - Foods', 'Packaged Foods & Meats', 'Seeds', 'Sugar', 'FMCG - Household Products', 'FMCG - Personal Products', 'FMCG - Tobacco'],
    'Energy': ['Oil & Gas - Equipment & Services', 'Oil & Gas - Exploration & Production', 'Oil & Gas - Refining & Marketing', 'Oil & Gas - Storage & Transportation'],
    'Financials': ['Private Banks', 'Public Banks', 'Asset Management', 'Investment Banking & Brokerage', 'Stock Exchanges & Ratings', 'Consumer Finance', 'Diversified Financials', 'Specialized Finance', 'Insurance', 'Home Financing', 'Payment Infrastructure'],
    'Health Care': ['Biotechnology', 'Health Care Equipment & Supplies', 'Hospitals & Diagnostic Centres', 'Labs & Life Sciences Services', 'Pharmaceuticals'],
    'Industrials': ['Airlines', 'Building Products - Ceramics', 'Building Products - Glass', 'Building Products - Granite', 'Building Products - Laminates', 'Building Products - Pipes', 'Business Support Services', 'Stationery', 'Construction & Engineering', 'Batteries', 'Cables', 'Electrical Components & Equipments', 'Heavy Electrical Equipments', 'Conglomerates', 'Agricultural & Farm Machinery', 'Heavy Machinery', 'Industrial Machinery', 'Rail', 'Shipbuilding', 'Tractors', 'Trucks & Buses', 'Employment Services', 'Airports', 'Dredging', 'Logistics', 'Ports', 'Roads', 'Renewable Energy Equipment & Services', 'Building Products - Prefab Structures', 'Commodities Trading', 'Aerospace & Defense Equipments'],
    'Information Technology': ['Communication & Networking', 'Electronic Equipments', 'Technology Hardware', 'IT Services & Consulting', 'Outsourced services', 'Software Services'],
    'Materials': ['Commodity Chemicals', 'Diversified Chemicals', 'Fertilizers & Agro Chemicals', 'Paints', 'Plastic Products', 'Specialty Chemicals', 'Cement', 'Packaging', 'Iron & Steel', 'Metals - Aluminium', 'Metals - Coke', 'Metals - Copper', 'Metals - Diversified', 'Metals - Lead', 'Mining - Coal', 'Mining - Copper', 'Mining - Diversified', 'Mining - Iron Ore', 'Mining - Manganese', 'Paper Products', 'Wood Products'],
    'Real Estate': ['Real Estate'],
    'Utilities': ['Power Infrastructure', 'Power Transmission & Distribution', 'Gas Distribution', 'Power Trading & Consultancy', 'Power Generation', 'Renewable Energy', 'Water Management'],
    'ETF': ['Gold', 'Equity', 'Debt']
}

# Reverse the dictionary to map each subsector to its sector
SUBSECTOR_TO_SECTOR_REVERSE = {subsector: sector for sector, subsectors in SUBSECTOR_TO_SECTOR.items() for subsector in subsectors}

def clean_column_name(column_name: str) -> str:
    """
    Clean a column name by removing non-alphanumeric characters, spaces, and '/' characters.

    Args:
        column_name (str): The column name to be cleaned.

    Returns:
        str: The cleaned column name.
    """
    # Keep only letters, numbers, spaces, and '/'
    cleaned_name = re.sub(r'[^a-zA-Z0-9 / -]', '', column_name)
    return cleaned_name

def split_all_stocks_csv_into_sector_files(all_stocks_file_path: str) -> None:
    """
    Split the 'all_stocks.csv' file into sector-wise CSV files.

    Args:
        all_stocks_file_path (str): The path to the 'all_stocks.csv' file.
    """
    try:
        all_stocks_df = pd.read_csv(all_stocks_file_path)
    except FileNotFoundError:
        print(f"Error: File '{all_stocks_file_path}' not found.")
        return

    # Map the subsector to sector
    all_stocks_df['Sector'] = all_stocks_df['Sub-Sector'].map(SUBSECTOR_TO_SECTOR_REVERSE)

    # Split the DataFrame into multiple DataFrames based on the 'Sector' column
    sector_dfs = {sector: group.drop(columns=['Sector']) for sector, group in all_stocks_df.groupby('Sector')}

    # Save the DataFrames to CSV
    for sector, sector_df in sector_dfs.items():
        if not sector_df.empty:
            file_path = os.path.join(TICKERTAPE_FOLDER_PATH, f'{sector}.csv')
            sector_df.to_csv(file_path, index=False)

def load_trendlyne_data() -> pd.DataFrame:
    """
    Load data from Trendlyne Excel files and create a combined DataFrame.

    Returns:
        pd.DataFrame: A DataFrame containing the combined data from Trendlyne Excel files.
    """
    trendlyne_dfs = []
    for file in os.listdir(TRENDLYNE_FOLDER_PATH):
        if file.endswith(".xlsx"):
            try:
                file_path = os.path.join(TRENDLYNE_FOLDER_PATH, file)
                df = pd.read_excel(file_path)
                df.drop(columns=['Stock Name', 'BSE code', 'ISIN', 'Current Price', 'Industry Name'], inplace=True)
                df.rename(columns={'NSE code': 'Ticker'}, inplace=True)
                trendlyne_dfs.append(df)
            except Exception as e:
                print(f"Error reading file '{file}': {e}")

    if not trendlyne_dfs:
        print("No Trendlyne data found.")
        return pd.DataFrame()

    trendlyne_data = pd.concat(trendlyne_dfs, ignore_index=True)
    return trendlyne_data

def load_tickertape_data() -> pd.DataFrame:
    """
    Load data from Tickertape CSV files and create a combined DataFrame.

    Returns:
        pd.DataFrame: A DataFrame containing the combined data from Tickertape CSV files.
    """
    tickertape_dfs = []
    for file_name in os.listdir(TICKERTAPE_FOLDER_PATH):
        if file_name.endswith('.csv'):
            try:
                file_path = os.path.join(TICKERTAPE_FOLDER_PATH, file_name)
                df = pd.read_csv(file_path)
                df.insert(2, 'Sector', file_name.split('.')[0])
                tickertape_dfs.append(df)
            except Exception as e:
                print(f"Error reading file '{file_name}': {e}")

    if not tickertape_dfs:
        print("No Tickertape data found.")
        return pd.DataFrame()

    tickertape_data = pd.concat(tickertape_dfs, ignore_index=True)
    return tickertape_data

def merge_data(trendlyne_data: pd.DataFrame, tickertape_data: pd.DataFrame) -> pd.DataFrame:
    """
    Merge Trendlyne and Tickertape data into a single DataFrame.

    Args:
        trendlyne_data (pd.DataFrame): The DataFrame containing Trendlyne data.
        tickertape_data (pd.DataFrame): The DataFrame containing Tickertape data.

    Returns:
        pd.DataFrame: A DataFrame containing the merged data from Trendlyne and Tickertape.
    """
    try:
        merged_data = pd.merge(tickertape_data, trendlyne_data, on='Ticker', how='inner')
    except Exception as e:
        print(f"Error merging data: {e}")
        return pd.DataFrame()

    return merged_data

def save_sector_data_to_excel(merged_data: pd.DataFrame, output_file: str) -> None:
    """
    Save sector-wise data to an Excel file.

    Args:
        merged_data (pd.DataFrame): The DataFrame containing the merged data.
        output_file (str): The path to the output Excel file.
    """
    try:
        with pd.ExcelWriter(output_file) as writer:
            # Iterate over unique sectors
            for sector in merged_data['Sector'].unique():
                sector_df = merged_data[merged_data['Sector'] == sector]
                sector_df.to_excel(writer, sheet_name=sector, index=False)
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

def categorize_market_cap(row: pd.Series) -> str:
    """
    Categorize the market cap of a company as 'Large', 'Medium', or 'Small'.

    Args:
        row (pd.Series): A row of data containing the 'Market Cap' column.

    Returns:
        str: The market cap category ('Large', 'Medium', or 'Small').
    """
    market_cap = row['Market Cap']
    if market_cap > 82000:
        return 'Large'
    elif 26000 < market_cap <= 82000:
        return 'Medium'
    else:
        return 'Small'

def preprocess_data(sector_data: pd.DataFrame) -> pd.DataFrame:
    """
    Preprocess the sector data by cleaning column names, filling missing values, and categorizing market cap.

    Args:
        sector_data (pd.DataFrame): The DataFrame containing the sector data.

    Returns:
        pd.DataFrame: The preprocessed DataFrame.
    """
    # Clean column names
    sector_data.columns = [clean_column_name(col) for col in sector_data.columns]

    # Fill missing values with 0 for specific columns
    columns_to_fill_zero = [
        'Insider Trades - 3M Cumulative',
        'Bulk Deals - 3M Cumulative',
        'FII Holding Change3M',
        'DII Holding Change3M',
        'MF Holding Change3M',
        'Promoter Holding Change3M',
    ]
    sector_data[columns_to_fill_zero].fillna(0, inplace=True)

    # Fill remaining missing values with 0 if more than 25% of the column is missing
    for col in sector_data.columns:
        if sector_data[col].isna().sum() > 0.25 * len(sector_data):
            sector_data[col].fillna(0, inplace=True)

    # Categorize market cap
    sector_data['Market Cap'] = sector_data.apply(categorize_market_cap, axis=1)

    return sector_data

def apply_knn_imputation(segment_df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply KNN imputation to a DataFrame segment.

    Args:
        segment_df (pd.DataFrame): The DataFrame segment to apply KNN imputation.

    Returns:
        pd.DataFrame: The DataFrame segment with missing values imputed using KNN.
    """
    # Exclude non-numeric columns
    numeric_cols = segment_df.select_dtypes(include=[np.number]).columns

    try:
        imputer = KNNImputer(n_neighbors=5)
        segment_df[numeric_cols] = imputer.fit_transform(segment_df[numeric_cols])
    except Exception as e:
        print(f"An error occurred during KNN imputation: {e}")
        segment_df[numeric_cols] = segment_df[numeric_cols].fillna(0)

    return segment_df

def impute_missing_values(preprocessed_data: pd.DataFrame) -> pd.DataFrame:
    """
    Impute missing values in the preprocessed data using KNN imputation for each market cap segment.

    Args:
        preprocessed_data (pd.DataFrame): The preprocessed DataFrame.

    Returns:
        pd.DataFrame: The DataFrame with missing values imputed using KNN imputation.
    """
    small_df = preprocessed_data[preprocessed_data['Market Cap'] == 'Small']
    medium_df = preprocessed_data[preprocessed_data['Market Cap'] == 'Medium']
    large_df = preprocessed_data[preprocessed_data['Market Cap'] == 'Large']

    # Apply KNN Imputation to each segment
    small_df_imputed = apply_knn_imputation(small_df)
    medium_df_imputed = apply_knn_imputation(medium_df)
    large_df_imputed = apply_knn_imputation(large_df)

    imputed_data = pd.concat([small_df_imputed, medium_df_imputed, large_df_imputed])
    return imputed_data

def create_inverse(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """
    Create an inverse column for a specified column in the DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to create the inverse column.
        column (str): The name of the column to create the inverse for.

    Returns:
        pd.DataFrame: The DataFrame with the inverse column added.
    """
    # Calculate inverse of specified column, handling zeros
    inverse_column = 1 / df[column].replace(0, np.nan)
    max_value = inverse_column.max()
    inverse_column = inverse_column.apply(lambda x: abs(x) + max_value if x < 0 else x)
    inverse_column.fillna(max_value + 1, inplace=True)

    df[column] = inverse_column
    return df

def feature_engineering(imputed_data: pd.DataFrame) -> pd.DataFrame:
    """
    Perform feature engineering on the imputed data.

    Args:
        imputed_data (pd.DataFrame): The DataFrame with imputed missing values.

    Returns:
        pd.DataFrame: The DataFrame with engineered features.
    """
    imputed_data['Altman Zscore'] = imputed_data['Altman Zscore'] - 3
    imputed_data['Fundamentals'] = imputed_data['Piotroski Score'] * imputed_data['Altman Zscore']
    imputed_data.drop(columns=['Piotroski Score', 'Altman Zscore'], inplace=True)

    investor_cols = ['FII Holding Change3M', 'DII Holding Change3M', 'MF Holding Change3M', 'Promoter Holding Change3M', 'Insider Trades - 3M Cumulative', 'Bulk Deals - 3M Cumulative']
    imputed_data['Investor Sentiment'] = imputed_data[investor_cols].mean(axis=1)
    insider_cols = ['Insider Trades - 3M Cumulative', 'Promoter Holding Change3M']
    imputed_data['Promoter Sentiment'] = imputed_data[insider_cols].mean(axis=1)
    imputed_data.drop(columns=investor_cols+insider_cols, inplace=True)

    price_cols = ['Price / Sales', 'Price / CFO']
    imputed_data['Pricing'] = imputed_data[price_cols].mean(axis=1)*imputed_data['PE Ratio']*imputed_data['PB Ratio']
    imputed_data.drop(columns=price_cols+['PE Ratio', 'PB Ratio'], inplace=True)

    valuation_cols = ['EV / Invested Capital','EV / Revenue Ratio', 'EV / Free Cash Flow', 'EV/EBITDA Ratio']
    imputed_data['Valuation'] = imputed_data[valuation_cols].mean(axis=1)
    imputed_data.drop(columns=valuation_cols, inplace=True)

    profitability_cols = ['Return on Investment', 'ROCE', 'Return on Equity', 'Return on Assets', 'Dividend Yield']
    imputed_data['Profitability'] = imputed_data[profitability_cols].mean(axis=1)*imputed_data['Net Profit Margin']
    imputed_data.drop(columns=profitability_cols+['Net Profit Margin'], inplace=True)

    imputed_data = create_inverse(imputed_data, 'Cash Conversion Cycle')
    turnover_cols = ['Interest Coverage Ratio', 'Asset Turnover Ratio', 'Inventory Turnover Ratio', 'Working Capital Turnover Ratio']
    imputed_data['Business Turnover'] = imputed_data[turnover_cols].mean(axis=1)*imputed_data['Current Ratio']*imputed_data['Cash Conversion Cycle']
    imputed_data.drop(columns=turnover_cols +['Current Ratio', 'Cash Conversion Cycle'], inplace=True)

    future_eps = imputed_data['1Y Forward EPS Growth'] - imputed_data['1Y Historical EPS Growth']
    future_rev = imputed_data['1Y Forward Revenue Growth'] - imputed_data['1Y Historical Revenue Growth']
    imputed_data['Future Growth'] = future_eps + future_rev
    imputed_data.drop(columns=['1Y Forward EPS Growth', '1Y Historical EPS Growth', '1Y Forward Revenue Growth', '1Y Historical Revenue Growth'], inplace=True)

    return imputed_data

def calculate_weighted_rank(engineered_data: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate the weighted rank for each row in the engineered data.

    Args:
        engineered_data (pd.DataFrame): The DataFrame with engineered features.

    Returns:
        pd.DataFrame: The DataFrame with the weighted rank column added.
    """
    # Define weights for each rank
    weights = {
        'Rank1': 0.2,
        'Rank2': 0.1,
        'Rank3': 0.15,
        'Rank4': 0.1,
        'Rank5': 0.1,
        'Rank6': 0.05,
        'Rank7': 0.1,
        'Rank8': 0.2
    }

    # Group by 'Market Cap' categories
    grouped = engineered_data.groupby('Market Cap')

    # Initialize an empty list to store the results
    result_dfs = []

    # Iterate over each group
    for name, group in grouped:
        # Calculate ranks within each group
        group['Rank1'] = group['Fundamentals'].rank(method='min', ascending=False)
        group['Rank2'] = group['Investor Sentiment'].rank(method='min', ascending=False)
        group['Rank3'] = group['Promoter Sentiment'].rank(method='min', ascending=False)
        group['Rank4'] = group['Pricing'].rank(method='min', ascending=True)
        group['Rank5'] = group['Valuation'].rank(method='min', ascending=True)
        group['Rank6'] = group['Profitability'].rank(method='min', ascending=False)
        group['Rank7'] = group['Business Turnover'].rank(method='min', ascending=False)
        group['Rank8'] = group['Future Growth'].rank(method='min', ascending=False)

        # Calculate weighted rank
        group['Weighted_rank'] = group[list(weights.keys())].mul(list(weights.values())).sum(axis=1)

        # Calculate final rank
        group['Rank'] = group['Weighted_rank'].rank(method='min', ascending=True)

        # Reorder columns
        columns_order = ['Rank', 'Name', 'Ticker', 'Sector', 'Sub-Sector', 'Market Cap', 'Fundamentals', 'Rank1', 'Investor Sentiment', 'Rank2', 'Promoter Sentiment', 'Rank3', 'Pricing', 'Rank4', 'Valuation', 'Rank5', 'Profitability', 'Rank6', 'Business Turnover', 'Rank7', 'Future Growth', 'Rank8', 'Weighted_rank']
        group = group[columns_order]

        # Append the result to the list
        result_dfs.append(group)

    # Concatenate the results
    ranked_data = pd.concat(result_dfs)
    ranked_data.sort_values(by='Rank', ascending=True, inplace=True)

    return ranked_data

def save_ranked_data_to_excel(ranked_data: pd.DataFrame, output_file: str) -> None:
    """
    Save the ranked data to an Excel file with separate sheets for each market cap category.

    Args:
        ranked_data (pd.DataFrame): The DataFrame containing the ranked data.
        output_file (str): The path to the output Excel file.
    """
    try:
        with pd.ExcelWriter(output_file) as writer:
            # Create a new workbook
            workbook = writer.book

            # Group the data by Market Cap and save them in different dataframes
            grouped = ranked_data.groupby('Market Cap')

            # Loop through each group
            for name, group in grouped:
                # Create a new sheet with the market cap category as the sheet name
                worksheet = workbook.add_worksheet(name)

                # Write the headers starting from the first row
                for col_num, column_title in enumerate(group.columns):
                    worksheet.write(0, col_num, column_title)

                # Write the data starting from the second row
                for row_num, row in enumerate(group.values):
                    for col_num, cell_value in enumerate(row):
                        worksheet.write(row_num + 1, col_num, cell_value)

            # Save the workbook
            writer.save()
    except Exception as e:
        print(f"Error saving ranked data to Excel: {e}")

# Main function
def main():
    # Split 'all_stocks.csv' into sector-wise CSV files
    all_stocks_file_path = os.path.join(TICKERTAPE_FOLDER_PATH, 'all_stocks.csv')
    split_all_stocks_csv_into_sector_files(all_stocks_file_path)

    # Load Trendlyne data
    trendlyne_data = load_trendlyne_data()

    # Load Tickertape data
    tickertape_data = load_tickertape_data()

    # Merge Trendlyne and Tickertape data
    merged_data = merge_data(trendlyne_data, tickertape_data)

    # Save sector-wise data to an Excel file
    save_sector_data_to_excel(merged_data, 'Sector_Universe_Data.xlsx')

    # Preprocess the data
    preprocessed_data = preprocess_data(merged_data)

    # Impute missing values
    imputed_data = impute_missing_values(preprocessed_data)

    # Perform feature engineering
    engineered_data = feature_engineering(imputed_data)

    # Calculate weighted rank
    ranked_data = calculate_weighted_rank(engineered_data)

    # Save ranked data to an Excel file
    save_ranked_data_to_excel(ranked_data, 'Sector Analysis New.xlsx')

if __name__ == "__main__":
    main()

