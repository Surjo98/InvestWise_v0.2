import os
import pandas as pd

# Path to the Trendlyne folder
folder_path_tl = "Trendlyne"
folder_path_tt = "Tickertape"

# List to store dataframes
dfs = []
for file in os.listdir(folder_path_tl):
    if file.endswith(".xlsx"):

        file_path = os.path.join(folder_path_tl, file)
        df = pd.read_excel(file_path)
        df.drop(columns=['Stock Name', 'BSE code', 'ISIN', 'Current Price', 'Industry Name'], inplace=True)
        df.rename(columns={'NSE code': 'Ticker'}, inplace=True)
        dfs.append(df)

df_tl = pd.concat(dfs, ignore_index=True)


dfs = []
for file_name in os.listdir(folder_path_tt):
    if file_name.endswith('.csv'):

        file_path = os.path.join(folder_path_tt, file_name)
        df = pd.read_csv(file_path)
        df.insert(2, 'Sector', file_name.split('.')[0])
        dfs.append(df)

df_tt = pd.concat(dfs, ignore_index=True)

# Concatenate sub-dfs into final df
df = pd.merge(df_tt, df_tl, on='Ticker', how='inner')

# Create a new Excel writer object
with pd.ExcelWriter('Sector_Universe_Data.xlsx') as writer:
    # Iterate over unique sectors
    for sector in df['Sector'].unique():
        sector_df = df[df['Sector'] == sector]
        sector_df.to_excel(writer, sheet_name=sector, index=False)

print("Combined data has been saved.")