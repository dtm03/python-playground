import pandas as pd

# File paths
file_emimi = r'MSCIEMIMI.csv'
file_smallcap = r'MSCISmallCap.csv'
file_world = r'MSCIWorld.xlsx'
file_europe = r'StoxxEurope600.xlsx'

# Read the CSV files
df_emimi = pd.read_csv(file_emimi, delimiter=',', encoding='utf-8', skiprows=2)
df_smallcap = pd.read_csv(file_smallcap, delimiter=',', encoding='utf-8', skiprows=2)

# Read the Excel file starting from the fourth row, with the correct column names starting from row 5
df_world = pd.read_excel(file_world, header=3)

# Read the Excel file starting from the twentieth row, with the correct column names starting from row 21
df_europe = pd.read_excel(file_europe, header=19)

# Clean the column names to remove any leading/trailing spaces
df_emimi.columns = df_emimi.columns.str.strip()
df_smallcap.columns = df_smallcap.columns.str.strip()
df_world.columns = df_world.columns.str.strip()
df_europe.columns = df_europe.columns.str.strip()

# Extract the ISIN column from the Excel files and drop any NaN values
isin_world = df_world['ISIN'].dropna().astype(str)
isin_europe = df_europe['ISIN'].dropna().astype(str)

# Extract the 'Name' column from each dataframe and drop any NaN values
names_emimi = df_emimi['Name'].dropna().astype(str)
names_smallcap = df_smallcap['Name'].dropna().astype(str)
names_world = df_world['Name'].dropna().astype(str)
names_europe = df_europe['Name'].dropna().astype(str)

# Check intersections between all different combinations of ETFs
# Find the intersection of the sets of World and Europe based on ISINs
common_isin = set(isin_world).intersection(set(isin_europe))

# Find the intersection of the sets of Small Cap and Europe based on CSV names
common_csv_names = set(names_smallcap).intersection(set(names_europe))
print(len(common_csv_names))

# Find the intersection of the sets of EMIMI and Europe based on names
common_names_emimi_europe = set(names_emimi).intersection(set(names_europe))

# Find the intersection of the sets of EMIMI and Small Cap based on names
common_names_emimi_smallcap = set(names_emimi).intersection(set(names_smallcap))

# Find the intersection of the sets of Small Cap and World based on names
common_names_smallcap_world = set(names_smallcap).intersection(set(names_world))

# Find the intersection of the sets of EMIMI and World based on names
common_names_emimi_world = set(names_emimi).intersection(set(names_world))

# Convert all the sets to sorted lists
common_isin_sorted = sorted(list(common_isin))
common_csv_names_sorted = sorted(list(common_csv_names))
common_names_emimi_europe_sorted = sorted(list(common_names_emimi_europe))
common_names_emimi_smallcap_sorted = sorted(list(common_names_emimi_smallcap))
common_names_smallcap_world_sorted = sorted(list(common_names_smallcap_world))
common_names_emimi_world_sorted = sorted(list(common_names_emimi_world))

# Total stocks in each ETF and combined and total unique stocks
total_emimi = len(names_emimi)
total_smallcap = len(names_smallcap)
total_world = len(names_world)
total_europe = len(names_europe)
total = total_emimi + total_smallcap + total_world + total_europe
total_not_europe = total_emimi + total_smallcap + total_world
print(f'Total stocks in EMIMI: {total_emimi}')
print(f'Total stocks in Small Cap: {total_smallcap}')
print(f'Total stocks in Excel: {total_world}')
print(f'Total stocks in Europe: {total_europe}')
print(f'Total stocks in all four: {total}')
print(f'Total not common stocks: {total - len(common_isin) - len(common_csv_names) - len(common_names_emimi_europe) - len(common_names_emimi_smallcap) - 
                                  len(common_names_smallcap_world) - len(common_names_emimi_world)}')
print(f'Total stocks in all but Europe: {total_not_europe}')
print(f'Total unique stocks not europe: {total_not_europe - len(common_csv_names) - len(common_names_emimi_smallcap) - len(common_names_smallcap_world) - len(common_names_emimi_world)}')