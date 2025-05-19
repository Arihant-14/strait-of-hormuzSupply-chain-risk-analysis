import pandas as pd

# Step 1: Load files
consumption_path = r"C:\Users\ariha\Downloads\Data Folders\Montly Oil Consumption 2013-23.xlsx"
production_path = r"C:\Users\ariha\Downloads\Data Folders\Monthly Oil Production 2013-23.xlsx"

consumption_df = pd.read_excel(consumption_path, index_col=0)
production_df = pd.read_excel(production_path, index_col=0)

# Step 2: Transpose
consumption_df_T = consumption_df.T
production_df_T = production_df.T

# Step 3: Show original index for debugging
print("\n🟡 Original consumption index (first 5):")
print(consumption_df_T.index[:5])

# Step 4: Convert index to datetime

# Consumption: originally strings like "2013 April"
consumption_df_T.index = pd.to_datetime(consumption_df_T.index.str.strip(), format="%Y %B", errors="coerce")

# Production: already datetime-like, just parse directly
production_df_T.index = pd.to_datetime(production_df_T.index, errors="coerce")

# Drop any rows where dates could not be parsed
consumption_df_T = consumption_df_T[~consumption_df_T.index.isna()]
production_df_T = production_df_T[~production_df_T.index.isna()]

print("\n✅ Parsed index sample:")
print(consumption_df_T.index[:3])

# Step 5: Align countries and dates
common_countries = consumption_df_T.columns.intersection(production_df_T.columns)
common_dates = consumption_df_T.index.intersection(production_df_T.index)

consumption_df_T = consumption_df_T.loc[common_dates, common_countries]
production_df_T = production_df_T.loc[common_dates, common_countries]

# Debug check
print("\n✅ Countries matched:", list(common_countries))
print("✅ Date range:", common_dates.min(), "to", common_dates.max())
print("✅ Shape:", consumption_df_T.shape)

# Step 6: Calculate net exports
net_exports = production_df_T - consumption_df_T

# Step 7: Hormuz dependency ratios
hormuz_ratios = {
    "Iran": 1.0,
    "Iraq": 0.95,
    "Kuwait": 0.90,
    "Qatar": 1.0,
    "Saudi Arabia": 0.65
}

# Only apply to relevant countries
hormuz_countries = [c for c in net_exports.columns if c in hormuz_ratios]
ratios_series = pd.Series(hormuz_ratios)

hormuz_volume = net_exports[hormuz_countries].multiply(ratios_series[hormuz_countries], axis=1)

# Step 8: Export to Excel
output_path = r"C:\Users\ariha\Downloads\Hormuz_Dependent_Exports_Timeseries_FIXED.xlsx"
hormuz_volume.to_excel(output_path)

print(f"\n✅ Exported to: {output_path}")





