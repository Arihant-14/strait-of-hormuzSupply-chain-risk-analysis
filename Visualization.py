#Bar plot for Hormuz Ratio by country:
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import seaborn as sns
import matplotlib.pyplot as plt

# Read the Hormuz Ratio data
hormuz_ratio = pd.read_excel('Winal Workin.xlsx', sheet_name='Hormuz Ratio')

# Create a bar plot for Hormuz Ratio
fig = px.bar(hormuz_ratio,
             x='Country',
             y='Ratio',
             title='Hormuz Ratio by Country',
             color='Ratio',
             color_continuous_scale='Viridis')
fig.update_layout(
    xaxis_title="Country",
    yaxis_title="Ratio",
    template="plotly_white"
)

# Save the plot
fig.write_html("hormuz_ratio_plot.html")


#plot line for Net Export Trends:

# Read the Net Export data
net_export = pd.read_excel('Winal Workin.xlsx', sheet_name='Net Export')

# Melt the dataframe to convert months to a single column
net_export_melted = net_export.melt(id_vars=['Country'], var_name='Month', value_name='Export')

# Create a line plot for Net Export trends
fig = px.line(net_export_melted,
              x='Month',
              y='Export',
              color='Country',
              title='Net Export Trends by Country Over Time')

fig.update_layout(
    xaxis_title="Month",
    yaxis_title="Net Export",
    template="plotly_white",
    showlegend=True
)

# Rotate x-axis labels for better readability
fig.update_xaxes(tickangle=45)

# Save the plot
fig.write_html("net_export_trends.html")

# Create a heatmap of exports over time
pivot_df = net_export_df_melted.pivot(index='Month', columns='Country', values='Export')
plt.figure(figsize=(15, 10))
sns.heatmap(pivot_df, cmap='YlOrRd', cbar_kws={'label': 'Net Exports (Thousand Barrels per Day)'})
plt.title('Heatmap of Net Oil Exports Over Time')
plt.xlabel('Country')
plt.ylabel('Month')
plt.tight_layout()
plt.show()

# Print summary statistics
print("\
Summary Statistics of Net Exports by Country:")
print(net_export_df_melted.groupby('Country')['Export'].describe())