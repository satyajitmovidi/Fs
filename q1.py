import pandas as pd

# Read the two Excel sheets into dataframes and convert team names to lowercase
df1 = pd.read_excel('table1.xlsx')
df1['Team Name'] = df1['Team Name'].str.lower()

df2 = pd.read_excel('table2.xlsx')
df2['name'] = df2['name'].str.lower()

# Rename the "User ID" column in df1 to "uid"
df1 = df1.rename(columns={'User ID': 'uid'})

# Merge the two dataframes on the common column, "uid"
merged_df = pd.merge(df1, df2, how='inner', on='uid')

# Group the merged dataframe by team name and calculate the average statements and reasons
grouped_df = merged_df.groupby('Team Name').agg({'total_statements': 'mean', 'total_reasons': 'mean'})

# Reset the index of the grouped dataframex
grouped_df = grouped_df.reset_index()

# Rename the columns of the grouped dataframe
grouped_df = grouped_df.rename(columns={'total_statements': 'Average Statements per Team', 'total_reasons': 'Average Responses per Team'})

# Format the numbers in the output dataframe to always have 2 digits after the decimal point
grouped_df = grouped_df.applymap(lambda x: '{:.2f}'.format(x) if isinstance(x, (int, float)) else x)

# Sort the resulting dataframe based on the "Average Statements per Team" and "Average Responses per Team" columns
grouped_df = grouped_df.sort_values(by=['Average Statements per Team', 'Average Responses per Team'], ascending=False)

# Save the resulting dataframe to an Excel file
grouped_df.to_excel('team_stats.xlsx', index=False)
