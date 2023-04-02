import pandas as pd

# Read the two Excel sheets into dataframes and convert team names to lowercase
df1 = pd.read_excel('table1.xlsx')
df1['Team Name'] = df1['Team Name'].str.lower()

df2 = pd.read_excel('table2.xlsx')
df2['name'] = df2['name'].str.lower()


df1 = df1.rename(columns={'User ID': 'uid'})

# Merge the two dataframes on the common column, "uid"
merged_df = pd.merge(df1, df2, how='inner', on='uid')

# Group the merged dataframe by name and uid and calculate the sum of statements and reasons
grouped_df = merged_df.groupby(['Name', 'uid']).agg({'total_statements': 'sum', 'total_reasons': 'sum'}).reset_index()

# Calculate the total number of statements and reasons for each person
grouped_df['Total'] = grouped_df['total_statements'] + grouped_df['total_reasons']

# Sort the resulting dataframe based on the "Total" column
sorted_df = grouped_df.sort_values(by='Total', ascending=False).reset_index(drop=True)

# Assign a rank to each person based on their position in the sorted dataframe
sorted_df['Rank'] = sorted_df.index + 1

# Select the columns for the output dataframe
output_df = sorted_df[['Rank', 'Name', 'uid', 'total_statements', 'total_reasons']]

# Rename the columns of the output dataframe
output_df = output_df.rename(columns={'total_statements': 'No. of Statements', 'total_reasons': 'No. of Reasons'})

# Save the resulting dataframe to an Excel file
output_df.to_excel('person_stats.xlsx', index=False)
