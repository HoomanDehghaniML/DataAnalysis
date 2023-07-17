import pandas as pd

# Read the first CSV file into a dataframe
df1 = pd.read_excel(r"C:\Users\Hooman Deghani\Python\Data Analysis\Outreach - Skyscraper\bike-theft.xlsx")

# Read the second CSV file into a dataframe
df2 = pd.read_excel(r"C:\Users\Hooman Deghani\Python\Data Analysis\Outreach - Skyscraper\bike-theft-2.xlsx")

# Merge the dataframes based on a common column (e.g., 'id')
merged_df = pd.concat([df1, df2]).drop_duplicates(subset=['Referring page URL'])

merged_df = merged_df.dropna(subset=['Recipient'])

# Save the merged dataframe to a new CSV file
merged_df.to_excel('merged_data.xlsx', index=False)