import pandas as pd
from datetime import datetime

# Define the source file path and the destination file path
source_file_path = r"G:\Shared drives\sq1 - marketing\seo\Internal Links\Data\url_directory.xlsx"
destination_directory = r"C:\Users\Hooman Deghani\Python\Data Analysis\Backups\URL_directory"

# Generate the current date and time as a string
current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")

# Create the destination file name by appending the current date and time
destination_file_name = f"url_directory_{current_datetime}.xlsx"
destination_file_path = f"{destination_directory}/{destination_file_name}"

# Read the Excel file into a DataFrame
df = pd.read_excel(source_file_path)

# Write the DataFrame to a new Excel file at the destination
df.to_excel(destination_file_path, index=False)

print("Excel file has been copied to the destination location.")