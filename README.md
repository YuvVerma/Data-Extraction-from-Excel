# Data-extraction-from-excel
import pandas as pd
import os

# Specify the file paths
source_file = "D:\Project 1\Phase1\Excel Source\Project-Management-Sample-Data.xlsx"  # Path to your source Excel file
destination_file = os.path.join(os.path.expanduser('~'), 'Documents', "D:\Project 1\Phase1OP\output1.xlsx")  # Path to your destination Excel file

# Load the source Excel file
df = pd.read_excel(source_file)

# Specify the columns you want to copy (you can use column names or indices)
columns_to_copy = ['Project Name', 'Task Name','Assigned to','Progress']  # Replace with your column names

# Create a new DataFrame with the specified columns
df_copy = df[columns_to_copy]

# Write the new DataFrame to the destination Excel file
df_copy.to_excel(destination_file, index=False)

print(f"Copied columns {columns_to_copy} from {source_file} to {destination_file}.")
