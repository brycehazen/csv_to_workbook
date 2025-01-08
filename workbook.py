import pandas as pd
import os

# Get the current working directory
current_dir = os.getcwd()

# Find a file with a .csv or .CSV extension in the current directory
for file in os.listdir(current_dir):
    if file.lower().endswith(".csv"):  # Case-insensitive check for .csv files
        csv_file = file
        break
else:
    raise FileNotFoundError("No CSV file found in the current directory.")

# Read the CSV with specified encoding
try:
    data = pd.read_csv(csv_file, dtype=str, encoding='iso-8859-1')  # Specify ISO-8859-1 encoding
except UnicodeDecodeError as e:
    raise Exception(f"Encoding issue with file {csv_file}: {e}")

# Create the output Excel filename
excel_file = os.path.splitext(csv_file)[0] + ".xlsx"

# Save to Excel
data.to_excel(excel_file, index=False)

print(f"CSV converted to Excel and saved as {excel_file}")
