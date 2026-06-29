import pandas as pd

file_path = "Best Efficiency Moulding Nov.25 to Jan.26.xlsx"

# Read Efficiency Nov.25 sheet with more rows
df = pd.read_excel(file_path, sheet_name="Efficiency Nov.25", header=None)
print("Full sheet (first 20 rows):")
print(df.head(20).to_string())
