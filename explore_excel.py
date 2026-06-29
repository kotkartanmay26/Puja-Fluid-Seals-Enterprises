import pandas as pd

file_path = "Best Efficiency Moulding Nov.25 to Jan.26.xlsx"

# Read all sheet names
xl = pd.ExcelFile(file_path)
print("Sheet names found:")
for i, sheet_name in enumerate(xl.sheet_names):
    print(f"{i+1}. {sheet_name}")

print("\n" + "="*50 + "\n")

# Explore each sheet
for sheet_name in xl.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print(f"Sheet: {sheet_name}")
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"First 5 rows:")
    print(df.head())
    print("\n" + "="*50 + "\n")
