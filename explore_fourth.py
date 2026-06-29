import pandas as pd

print("=== Exploring List of Employees M - Copy.xls ===")
file1 = "List of Employees M - Copy.xls"
xl1 = pd.ExcelFile(file1)
print("Sheet names:")
for sheet in xl1.sheet_names:
    print(f"- {sheet}")
    df = pd.read_excel(file1, sheet_name=sheet)
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"First 10 rows:")
    print(df.head(10).to_string())
    print("\n" + "="*50 + "\n")

print("\n=== Exploring List of Employees NM - Copy.xls ===")
file2 = "List of Employees NM - Copy.xls"
xl2 = pd.ExcelFile(file2)
print("Sheet names:")
for sheet in xl2.sheet_names:
    print(f"- {sheet}")
    df = pd.read_excel(file2, sheet_name=sheet)
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"First 10 rows:")
    print(df.head(10).to_string())
    print("\n" + "="*50 + "\n")
