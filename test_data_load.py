import pandas as pd
import os
import re

# Test first file
print("Testing first file...")
file1 = "Best Efficiency Moulding Nov.25 to Jan.26.xlsx"

def load_sheet_data(sheet_name, file_path):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row = df.iloc[2]
        df_data = df.iloc[3:].copy()
        df_data.columns = header_row
        df_data.columns = df_data.columns.astype(str).str.strip()
        
        operator_col = None
        date_columns = []
        for col in df_data.columns:
            col_str = str(col).lower()
            if 'operator' in col_str or 'name' in col_str:
                operator_col = col
            try:
                float(col)
                date_columns.append(col)
            except:
                pass
        
        if operator_col is None and len(df_data.columns) > 2:
            operator_col = df_data.columns[2]
        
        if operator_col:
            id_vars = [operator_col]
            id_vars = [col for col in id_vars if col in df_data.columns]
            value_vars = [col for col in date_columns if col in df_data.columns]
            if id_vars and value_vars:
                df_melted = pd.melt(df_data, id_vars=id_vars, value_vars=value_vars, 
                                   var_name='Date', value_name='Efficiency')
                df_melted['Sheet'] = sheet_name
                df_melted['Efficiency'] = pd.to_numeric(df_melted['Efficiency'], errors='coerce')
                df_melted = df_melted.dropna(subset=['Efficiency'])
                return df_melted
        return None
    except Exception as e:
        print(f"Error loading sheet {sheet_name}: {e}")
        return None

xl1 = pd.ExcelFile(file1)
all_data = []
for sheet in xl1.sheet_names:
    df = load_sheet_data(sheet, file1)
    if df is not None and len(df) >0:
        all_data.append(df)
print(f"Loaded {len(all_data)} sheets from first file!")
if all_data:
    df1 = pd.concat(all_data, ignore_index=True)
    print(f"First file data shape: {df1.shape}")

# Test second file
print("\nTesting second file...")
file2 = "Report System 19052026.xls"

def parse_production_activity(file_path):
    xl = pd.ExcelFile(file_path)
    all_reports = []
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            date_str = None
            for i in range(min(5, len(df))):
                row_str = str(df.iloc[i].values)
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})', row_str)
                if date_match:
                    date_str = date_match.group(1)
                    break
            
            report_data = {'Sheet': sheet_name, 'Date': date_str}
            
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    val = df.iloc[i, j]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if 'Compounding' in val_str and 'Weight' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Compounding Weight (Kg)'] = float(num_val.group(1))
                        elif 'Chemlok' in val_str and 'Appllied' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Chemlok Applied (Nos)'] = float(num_val.group(1))
                        elif 'Metal Bush' in val_str and 'Production' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Metal Bush Production (Nos)'] = float(num_val.group(1))
                        elif 'Production' in val_str and 'Molding' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Production - Molding (Nos)'] = float(num_val.group(1))
                        elif 'Total Employement' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Total Employment'] = int(val_num)
                        elif 'Present' in val_str and not 'Preforming' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Present Employees'] = int(val_num)
                        elif 'Absent' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Absent Employees'] = int(val_num)
                        elif 'Late' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Late Employees'] = int(val_num)
                        elif 'Efficiency-Chemlok' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Efficiency - Chemlok (%)'] = float(val_num)
                        elif 'Efficiency-Compounding' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Efficiency - Compounding (%)'] = float(val_num)
                        elif 'Efficiency-Toolroom' in val_str and 'production' not in val_str.lower():
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Efficiency - Toolroom (%)'] = float(val_num)
                        elif 'Efficiency-Molding' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Efficiency - Molding (%)'] = float(val_num)
                        elif 'Energy Units' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Energy Consumption (Units)'] = float(num_val.group(1))
                        elif 'Scrap-Rubber' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Scrap - Rubber (Kg)'] = float(num_val.group(1))
                        elif 'Sales Rs' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                val_num = df.iloc[i, j+1]
                                if isinstance(val_num, (int, float)):
                                    report_data['Sales (Rs)'] = float(val_num)
                        elif 'Preforming' in val_str and 'Weight' in val_str:
                            if j+1 < len(df.columns) and pd.notna(df.iloc[i, j+1]):
                                next_val = str(df.iloc[i, j+1]).strip()
                                num_val = re.search(r'(\d+\.?\d*)', next_val)
                                if num_val:
                                    report_data['Total Preforming Weight (Kg)'] = float(num_val.group(1))
            
            if len(report_data) > 2:
                all_reports.append(report_data)
                
        except Exception as e:
            print(f"Error parsing sheet {sheet_name}: {e}")
    
    return pd.DataFrame(all_reports)

df_reports = parse_production_activity(file2)
print(f"Loaded {len(df_reports)} reports from second file!")
if len(df_reports) >0:
    print(f"Second file data shape: {df_reports.shape}")
    print("\nSample data:")
    print(df_reports.head())

print("\n✅ All tests passed!")
