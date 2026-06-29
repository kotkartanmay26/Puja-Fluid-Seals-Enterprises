import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Custom CSS
st.markdown("""
<style>
/* Main background */
[data-testid="stAppViewContainer"] {
    background-color: #F3F6FB;
}

/* Sidebar styling */
[data-testid="stSidebar"] {
    background-color: #0F172A;
    padding: 20px 15px;
}

[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    color: #FFFFFF !important;
}

/* Global font styling */
html, body, [class*="css"] {
    font-family: 'Segoe UI', Roboto, sans-serif;
}

/* Title styling */
h1 {
    font-size: 40px !important;
    font-weight: 700;
    color: #111827;
    margin-bottom: 24px;
}

h2, h3 {
    font-size: 26px !important;
    font-weight: 700;
    color: #111827;
    margin-top: 32px;
    margin-bottom: 20px;
}

/* Normal text */
p, div, span {
    font-size: 16px !important;
    color: #475569;
}

/* Button styling */
button[kind="primary"] {
    background-color: #2563EB !important;
    border-radius: 12px !important;
    border: none !important;
    padding: 12px 24px !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    transition: all 0.3s ease !important;
}

button[kind="primary"]:hover {
    background-color: #1D4ED8 !important;
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.4) !important;
}

/* Divider styling */
hr {
    border: none;
    border-top: 1px solid #E2E8F0;
    margin: 32px 0;
}

/* Sidebar text */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] div {
    font-size: 16px !important;
}

/* Filter widgets */
[data-testid="stMultiSelect"] label,
[data-testid="stSelectbox"] label {
    font-size: 15px !important;
    font-weight: 500;
    color: #FFFFFF;
}

/* Fix sidebar select/multiselect text colors */
[data-testid="stSidebar"] [data-testid="stMultiSelect"],
[data-testid="stSidebar"] [data-testid="stSelectbox"],
[data-testid="stSidebar"] [data-testid="stTextInput"] {
    background-color: rgba(255,255,255,0.1);
    border-radius: 12px;
}

[data-testid="stSidebar"] [data-testid="stMultiSelect"] div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] div {
    color: #FFFFFF !important;
}

/* KPI Metric Cards */
[data-testid="stMetric"] {
    background-color: white;
    border-radius: 14px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    padding: 24px;
    margin: 0;
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

[data-testid="stMetric"]:hover {
    transform: translateY(-4px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.12);
}

[data-testid="stMetricValue"] {
    font-size: 34px !important;
    font-weight: 700;
    color: #2563EB;
}

[data-testid="stMetricLabel"] p {
    font-size: 16px !important;
    font-weight: 500;
    color: #64748B;
}

/* Upload box styling */
[data-testid="stFileUploaderDropzone"] {
    background-color: white;
    border: 2px dashed #2563EB;
    border-radius: 12px;
    padding: 32px 24px;
}

[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #1D4ED8;
}

/* DataFrame styling */
div[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}

div[data-testid="stDataFrame"] th {
    background-color: #F1F5F9 !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    color: #334155 !important;
    padding: 16px !important;
    position: sticky;
    top: 0;
}

div[data-testid="stDataFrame"] td {
    font-size: 15px !important;
    padding: 14px !important;
}

div[data-testid="stDataFrame"] tr:nth-child(even) {
    background-color: #F8FAFC !important;
}

div[data-testid="stDataFrame"] tr:nth-child(odd) {
    background-color: white !important;
}

div[data-testid="stDataFrame"] tr:hover {
    background-color: #E6F0FF !important;
}
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Employee Management Dashboard", page_icon="👥", layout="wide")

st.title("👥 Employee Management Dashboard")
st.markdown("---")

# Single file upload option
uploaded_file = st.file_uploader("📁 Upload Employee List Excel File (Expected: List of Employees M - Copy.xls or List of Employees NM - Copy.xls)", type=['xlsx', 'xls'])

default_file_path = "List of Employees M - Copy.xls"

if uploaded_file is not None:
    file_to_use = uploaded_file
    st.success(f"✅ Successfully uploaded: {uploaded_file.name}")
elif os.path.exists(default_file_path):
    file_to_use = default_file_path
    st.info(f"Using default file: {default_file_path}")
elif os.path.exists("List of Employees NM - Copy.xls"):
    file_to_use = "List of Employees NM - Copy.xls"
    st.info(f"Using default file: List of Employees NM - Copy.xls")
else:
    st.error("❌ No file uploaded and no default files found.")
    st.stop()

# Function to load and clean employee data
def load_employee_data(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        all_dfs = []
        
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Clean header - find the row with SR.NO. or similar
            header_row = None
            for i in range(min(10, len(df))):
                row_str = str(df.iloc[i].values).lower()
                if 'sr.no' in row_str or 'sr no' in row_str or 'name' in row_str:
                    header_row = i
                    break
            
            if header_row is not None:
                # Use the found header row
                df.columns = df.iloc[header_row]
                df = df.iloc[header_row+1:].reset_index(drop=True)
                # Clean column names and make them unique
                cleaned_cols = []
                seen = {}
                for col in df.columns:
                    col_str = str(col).strip()
                    if col_str == 'nan' or not col_str:
                        col_str = f'Column_{len(cleaned_cols)+1}'
                    
                    if col_str in seen:
                        seen[col_str] += 1
                        col_str = f"{col_str}_{seen[col_str]}"
                    else:
                        seen[col_str] = 0
                    
                    cleaned_cols.append(col_str)
                
                df.columns = cleaned_cols
                # Drop empty rows
                df = df.dropna(how='all')
                if len(df) > 0:
                    all_dfs.append(df)
        
        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True)
        else:
            # Fallback: just load first sheet
            df = pd.read_excel(file_path)
            # Clean columns even for fallback
            cleaned_cols = []
            seen = {}
            for col in df.columns:
                col_str = str(col).strip()
                if col_str == 'nan' or not col_str:
                    col_str = f'Column_{len(cleaned_cols)+1}'
                
                if col_str in seen:
                    seen[col_str] += 1
                    col_str = f"{col_str}_{seen[col_str]}"
                else:
                    seen[col_str] = 0
                
                cleaned_cols.append(col_str)
            
            df.columns = cleaned_cols
            return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        import traceback
        st.error(f"Details: {traceback.format_exc()}")
        return pd.DataFrame()

# Load data
df = load_employee_data(file_to_use)

if len(df) > 0:
    st.sidebar.header("📊 Filters")
    
    # Display column names for debugging (optional)
    # st.sidebar.write("Columns:", list(df.columns))

    # Find department column
    dept_col = None
    for col in df.columns:
        col_str = str(col).lower()
        if 'department' in col_str or 'dept' in col_str:
            dept_col = col
            break

    # If not found, try a fallback logic
    if not dept_col and len(df.columns) > 2:
        dept_col = df.columns[2]  # Try 3rd column (index 2)

    if dept_col and dept_col in df.columns and df[dept_col].notna().any():
        departments = sorted(df[dept_col].dropna().unique())
        selected_depts = st.sidebar.multiselect(
            "Select Departments",
            options=departments,
            default=departments
        )
    else:
        selected_depts = []
    
    # Apply filter
    df_filtered = df.copy()
    if selected_depts and dept_col and dept_col in df.columns:
        df_filtered = df_filtered[df_filtered[dept_col].isin(selected_depts)]
    
    if len(df_filtered) > 0:
        # Key Metrics
        col1, col2, col3 = st.columns(3)
        
        total_employees = len(df_filtered)
        total_depts = df_filtered[dept_col].nunique() if dept_col is not None else 0
        
        col1.metric("👥 Total Employees", total_employees)
        col2.metric("🏢 Total Departments", total_depts)
        col3.metric("📊 Total Columns", len(df_filtered.columns))
        
        st.markdown("---")
        
        # Charts
        col_a, col_b = st.columns(2)
        
        # Define color palette
        color_palette = ['#2563EB', '#60A5FA', '#38BDF8', '#10B981', '#14B8A6', '#F59E0B', '#F97316', '#EF4444', '#8B5CF6', '#64748B']
        
        with col_a:
            if dept_col is not None and dept_col in df.columns:
                st.subheader("🏢 Department Distribution")
                dept_counts = df_filtered[dept_col].value_counts().reset_index()
                dept_counts.columns = [dept_col, 'Count']
                fig_dept = px.bar(dept_counts, x=dept_col, y='Count',
                                  title='Employees by Department', color='Count',
                                  color_continuous_scale=color_palette, text_auto=True,
                                  template='plotly_white')
                fig_dept.update_layout(
                    height=500,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font=dict(size=20, color='#111827'),
                    xaxis_title_font=dict(size=16),
                    yaxis_title_font=dict(size=16),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    legend_font=dict(size=14),
                    xaxis_tickangle=-45,
                    hovermode='x unified',
                    margin=dict(l=40, r=40, t=60, b=40)
                )
                fig_dept.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_dept, use_container_width=True)
        
        with col_b:
            # Try to find an experience or numeric column for histogram
            num_cols = df_filtered.select_dtypes(include=['int64', 'float64']).columns
            if len(num_cols) > 0:
                exp_col = num_cols[0]
                st.subheader("📊 Experience Distribution")
                fig_exp = px.histogram(df_filtered, x=exp_col, title=f'{exp_col} Distribution',
                                     color_discrete_sequence=['#2563EB'],
                                     template='plotly_white')
                fig_exp.update_layout(
                    height=500,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font=dict(size=20, color='#111827'),
                    xaxis_title_font=dict(size=16),
                    yaxis_title_font=dict(size=16),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    legend_font=dict(size=14),
                    hovermode='x unified',
                    margin=dict(l=40, r=40, t=60, b=40)
                )
                fig_exp.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_exp, use_container_width=True)
        
        st.markdown("---")
        
        # Detailed data
        st.subheader("📋 Employee Details")
        st.dataframe(df_filtered, use_container_width=True)
        
        # Download option
        csv = df_filtered.to_csv(index=False)
        st.download_button(
            "📥 Download Data as CSV",
            csv,
            "employee_data.csv",
            "text/csv"
        )
        
    else:
        st.warning("No data available with selected filters.")
else:
    st.error("Could not load valid employee data.")
