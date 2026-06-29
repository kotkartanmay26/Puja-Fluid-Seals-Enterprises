import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os
from io import BytesIO

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

st.set_page_config(page_title="Production Efficiency Dashboard", page_icon="📈", layout="wide")

st.title("📈 Production Efficiency Dashboard")
st.markdown("---")

# File upload section
uploaded_file = st.file_uploader("📁 Upload Excel File for Production Efficiency Data (Expected: Best Efficiency Moulding Nov.25 to Jan.26.xlsx)", type=['xlsx', 'xls'])

default_file_path = "Best Efficiency Moulding Nov.25 to Jan.26.xlsx"

if uploaded_file is not None:
    file_to_use = uploaded_file
    st.success(f"✅ Successfully uploaded: {uploaded_file.name}")
elif os.path.exists(default_file_path):
    file_to_use = default_file_path
    st.info(f"Using default file: {default_file_path}")
else:
    st.error(f"❌ No file uploaded and default file not found at: {default_file_path}")
    st.stop()

# Function to load and process data from a sheet
def load_sheet_data(sheet_name, file_path):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # Skip first 2 rows, use row 2 as header
        header_row = df.iloc[2]
        df_data = df.iloc[3:].copy()
        df_data.columns = header_row
        
        # Clean column names
        df_data.columns = df_data.columns.astype(str).str.strip()
        
        # Identify operator name column and date columns
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
        
        # If operator column not found, use the 3rd column (index 2)
        if operator_col is None and len(df_data.columns) > 2:
            operator_col = df_data.columns[2]
        
        if operator_col:
            # Melt the data to long format
            id_vars = [operator_col]
            id_vars = [col for col in id_vars if col in df_data.columns]
            value_vars = [col for col in date_columns if col in df_data.columns]
            
            if id_vars and value_vars:
                df_melted = pd.melt(df_data, id_vars=id_vars, value_vars=value_vars, 
                                   var_name='Date', value_name='Efficiency')
                df_melted['Sheet'] = sheet_name
                
                # Clean efficiency values
                df_melted['Efficiency'] = pd.to_numeric(df_melted['Efficiency'], errors='coerce')
                
                # Remove rows with NaN efficiency
                df_melted = df_melted.dropna(subset=['Efficiency'])
                
                return df_melted
        return None
    except Exception as e:
        st.warning(f"Could not load sheet {sheet_name}: {e}")
        return None

# Load all efficiency sheets
xl = pd.ExcelFile(file_to_use)
efficiency_sheets = [sheet for sheet in xl.sheet_names if 'efficiency' in sheet.lower() or 'nov' in sheet.lower() or 'dec' in sheet.lower() or 'jan' in sheet.lower() or 'feb' in sheet.lower()]
presentation_sheets = [sheet for sheet in xl.sheet_names if 'presentation' in sheet.lower()]

all_data = []

for sheet in efficiency_sheets + presentation_sheets:
    df = load_sheet_data(sheet, file_to_use)
    if df is not None and len(df) > 0:
        all_data.append(df)

if all_data:
    df_combined = pd.concat(all_data, ignore_index=True)
    
    # Find operator column
    operator_col = None
    for col in df_combined.columns:
        if 'operator' in str(col).lower() or 'name' in str(col).lower():
            operator_col = col
            break
    if operator_col is None:
        for col in df_combined.columns:
            if col != 'Date' and col != 'Efficiency' and col != 'Sheet':
                operator_col = col
                break

    if operator_col:
        df_combined = df_combined.rename(columns={operator_col: 'Operator'})
        
        # Sidebar filters
        st.sidebar.header("📊 Filters")
        
        # Sheet filter
        selected_sheets = st.sidebar.multiselect(
            "Select Month/Sheet",
            options=sorted(df_combined['Sheet'].unique()),
            default=sorted(df_combined['Sheet'].unique())
        )
        
        # Operator filter
        selected_operators = st.sidebar.multiselect(
            "Select Operators",
            options=sorted(df_combined['Operator'].dropna().unique()),
            default=sorted(df_combined['Operator'].dropna().unique())[:10]
        )
        
        # Apply filters
        df_filtered = df_combined[
            (df_combined['Sheet'].isin(selected_sheets)) & 
            (df_combined['Operator'].isin(selected_operators))
        ]

        if len(df_filtered) > 0:
            # Key Metrics
            col1, col2, col3, col4 = st.columns(4)
            
            avg_efficiency = df_filtered['Efficiency'].mean()
            median_efficiency = df_filtered['Efficiency'].median()
            max_efficiency = df_filtered['Efficiency'].max()
            min_efficiency = df_filtered['Efficiency'].min()
            
            col1.metric("📊 Average Efficiency", f"{avg_efficiency:.2f}%")
            col2.metric("📈 Median Efficiency", f"{median_efficiency:.2f}%")
            col3.metric("🏆 Maximum Efficiency", f"{max_efficiency:.2f}%")
            col4.metric("📉 Minimum Efficiency", f"{min_efficiency:.2f}%")

            st.markdown("---")
            
            # Top row of charts
            col_a, col_b = st.columns(2)
            
            # Define color palette
            color_palette = ['#2563EB', '#60A5FA', '#38BDF8', '#10B981', '#14B8A6', '#F59E0B', '#F97316', '#EF4444', '#8B5CF6', '#64748B']
            
            with col_a:
                st.subheader("📈 Average Efficiency by Month")
                sheet_avg = df_filtered.groupby('Sheet')['Efficiency'].mean().reset_index()
                fig_sheet = px.bar(sheet_avg, x='Sheet', y='Efficiency',
                                     title='Average Efficiency by Month',
                                     color='Efficiency',
                                     color_continuous_scale=color_palette,
                                     text_auto='.2f',
                                     template='plotly_white')
                fig_sheet.update_layout(
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
                fig_sheet.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_sheet, use_container_width=True)

            with col_b:
                st.subheader("👥 Top 10 Operators by Average Efficiency")
                operator_avg = df_filtered.groupby('Operator')['Efficiency'].mean().reset_index()
                operator_avg = operator_avg.sort_values('Efficiency', ascending=False).head(10)
                fig_top_ops = px.bar(operator_avg, x='Operator', y='Efficiency',
                                    title='Top 10 Operators',
                                    color='Efficiency',
                                    color_continuous_scale=color_palette,
                                    text_auto='.2f',
                                    template='plotly_white')
                fig_top_ops.update_layout(
                    height=500,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font=dict(size=20, color='#111827'),
                    xaxis_title_font=dict(size=16),
                    yaxis_title_font=dict(size=16),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    xaxis_tickangle=-45,
                    legend_font=dict(size=14),
                    hovermode='x unified',
                    margin=dict(l=40, r=40, t=60, b=40)
                )
                fig_top_ops.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_top_ops, use_container_width=True)

            st.markdown("---")

            # Efficiency distribution
            st.subheader("📊 Efficiency Distribution")
            fig_dist = px.histogram(df_filtered, x='Efficiency', nbins=20,
                                     title='Distribution of Efficiency Scores',
                                     color_discrete_sequence=['#2563EB'],
                                     template='plotly_white')
            fig_dist.update_layout(
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
            fig_dist.update_traces(
                marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                textfont_size=14
            )
            st.plotly_chart(fig_dist, use_container_width=True)

            st.markdown("---")

            # Operator performance heatmap
            st.subheader("🔥 Operator Efficiency Heatmap")
            pivot_df = df_filtered.pivot_table(values='Efficiency', index='Operator', columns='Sheet', aggfunc='mean')
            fig_heatmap = px.imshow(pivot_df,
                                     title='Operator Efficiency Across Months',
                                     color_continuous_scale=color_palette,
                                     aspect='auto',
                                     template='plotly_white')
            fig_heatmap.update_layout(
                height=650,
                paper_bgcolor='white',
                plot_bgcolor='white',
                title_font=dict(size=20, color='#111827'),
                xaxis_title_font=dict(size=16),
                yaxis_title_font=dict(size=16),
                xaxis_tickfont=dict(size=14),
                yaxis_tickfont=dict(size=14),
                legend_font=dict(size=14),
                margin=dict(l=40, r=40, t=60, b=40)
            )
            st.plotly_chart(fig_heatmap, use_container_width=True)

            st.markdown("---")

            # Detailed data
            st.subheader("📋 Detailed Data")
            
            # Show operator-wise summary
            st.write("### Operator-wise Summary")
            operator_summary = df_filtered.groupby(['Sheet', 'Operator']).agg({
                'Efficiency': ['mean', 'median', 'count', 'max', 'min']
            }).round(2)
            operator_summary.columns = ['Average', 'Median', 'Data Points', 'Maximum', 'Minimum']
            st.dataframe(operator_summary, use_container_width=True)

            # Show raw data
            with st.expander("View Raw Data"):
                st.dataframe(df_filtered, use_container_width=True)

        else:
            st.warning("No data available with the selected filters. Please adjust your filters.")
    else:
        st.error("Could not identify operator column in the data.")
else:
    st.error("Could not load any valid data from the Excel file.")
