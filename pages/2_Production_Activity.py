import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re

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

st.set_page_config(page_title="Production Activity Dashboard", page_icon="📋", layout="wide")

st.title("📋 Production Activity Dashboard")
st.markdown("---")

# File upload section
uploaded_file = st.file_uploader("📁 Upload Excel File for Production Activity Data (Expected: Report System 19052026.xls)", type=['xlsx', 'xls'])

default_file_path = "Report System 19052026.xls"

if uploaded_file is not None:
    file_to_use = uploaded_file
    st.success(f"✅ Successfully uploaded: {uploaded_file.name}")
elif os.path.exists(default_file_path):
    file_to_use = default_file_path
    st.info(f"Using default file: {default_file_path}")
else:
    st.error(f"❌ No file uploaded and default file not found at: {default_file_path}")
    st.stop()

# Function to parse production activity data
def parse_production_activity(file_path):
    xl = pd.ExcelFile(file_path)
    all_reports = []
    
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Extract date from first row
            date_str = None
            for i in range(min(5, len(df))):
                row_str = str(df.iloc[i].values)
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})', row_str)
                if date_match:
                    date_str = date_match.group(1)
                    break
            
            report_data = {
                'Sheet': sheet_name,
                'Date': date_str
            }
            
            # Parse key-value pairs from the sheet
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    val = df.iloc[i, j]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        
                        # Check for key metrics
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
                        
                        elif 'Production' in val_str and 'Moulding' in val_str:
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
                        
                        elif 'Efficiency-Moulding' in val_str:
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
            st.warning(f"Error parsing sheet {sheet_name}: {e}")
    
    return pd.DataFrame(all_reports)

# Load and parse data
df_reports = parse_production_activity(file_to_use)

if len(df_reports) > 0:
    st.sidebar.header("📊 Filters")
    selected_sheets = st.sidebar.multiselect(
        "Select Sheets/Days",
        options=sorted(df_reports['Sheet'].unique()),
        default=sorted(df_reports['Sheet'].unique())
    )
    
    df_filtered = df_reports[df_reports['Sheet'].isin(selected_sheets)]
    
    if len(df_filtered) > 0:
        # Key Metrics
        st.subheader("📊 Key Metrics Summary")
        metric_cols = [col for col in df_filtered.columns if col not in ['Sheet', 'Date']]
        
        for metric in metric_cols:
            if df_filtered[metric].notna().any():
                col1, col2, col3 = st.columns(3)
                avg_val = df_filtered[metric].mean()
                max_val = df_filtered[metric].max()
                min_val = df_filtered[metric].min()
                
                col1.metric(f"Average {metric}", f"{avg_val:.2f}")
                col2.metric(f"Maximum {metric}", f"{max_val:.2f}")
                col3.metric(f"Minimum {metric}", f"{min_val:.2f}")
                st.markdown("---")
        
        # Employment status charts
        st.subheader("👥 Employment Status")
        employment_cols = [col for col in ['Total Employment', 'Present Employees', 'Absent Employees', 'Late Employees'] if col in df_filtered.columns]
        
        # Define color palette
        color_palette = ['#2563EB', '#60A5FA', '#38BDF8', '#10B981', '#14B8A6', '#F59E0B', '#F97316', '#EF4444', '#8B5CF6', '#64748B']
        
        if employment_cols:
            # Bar chart for employment
            fig_emp = go.Figure()
            for idx, col in enumerate(employment_cols):
                if df_filtered[col].notna().any():
                    fig_emp.add_trace(go.Bar(
                        x=df_filtered['Sheet'],
                        y=df_filtered[col],
                        name=col,
                        marker_color=color_palette[idx % len(color_palette)]
                    ))
            fig_emp.update_layout(
                template='plotly_white',
                title='Employment Status by Day',
                barmode='group',
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
            st.plotly_chart(fig_emp, use_container_width=True)
        
        st.markdown("---")
        
        # Efficiency charts
        st.subheader("⚡ Department Efficiencies")
        efficiency_cols = [col for col in df_filtered.columns if 'Efficiency' in col]
        
        if efficiency_cols:
            fig_eff = go.Figure()
            for idx, col in enumerate(efficiency_cols):
                if df_filtered[col].notna().any():
                    fig_eff.add_trace(go.Bar(
                        x=df_filtered['Sheet'],
                        y=df_filtered[col],
                        name=col,
                        marker_color=color_palette[idx % len(color_palette)]
                    ))
            fig_eff.update_layout(
                template='plotly_white',
                title='Department Efficiencies by Day',
                barmode='group',
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
            st.plotly_chart(fig_eff, use_container_width=True)
        
        st.markdown("---")
        
        # Production metrics
        st.subheader("🏭 Production Metrics")
        production_cols = [col for col in df_filtered.columns if any(x in col for x in ['Production', 'Compounding', 'Chemlok', 'Preforming', 'Metal Bush'])]
        
        if production_cols:
            for idx, col in enumerate(production_cols):
                if df_filtered[col].notna().any():
                    fig_prod = px.bar(df_filtered, x='Sheet', y=col, title=col, color_discrete_sequence=[color_palette[idx % len(color_palette)]], template='plotly_white')
                    fig_prod.update_layout(
                        height=450,
                        paper_bgcolor='white',
                        plot_bgcolor='white',
                        title_font=dict(size=20, color='#111827'),
                        xaxis_title_font=dict(size=16),
                        yaxis_title_font=dict(size=16),
                        xaxis_tickfont=dict(size=14),
                        yaxis_tickfont=dict(size=14),
                        hovermode='x unified',
                        margin=dict(l=40, r=40, t=60, b=40)
                    )
                    fig_prod.update_traces(
                        marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                        textfont_size=14
                    )
                    st.plotly_chart(fig_prod, use_container_width=True)
        
        st.markdown("---")
        
        # Detailed data
        st.subheader("📋 Detailed Reports")
        st.dataframe(df_filtered, use_container_width=True)
        
        # Download option
        csv = df_filtered.to_csv(index=False)
        st.download_button(
            "📥 Download Data as CSV",
            csv,
            "production_activity_data.csv",
            "text/csv"
        )
        
    else:
        st.warning("No data available with selected filters.")
else:
    st.error("Could not parse any valid production activity data from the Excel file.")
