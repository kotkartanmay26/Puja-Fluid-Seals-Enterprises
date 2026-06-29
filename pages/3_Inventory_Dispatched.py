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

st.set_page_config(page_title="Inventory Dispatched Dashboard", page_icon="📦", layout="wide")

st.title("📦 Inventory Dispatched Dashboard")
st.markdown("---")

# File upload section
uploaded_file = st.file_uploader("📁 Upload Excel File for Inventory Data (Expected: Total Inventory Dispatched 2025-26.xlsx)", type=['xlsx', 'xls'])

default_file_path = "Total Inventory Dispatched 2025-26.xlsx"

if uploaded_file is not None:
    file_to_use = uploaded_file
    st.success(f"✅ Successfully uploaded: {uploaded_file.name}")
elif os.path.exists(default_file_path):
    file_to_use = default_file_path
    st.info(f"Using default file: {default_file_path}")
else:
    st.error(f"❌ No file uploaded and default file not found at: {default_file_path}")
    st.stop()

# Load data
try:
    df = pd.read_excel(file_to_use)
    # Clean column names - strip whitespace and make case-insensitive
    df.columns = [str(col).strip() for col in df.columns]
    st.sidebar.header("📊 Filters")
    
    # Customer filter
    # Find customer column by name variations
    customer_col = None
    for col in df.columns:
        col_lower = col.lower()
        if 'customer' in col_lower:
            customer_col = col
            break
    
    if customer_col and df[customer_col].notna().any():
        customers = sorted(df[customer_col].dropna().unique())
        selected_customers = st.sidebar.multiselect(
            "Select Customers",
            options=customers,
            default=customers  # Select all by default!
        )
    else:
        selected_customers = []
    
    # Item filter
    # Find item column by name variations
    item_col = None
    for col in df.columns:
        col_lower = col.lower()
        if 'item' in col_lower or 'code' in col_lower:
            item_col = col
            break
    
    if item_col and df[item_col].notna().any():
        items = sorted(df[item_col].dropna().unique())
        selected_items = st.sidebar.multiselect(
            "Select Item Codes",
            options=items,
            default=items  # Select all by default!
        )
    else:
        selected_items = []
    
    # Apply filters
    df_filtered = df.copy()
    if selected_customers and customer_col:
        df_filtered = df_filtered[df_filtered[customer_col].isin(selected_customers)]
    if selected_items and item_col:
        df_filtered = df_filtered[df_filtered[item_col].isin(selected_items)]
    
    if len(df_filtered) > 0:
        # Key Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        # Find Qty column
        qty_col = None
        for col in df.columns:
            if 'qty' in col.lower() or 'quantity' in col.lower():
                qty_col = col
                break
        if not qty_col and 'Qty' in df.columns:
            qty_col = 'Qty'
        
        # Find Amount column
        amount_col = None
        for col in df.columns:
            if 'amount' in col.lower() or 'value' in col.lower() or 'total' in col.lower():
                amount_col = col
                break
        if not amount_col and 'Amount' in df.columns:
            amount_col = 'Amount'
        
        total_qty = df_filtered[qty_col].sum() if qty_col is not None and qty_col in df.columns else 0
        total_amount = df_filtered[amount_col].sum() if amount_col is not None and amount_col in df.columns else 0
        unique_customers = df_filtered[customer_col].nunique() if customer_col is not None and customer_col in df.columns else 0
        unique_items = df_filtered[item_col].nunique() if item_col is not None and item_col in df.columns else 0
        
        col1.metric("📦 Total Quantity Dispatched", f"{total_qty:,.0f}")
        col2.metric("💰 Total Amount", f"₹{total_amount:,.2f}")
        col3.metric("👥 Unique Customers", f"{unique_customers}")
        col4.metric("📦 Unique Items", f"{unique_items}")
        
        st.markdown("---")
        
        # Top charts
        col_a, col_b = st.columns(2)
        
        # Define color palette
        color_palette = ['#2563EB', '#60A5FA', '#38BDF8', '#10B981', '#14B8A6', '#F59E0B', '#F97316', '#EF4444', '#8B5CF6', '#64748B']
        
        with col_a:
            if customer_col is not None and customer_col in df.columns and amount_col is not None and amount_col in df.columns:
                st.subheader("🏆 Top 10 Customers by Amount")
                customer_amount = df_filtered.groupby(customer_col)[amount_col].sum().reset_index()
                customer_amount = customer_amount.sort_values(amount_col, ascending=False).head(10)
                fig_customer = px.bar(customer_amount, x=customer_col, y=amount_col,
                                     title='Top 10 Customers', color=amount_col,
                                     color_continuous_scale=color_palette, text_auto=',.0f',
                                     template='plotly_white')
                fig_customer.update_layout(
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
                fig_customer.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_customer, use_container_width=True)
        
        with col_b:
            if item_col is not None and item_col in df.columns and amount_col is not None and amount_col in df.columns:
                st.subheader("📦 Top 10 Items by Amount")
                item_amount = df_filtered.groupby(item_col)[amount_col].sum().reset_index()
                item_amount = item_amount.sort_values(amount_col, ascending=False).head(10)
                fig_item = px.bar(item_amount, x=item_col, y=amount_col,
                                    title='Top 10 Items', color=amount_col,
                                    color_continuous_scale=color_palette, text_auto=',.0f',
                                    template='plotly_white')
                fig_item.update_layout(
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
                fig_item.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_item, use_container_width=True)
        
        st.markdown("---")
        
        # More visualizations
        col_c, col_d = st.columns(2)
        
        with col_c:
            if customer_col is not None and customer_col in df.columns and qty_col is not None and qty_col in df.columns:
                st.subheader("📦 Top 10 Customers by Quantity")
                customer_qty = df_filtered.groupby(customer_col)[qty_col].sum().reset_index()
                customer_qty = customer_qty.sort_values(qty_col, ascending=False).head(10)
                fig_customer_qty = px.bar(customer_qty, x=customer_col, y=qty_col,
                                          title='Top 10 Customers by Qty', color=qty_col,
                                          color_continuous_scale=color_palette, text_auto=',.0f',
                                          template='plotly_white')
                fig_customer_qty.update_layout(
                    height=450,
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
                fig_customer_qty.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_customer_qty, use_container_width=True)
        
        with col_d:
            if item_col is not None and item_col in df.columns and qty_col is not None and qty_col in df.columns:
                st.subheader("🏷️ Top 10 Items by Quantity")
                item_qty = df_filtered.groupby(item_col)[qty_col].sum().reset_index()
                item_qty = item_qty.sort_values(qty_col, ascending=False).head(10)
                fig_item_qty = px.bar(item_qty, x=item_col, y=qty_col,
                                      title='Top 10 Items by Qty', color=qty_col,
                                      color_continuous_scale=color_palette, text_auto=',.0f',
                                      template='plotly_white')
                fig_item_qty.update_layout(
                    height=450,
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
                fig_item_qty.update_traces(
                    marker=dict(line=dict(width=1, color='rgba(0,0,0,0.1)')),
                    textfont_size=14
                )
                st.plotly_chart(fig_item_qty, use_container_width=True)
        
        st.markdown("---")
        
        # Detailed data
        st.subheader("📋 Detailed Dispatched Inventory")
        st.dataframe(df_filtered, use_container_width=True)
        
        # Download option
        csv = df_filtered.to_csv(index=False)
        st.download_button(
            "📥 Download Data as CSV",
            csv,
            "inventory_dispatched.csv",
            "text/csv"
        )
        
    else:
        st.warning("No data available with selected filters.")
        
except Exception as e:
    st.error(f"Error loading data: {e}")
    import traceback
    st.error(f"Details: {traceback.format_exc()}")
