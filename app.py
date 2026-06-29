import streamlit as st

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

/* Card styling for containers */
.stContainer {
    background-color: white;
    border-radius: 14px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    padding: 20px;
}
</style>
""", unsafe_allow_html=True)

st.set_page_config(
    page_title="Enterprise Dashboards",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Enterprise Dashboard Suite")
st.markdown("---")

st.write("Welcome to the Enterprise Dashboard Suite! Select a dashboard from the sidebar to get started.")
