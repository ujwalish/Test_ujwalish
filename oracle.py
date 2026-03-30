import streamlit as st
import pandas as pd
import plotly.express as px

# Page configuration
st.set_page_config(page_title="Oracle Report Dashboard", layout="wide")

# Title and description
st.title("📊 Oracle Report Interactive Dashboard")
st.markdown("View and analyze subscription account data with interactive filters and visualizations")

# Load data
@st.cache_data
def load_data():
    try:
        df = pd.read_csv("27-Mar-2026 Oracle Report!.csv", encoding='latin1')
    except Exception as e:
        try:
            df = pd.read_csv("27-Mar-2026 Oracle Report!.csv", encoding='iso-8859-1')
        except:
            df = pd.read_csv("27-Mar-2026 Oracle Report!.csv", encoding='cp1252')
    return df

df = load_data()

# Sidebar filters
st.sidebar.header("🔍 Filters")

# Status filter
account_status = st.sidebar.multiselect(
    "Account Status",
    options=df["Account Status"].unique(),
    default=df["Account Status"].unique()
)

# Sector filter
sector = st.sidebar.multiselect(
    "Sector",
    options=sorted(df["Sector"].dropna().unique()),
    default=None
)

# Connectivity filter
connectivity = st.sidebar.multiselect(
    "Connectivity Status",
    options=df["Connectivity Status"].unique(),
    default=df["Connectivity Status"].unique()
)

# Apply filters
filtered_df = df[
    (df["Account Status"].isin(account_status)) &
    (df["Connectivity Status"].isin(connectivity))
]

if sector:
    filtered_df = filtered_df[filtered_df["Sector"].isin(sector)]

# Search box
search_term = st.sidebar.text_input("🔎 Search by User/Company Name")
if search_term:
    filtered_df = filtered_df[
        (filtered_df["User Name"].str.contains(search_term, case=False, na=False)) |
        (filtered_df["Subscription Name"].str.contains(search_term, case=False, na=False))
    ]

st.sidebar.divider()
st.sidebar.info(f"Showing {len(filtered_df)} of {len(df)} records")

# Display data summary (AFTER filters applied)
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    total_records = filtered_df["User Subscription Id"].nunique()
    st.metric("Total Records", total_records)
with col2:
    unique_orgs = filtered_df["Subscription Name"].nunique()
    st.metric("Unique Organizations", unique_orgs)
with col3:
    active_count = len(filtered_df[filtered_df["Account Status"] == "Active"])
    st.metric("Active Accounts", active_count)
with col4:
    inactive_count = len(filtered_df[filtered_df["Account Status"] == "Inactive"])
    st.metric("Inactive Accounts", inactive_count)
with col5:
    total_amount = filtered_df["Amount"].sum()
    st.metric("Total Amount", f"NPR {total_amount:,.0f}")

# ISP Sector Card
st.subheader("ISP Sector Summary")
col_isp1, col_isp2, col_isp3 = st.columns(3)
isp_df = filtered_df[filtered_df["Sector"] == "ISP"]
with col_isp1:
    isp_count = isp_df["User Subscription Id"].nunique()
    st.metric("ISP Records", isp_count)
with col_isp2:
    isp_orgs = isp_df["Subscription Name"].nunique()
    st.metric("ISP Organizations", isp_orgs)
with col_isp3:
    isp_amount = isp_df["Amount"].sum()
    st.metric("ISP Total Amount", f"NPR {isp_amount:,.0f}")

st.divider()

# Tabs for different views
tab1, tab2, tab3, tab4 = st.tabs(["📋 Data Table", "📈 Analytics", "🗺️ Regional View", "💰 Financial"])

# Tab 1: Data Table
with tab1:
    st.subheader("Subscription Data")
    
    # Column selector for display
    display_cols = st.multiselect(
        "Select columns to display:",
        options=df.columns.tolist(),
        default=["Subscription Name", "User Name", "Account Status", "Sector", "Amount", "Region Of Branch", "Connectivity Status"]
    )
    
    st.dataframe(
        filtered_df[display_cols],
        use_container_width=True,
        height=400
    )
    
    # Download button
    csv = filtered_df.to_csv(index=False)
    st.download_button(
        label="📥 Download as CSV",
        data=csv,
        file_name="oracle_report_filtered.csv",
        mime="text/csv"
    )

# Tab 2: Analytics
with tab2:
    col1, col2 = st.columns(2)
    
    with col1:
        # Accounts by Status
        status_counts = filtered_df["Account Status"].value_counts()
        fig_status = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title="Account Status Distribution",
            color_discrete_map={"Active": "#28a745", "Inactive": "#dc3545"}
        )
        st.plotly_chart(fig_status, use_container_width=True)
    
    with col2:
        # Accounts by Sector
        sector_counts = filtered_df["Sector"].value_counts().head(10)
        df_sector = pd.DataFrame({
            "Sector": sector_counts.index,
            "Count": sector_counts.values
        })
        fig_sector = px.bar(
            df_sector,
            x="Count",
            y="Sector",
            orientation="h",
            title="Top 10 Sectors"
        )
        st.plotly_chart(fig_sector, use_container_width=True)
    
    col3, col4 = st.columns(2)
    
    with col3:
        # Amount by Sector
        amount_by_sector = filtered_df.groupby("Sector")["Amount"].sum().sort_values(ascending=False).head(10)
        df_amt = pd.DataFrame({
            "Sector": amount_by_sector.index,
            "Amount": amount_by_sector.values
        })
        fig_amount = px.bar(
            df_amt,
            x="Sector",
            y="Amount",
            title="Top 10 Sectors by Total Amount"
        )
        fig_amount.update_xaxes(tickangle=-45)
        st.plotly_chart(fig_amount, use_container_width=True)
    
    with col4:
        # Connectivity Status
        conn_counts = filtered_df["Connectivity Status"].value_counts()
        fig_conn = px.pie(
            values=conn_counts.values,
            names=conn_counts.index,
            title="Connectivity Status"
        )
        st.plotly_chart(fig_conn, use_container_width=True)

# Tab 3: Regional View
with tab3:
    col1, col2 = st.columns(2)
    
    with col1:
        # Accounts by Region
        region_counts = filtered_df["Region Of Branch"].value_counts().head(12)
        df_region = pd.DataFrame({
            "Region": region_counts.index,
            "Count": region_counts.values
        })
        fig_region = px.bar(
            df_region,
            x="Count",
            y="Region",
            orientation="h",
            title="Accounts by Region"
        )
        st.plotly_chart(fig_region, use_container_width=True)
    
    with col2:
        # Amount by Region
        amount_by_region = filtered_df.groupby("Region Of Branch")["Amount"].sum().sort_values(ascending=False).head(12)
        df_reg_amt = pd.DataFrame({
            "Region": amount_by_region.index,
            "Amount": amount_by_region.values
        })
        fig_amt_region = px.bar(
            df_reg_amt,
            x="Amount",
            y="Region",
            orientation="h",
            title="Total Amount by Region"
        )
        st.plotly_chart(fig_amt_region, use_container_width=True)

# Tab 4: Financial
with tab4:
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Total Amount", f"NPR {filtered_df['Amount'].sum():,.0f}")
        st.metric("Average Amount", f"NPR {filtered_df['Amount'].mean():,.0f}")
        st.metric("Max Amount", f"NPR {filtered_df['Amount'].max():,.0f}")
    
    with col2:
        st.metric("Min Amount", f"NPR {filtered_df['Amount'].min():,.0f}")
        st.metric("Total Dealt Amount", f"NPR {filtered_df['Dealt Amount'].sum():,.0f}")
        st.metric("Active Accounts Value", f"NPR {filtered_df[filtered_df['Account Status'] == 'Active']['Amount'].sum():,.0f}")
    
    st.divider()
    
    # Amount distribution
    fig_dist = px.histogram(
        filtered_df[filtered_df['Amount'] > 0],
        x="Amount",
        nbins=50,
        title="Amount Distribution",
        labels={"Amount": "Amount (NPR)"}
    )
    st.plotly_chart(fig_dist, use_container_width=True)
