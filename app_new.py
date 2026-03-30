import streamlit as st
import pandas as pd
import plotly.express as px

# Page title
st.set_page_config(page_title="Sales Dashboard", layout="wide")
st.title("Sales Performance Dashboard")

# Load data
df = pd.read_csv("sales_test.csv")

# Sidebar filter
st.sidebar.header("Filters")
region = st.sidebar.selectbox("Select Region", ["All"] + list(df["Region"].unique()))

if region != "All":
    df = df[df["Region"] == region]

customer = st.sidebar.selectbox("Select Customer", ["All"] + list(df["Customer"].unique()))

if customer != "All":
    df = df[df["Customer"] == customer]

# KPI Metrics
total_sales = df["Sales"].sum()
avg_sales = df["Sales"].mean()
total_customers = df["Customer"].nunique()

col1, col2, col3 = st.columns(3)

col1.metric("Total Sales", f"{total_sales}")
col2.metric("Average Sales", f"{avg_sales:.2f}")
col3.metric("Customers", total_customers)

# Charts
col4, col5 = st.columns(2)

with col4:
    fig1 = px.bar(df, x="Customer", y="Sales", color="Region", title="Sales by Customer")
    st.plotly_chart(fig1, use_container_width=True)

with col5:
    fig2 = px.pie(df, names="Region", values="Sales", title="Sales by Region")
    st.plotly_chart(fig2, use_container_width=True)

# Data Table
st.subheader("Sales Data")
st.dataframe(df)