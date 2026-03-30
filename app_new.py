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
category = st.sidebar.selectbox("Select Category", ["All"] + list(df["Category"].unique()))

if category != "All":
    df = df[df["Category"] == category]

product = st.sidebar.selectbox("Select Product", ["All"] + list(df["Product"].unique()))

if product != "All":
    df = df[df["Product"] == product]

# KPI Metrics
total_sales = df["Amount"].sum()
avg_sales = df["Amount"].mean()
total_products = df["Product"].nunique()

col1, col2, col3 = st.columns(3)

col1.metric("Total Sales", f"${total_sales:,.2f}")
col2.metric("Average Sales", f"${avg_sales:.2f}")
col3.metric("Products", total_products)

# Charts
col4, col5 = st.columns(2)

with col4:
    fig1 = px.bar(df, x="Product", y="Amount", color="Category", title="Sales by Product")
    st.plotly_chart(fig1, use_container_width=True)

with col5:
    fig2 = px.pie(df, names="Category", values="Amount", title="Sales by Category")
    st.plotly_chart(fig2, use_container_width=True)

# Data Table
st.subheader("Sales Data")
st.dataframe(df)