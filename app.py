# ==========================================================
# SALES ORDER COMPLETE ANALYTICS DASHBOARD
# (Ultimate Full Version – All Features Added)
# ==========================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import re
import numpy as np
from datetime import datetime

st.set_page_config(page_title="Sales Order Dashboard", layout="wide")
st.title("📊 Complete Sales Order Analytics Dashboard")

# ----------------------------------------------------------
# PROFESSIONAL COLOR PALETTE
# ----------------------------------------------------------
COLOR_PRIMARY = "#1f77b4"
COLOR_SECONDARY = "#2ca02c"
COLOR_ALERT = "#d62728"
COLOR_WARNING = "#ff7f0e"
COLOR_PURPLE = "#9467bd"

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    # ======================================================
    # LOAD DATA
    # ======================================================
    df = pd.read_excel(uploaded_file)

    # ---------------- CLEAN DATES ----------------
    df['Po_Date'] = pd.to_datetime(df['Po_Date'], errors='coerce', dayfirst=True)
    df['Scheduled_Date'] = pd.to_datetime(df['Scheduled_Date'], errors='coerce', dayfirst=True)
    df['Invoice_Dates'] = pd.to_datetime(df['Invoice_Dates'], errors='coerce', dayfirst=True)

    # ---------------- CLEAN NUMERIC ----------------
    df['PO_Value'] = pd.to_numeric(df['PO_Value'], errors='coerce')
    df['Supplied_Value'] = pd.to_numeric(df['Supplied_Value'], errors='coerce')

    df = df.dropna(subset=['Po_Date'])

    # ======================================================
    # SO LEVEL DATA (NO DUPLICATION)
    # ======================================================
    so_df = df.copy()
    so_df['Order_Month'] = so_df['Po_Date'].dt.to_period('M').astype(str)
    so_df['Order_Year'] = so_df['Po_Date'].dt.year
    so_df['Order_Quarter'] = so_df['Po_Date'].dt.to_period('Q').astype(str)

    # ======================================================
    # EXPAND PRODUCT LEVEL
    # ======================================================
    expanded_rows = []

    for _, row in df.iterrows():
        item_details = str(row['Item_Qty_Details']).split('\n')
        for item in item_details:
            parts = item.split('|')
            if len(parts) >= 4:
                product_name = parts[0].strip()
                po_qty = re.search(r'PO Qty:\s*([\d\.]+)', item)
                supplied_qty = re.search(r'Supplied Qty:\s*([\d\.]+)', item)
                po_qty = float(po_qty.group(1)) if po_qty else 0
                supplied_qty = float(supplied_qty.group(1)) if supplied_qty else 0

                new_row = row.to_dict()
                new_row['Product_Name'] = product_name
                new_row['PO_Qty'] = po_qty
                new_row['Supplied_Qty'] = supplied_qty
                expanded_rows.append(new_row)

    clean_df = pd.DataFrame(expanded_rows)

    # ------------------------------------------------------
    # PRODUCT LEVEL CALCULATIONS
    # ------------------------------------------------------
    clean_df['Lead_Time'] = (clean_df['Invoice_Dates'] - clean_df['Po_Date']).dt.days
    clean_df['Schedule_Delay'] = (clean_df['Invoice_Dates'] - clean_df['Scheduled_Date']).dt.days

    clean_df['Delivery_Status'] = np.where(
        clean_df['Invoice_Dates'] <= clean_df['Scheduled_Date'],
        'On-Time', 'Delayed'
    )

    clean_df['Order_Status'] = clean_df['Invoice_Dates'].apply(
        lambda x: 'Closed' if pd.notnull(x) else 'Open'
    )

    # Revenue Allocation FIX
    clean_df['Allocated_Value'] = np.where(
        clean_df['PO_Qty'] > 0,
        (clean_df['PO_Qty'] /
         clean_df.groupby('So_No')['PO_Qty'].transform('sum')) *
        clean_df['PO_Value'],
        0
    )

    # ======================================================
    # KPI SECTION
    # ======================================================
    total_sales = round(so_df['PO_Value'].sum(), 2)
    total_orders = so_df['So_No'].nunique()
    avg_lead_time = round(clean_df['Lead_Time'].mean(), 1)
    on_time_perc = round(
        clean_df['Delivery_Status'].value_counts(normalize=True)
        .get('On-Time', 0) * 100, 1)
    pending_value = so_df[so_df['Invoice_Dates'].isna()]['PO_Value'].sum()

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Sales", f"{total_sales:,.0f}")
    col2.metric("Total Orders", total_orders)
    col3.metric("Avg Lead Time", avg_lead_time)
    col4.metric("On-Time %", f"{on_time_perc}%")
    col5.metric("Pending Value", f"{pending_value:,.0f}")

    st.divider()

    # ======================================================
    # MONTHLY SALES
    # ======================================================
    monthly_sales = so_df.groupby('Order_Month')['PO_Value'].sum().reset_index()

    st.plotly_chart(
        px.line(monthly_sales, x='Order_Month', y='PO_Value',
                title="Monthly Sales Trend",
                markers=True,
                color_discrete_sequence=[COLOR_PRIMARY]),
        use_container_width=True)

    # Sales Growth
    monthly_sales['Sales_Growth_MoM'] = monthly_sales['PO_Value'].pct_change() * 100

    st.plotly_chart(
        px.line(monthly_sales, x='Order_Month',
                y='Sales_Growth_MoM',
                title="Sales Growth % (MoM)",
                markers=True,
                color_discrete_sequence=[COLOR_PURPLE]),
        use_container_width=True)

    # ======================================================
    # QUARTERLY & YEARLY
    # ======================================================
    quarterly_sales = so_df.groupby('Order_Quarter')['PO_Value'].sum().reset_index()
    yearly_sales = so_df.groupby('Order_Year')['PO_Value'].sum().reset_index()

    st.plotly_chart(px.bar(quarterly_sales,
                           x='Order_Quarter',
                           y='PO_Value',
                           title="Quarterly Sales",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    st.plotly_chart(px.line(yearly_sales,
                            x='Order_Year',
                            y='PO_Value',
                            markers=True,
                            title="Year-wise Sales",
                            color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    # ======================================================
    # TOP CUSTOMERS & PRODUCTS
    # ======================================================
    top_customers = so_df.groupby('Customer_Name')['PO_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_customers,
                           x='Customer_Name',
                           y='PO_Value',
                           title="Top 10 Customers",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    top_products = clean_df.groupby('Product_Name')['Allocated_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_products,
                           x='Product_Name',
                           y='Allocated_Value',
                           title="Top 10 Products by Revenue",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    # ======================================================
    # RFM ANALYSIS
    # ======================================================
    st.subheader("📊 Customer RFM Analysis")

    rfm = clean_df.groupby('Customer_Name').agg({
        'Invoice_Dates': lambda x: (pd.Timestamp.today() - x.max()).days,
        'So_No': 'nunique',
        'Allocated_Value': 'sum'
    }).reset_index()

    rfm.columns = ['Customer_Name', 'Recency', 'Frequency', 'Monetary']

    st.plotly_chart(
        px.scatter(rfm, x='Recency', y='Monetary',
                   size='Frequency', color='Monetary',
                   title="RFM Distribution"),
        use_container_width=True)

    # ======================================================
    # DELIVERY PERFORMANCE
    # ======================================================
    st.subheader("🚨 Late Delivery % by Customer")

    delay_df = clean_df.groupby(['Customer_Name','Delivery_Status'])['So_No'].count().unstack().fillna(0)

    if 'Delayed' in delay_df.columns:
        delay_df['Delay_%'] = (delay_df['Delayed'] /
                               delay_df.sum(axis=1)) * 100
        delay_df = delay_df.sort_values('Delay_%', ascending=False).head(10)

        st.plotly_chart(
            px.bar(delay_df.reset_index(),
                   x='Customer_Name',
                   y='Delay_%',
                   title="Top Customers by Delay %",
                   text_auto=True,
                   color_discrete_sequence=[COLOR_ALERT]),
            use_container_width=True)

    # ======================================================
    # PARETO ANALYSIS
    # ======================================================
    st.subheader("📌 Revenue Pareto (80/20)")

    customer_sales = so_df.groupby('Customer_Name')['PO_Value'].sum().sort_values(ascending=False).reset_index()
    customer_sales['Cumulative_%'] = customer_sales['PO_Value'].cumsum() / customer_sales['PO_Value'].sum() * 100

    st.plotly_chart(
        px.line(customer_sales,
                x=customer_sales.index,
                y='Cumulative_%',
                markers=True,
                title="Cumulative Revenue Contribution",
                color_discrete_sequence=[COLOR_PURPLE]),
        use_container_width=True)

    # ======================================================
    # AGING & LEAD TIME
    # ======================================================
    st.subheader("⏳ Aging of Open Orders")

    open_orders = clean_df[clean_df['Order_Status'] == 'Open'].copy()
    if not open_orders.empty:
        open_orders['Aging_Days'] = (pd.Timestamp.today() - open_orders['Po_Date']).dt.days
        st.plotly_chart(px.histogram(open_orders,
                                     x='Aging_Days',
                                     nbins=20,
                                     title="Open Orders Aging",
                                     color_discrete_sequence=[COLOR_WARNING]),
                        use_container_width=True)

    st.subheader("🚚 Lead Time Distribution")

    closed_orders = clean_df[clean_df['Order_Status'] == 'Closed'].copy()
    closed_orders['Lead_Time'] = (closed_orders['Invoice_Dates'] - closed_orders['Po_Date']).dt.days

    st.plotly_chart(px.histogram(closed_orders,
                                 x='Lead_Time',
                                 nbins=20,
                                 title="Lead Time Distribution",
                                 color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

else:
    st.info("Please upload your Excel file to start analysis.")
