# ==========================================================
# SALES ORDER COMPLETE ANALYTICS DASHBOARD
# (MASTER FULL VERSION – NOTHING REMOVED – ALL FIXED)
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
    df.columns = df.columns.str.strip()

    # ---------------- SAFE DATE CLEANING ----------------
    date_cols = ['Po_Date', 'Scheduled_Date', 'Invoice_Dates']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    # ---------------- SAFE NUMERIC CLEANING ----------------
    numeric_cols = ['PO_Value', 'Supplied_Value']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['Po_Date'])

    # ======================================================
    # SO LEVEL DATA (STRICTLY UNIQUE SO)
    # ======================================================
    so_df = df.drop_duplicates(subset=['So_No']).copy()

    so_df['Order_Month'] = so_df['Po_Date'].dt.to_period('M').astype(str)
    so_df['Order_Year'] = so_df['Po_Date'].dt.year
    so_df['Order_Quarter'] = so_df['Po_Date'].dt.to_period('Q').astype(str)

    # ======================================================
    # PRODUCT EXPANSION
    # ======================================================
    expanded_rows = []

    for _, row in df.iterrows():
        item_details = str(row.get('Item_Qty_Details', '')).split('\n')
        for item in item_details:
            product_name = item.split('|')[0].strip()

            po_qty_match = re.search(r'PO Qty:\s*([\d\.]+)', item)
            supplied_qty_match = re.search(r'Supplied Qty:\s*([\d\.]+)', item)

            po_qty = float(po_qty_match.group(1)) if po_qty_match else 0
            supplied_qty = float(supplied_qty_match.group(1)) if supplied_qty_match else 0

            new_row = row.to_dict()
            new_row['Product_Name'] = product_name
            new_row['PO_Qty'] = po_qty
            new_row['Supplied_Qty'] = supplied_qty

            expanded_rows.append(new_row)

    clean_df = pd.DataFrame(expanded_rows)

    # ======================================================
    # PRODUCT LEVEL CALCULATIONS
    # ======================================================
    # ------------------------------
# Safe Date Conversion
# ------------------------------

    clean_df['Po_Date'] = pd.to_datetime(clean_df['Po_Date'], errors='coerce')
    clean_df['Invoice_Dates'] = pd.to_datetime(clean_df['Invoice_Dates'], errors='coerce')

# ------------------------------
# Lead Time Calculation
# ------------------------------

    clean_df['Lead_Time'] = (
        clean_df['Invoice_Dates'] - clean_df['Po_Date']
          ).dt.days

# Remove invalid lead times
    clean_df = clean_df[clean_df['Lead_Time'].notna()]
    clean_df = clean_df[clean_df['Lead_Time'] >= 0]

    
    clean_df['Schedule_Delay'] = (clean_df['Invoice_Dates'] - clean_df['Scheduled_Date']).dt.days

    clean_df['Delivery_Status'] = np.where(
        clean_df['Invoice_Dates'] <= clean_df['Scheduled_Date'],
        'On-Time', 'Delayed'
    )

    clean_df['Order_Month'] = clean_df['Po_Date'].dt.to_period('M').astype(str)
    clean_df['Order_Year'] = clean_df['Po_Date'].dt.year
    clean_df['Order_Quarter'] = clean_df['Po_Date'].dt.to_period('Q').astype(str)

    clean_df['Order_Status'] = np.where(
        clean_df['Invoice_Dates'].notna(), 'Closed', 'Open'
    )

    # ======================================================
    # SAFE REVENUE ALLOCATION (NO DUPLICATION)
    # ======================================================

    clean_df['Total_PO_Qty_Per_SO'] = clean_df.groupby('So_No')['PO_Qty'].transform('sum')

    clean_df['Allocated_Value'] = np.where(
        clean_df['Total_PO_Qty_Per_SO'] > 0,
        (clean_df['PO_Qty'] / clean_df['Total_PO_Qty_Per_SO']) * clean_df['PO_Value'],
        0
    )

    # ======================================================
    # CUSTOMER DORMANCY
    # ======================================================

    customer_last_invoice = so_df.groupby('Customer_Name')['Invoice_Dates'].max().reset_index()
    customer_last_invoice.rename(columns={'Invoice_Dates': 'Last_Invoice_Date'}, inplace=True)

    so_df = so_df.merge(customer_last_invoice, on='Customer_Name', how='left')
    so_df['Dormancy_Days'] = (pd.Timestamp.today() - so_df['Last_Invoice_Date']).dt.days

    # ======================================================
    # KPI SECTION
    # ======================================================

    total_sales = so_df['PO_Value'].sum()
    total_orders = so_df['So_No'].nunique()
    avg_lead_time = clean_df['Lead_Time'].mean()
    pending_value = so_df[so_df['Invoice_Dates'].isna()]['PO_Value'].sum()

    on_time_perc = round(
        clean_df['Delivery_Status'].value_counts(normalize=True)
        .get('On-Time', 0) * 100, 1)

    top_customer_value = so_df.groupby('Customer_Name')['PO_Value'].sum().sort_values(ascending=False)
    top_20_value = top_customer_value.head(max(1, int(len(top_customer_value) * 0.2))).sum()

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Sales", f"{total_sales:,.0f}")
    col2.metric("Total Orders", total_orders)
    col3.metric("Avg Lead Time", f"{avg_lead_time:.1f}")
    col4.metric("On-Time %", f"{on_time_perc}%")
    col5.metric("Pending Value", f"{pending_value:,.0f}")

    st.write("Top 20% Customer Contribution:",
             round((top_20_value / total_sales) * 100, 1), "%")

    st.divider()

    # ======================================================
    # ALL ORIGINAL CHARTS (EVERYTHING INCLUDED)
    # ======================================================

    # Monthly Sales
    monthly_sales = so_df.groupby('Order_Month')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.line(monthly_sales, x='Order_Month', y='PO_Value',
                            title="Monthly Sales Value Trend", markers=True,
                            color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    # Monthly SO Count
    monthly_so_count = so_df.groupby('Order_Month')['So_No'].count().reset_index()
    st.plotly_chart(px.line(monthly_so_count, x='Order_Month', y='So_No',
                            title="Monthly SO Count Trend", markers=True,
                            color_discrete_sequence=[COLOR_PURPLE]),
                    use_container_width=True)

    # Average Order Value
    monthly_aov = so_df.groupby('Order_Month')['PO_Value'].mean().reset_index()
    st.plotly_chart(px.line(monthly_aov, x='Order_Month', y='PO_Value',
                            title="Average Order Value Trend", markers=True,
                            color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    # Sales Growth
    monthly_sales['Sales_Growth_MoM'] = monthly_sales['PO_Value'].pct_change() * 100
    st.plotly_chart(px.line(monthly_sales, x='Order_Month',
                            y='Sales_Growth_MoM',
                            title="Sales Growth % (MoM)", markers=True,
                            color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    # Top Customers by Value
    top_customers = so_df.groupby('Customer_Name')['PO_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_customers, x='Customer_Name', y='PO_Value',
                           title="Top 10 Customers by Value",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    # Top Customers by Quantity
    top_customers_qty = clean_df.groupby('Customer_Name')['PO_Qty'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_customers_qty, x='Customer_Name', y='PO_Qty',
                           title="Top 10 Customers by Quantity",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    # Top Products by Quantity
    top_products = clean_df.groupby('Product_Name')['PO_Qty'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_products, x='Product_Name', y='PO_Qty',
                           title="Top 10 Products by Quantity",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    # Top Products by Value
    top_products_val = clean_df.groupby('Product_Name')['Allocated_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_products_val, x='Product_Name', y='Allocated_Value',
                           title="Top 10 Products by Value",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_PURPLE]),
                    use_container_width=True)

    # Customer Contribution
    customer_contrib = so_df.groupby('Customer_Name')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.pie(customer_contrib, values='PO_Value',
                           names='Customer_Name',
                           title="Customer Contribution to Revenue"),
                    use_container_width=True)

    # Region Contribution
    region_contrib = so_df.groupby('Site_Address')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.pie(region_contrib, values='PO_Value',
                           names='Site_Address',
                           title="Region-wise Contribution to Revenue"),
                    use_container_width=True)

    # Open vs Closed
    status_counts = clean_df.groupby('Order_Status')['So_No'].nunique().reset_index()
    st.plotly_chart(px.bar(status_counts, x='Order_Status', y='So_No',
                           title="Open vs Closed Orders",
                           text_auto=True,
                           color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    # Aging
    open_orders = clean_df[clean_df['Order_Status'] == 'Open'].copy()
    if not open_orders.empty:
        open_orders['Aging_Days'] = (pd.Timestamp.today() - open_orders['Po_Date']).dt.days
        open_orders = open_orders[open_orders['Aging_Days'] >= 0]

        st.plotly_chart(px.histogram(open_orders, x='Aging_Days',
                                     nbins=20,
                                     title="Aging Distribution (Open Orders)",
                                     color_discrete_sequence=[COLOR_PRIMARY]),
                        use_container_width=True)

    # Order-to-Delivery
    closed_orders = clean_df[clean_df['Order_Status'] == 'Closed'].copy()
    closed_orders['Order_to_Delivery_LT'] = (
        closed_orders['Invoice_Dates'] - closed_orders['Po_Date']
    ).dt.days

    closed_orders = closed_orders[closed_orders['Order_to_Delivery_LT'] >= 0]

    st.plotly_chart(px.histogram(closed_orders,
                                 x='Order_to_Delivery_LT',
                                 nbins=20,
                                 title="Order-to-Delivery Lead Time",
                                 color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

else:
    st.info("Please upload your Excel file to start analysis.")

