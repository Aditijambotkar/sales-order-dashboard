# ==========================================================
# SALES ORDER COMPLETE ANALYTICS DASHBOARD
# (Full conversion of your Colab code - nothing removed)
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

    # ======================================================
    # CLEAN DATE COLUMNS
    # ======================================================
    df['Po_Date'] = pd.to_datetime(df['Po_Date'], errors='coerce', dayfirst=True)
    df['Scheduled_Date'] = pd.to_datetime(df['Scheduled_Date'], errors='coerce', dayfirst=True)
    df['Invoice_Dates'] = pd.to_datetime(df['Invoice_Dates'], errors='coerce', dayfirst=True)

    # ======================================================
    # CLEAN NUMERIC COLUMNS
    # ======================================================
    df['PO_Value'] = pd.to_numeric(df['PO_Value'], errors='coerce')
    df['Supplied_Value'] = pd.to_numeric(df['Supplied_Value'], errors='coerce')

    # ======================================================
    # EXPAND ITEM_QTY_DETAILS
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

    # ======================================================
    # ENSURE DATETIME
    # ======================================================
    date_cols = ['Po_Date', 'Invoice_Dates', 'Scheduled_Date']
    for col in date_cols:
        if col in clean_df.columns:
            clean_df[col] = pd.to_datetime(clean_df[col], errors='coerce')

    # ======================================================
    # DROP MISSING CRITICAL DATES
    # ======================================================
    clean_df = clean_df.dropna(subset=['Po_Date', 'Scheduled_Date'])

    # ======================================================
    # CALCULATED FIELDS
    # ======================================================
    clean_df['Lead_Time'] = (clean_df['Invoice_Dates'] - clean_df['Po_Date']).dt.days
    clean_df['Schedule_Delay'] = (clean_df['Invoice_Dates'] - clean_df['Scheduled_Date']).dt.days

    clean_df['Delivery_Status'] = np.where(
        clean_df['Invoice_Dates'] <= clean_df['Scheduled_Date'],
        'On-Time', 'Delayed'
    )

    clean_df['Order_Month'] = clean_df['Po_Date'].dt.to_period('M').astype(str)
    clean_df['Order_Year'] = clean_df['Po_Date'].dt.year
    clean_df['Order_Quarter'] = clean_df['Po_Date'].dt.to_period('Q').astype(str)

    # Customer last invoice
    customer_last_invoice = clean_df.groupby('Customer_Name')['Invoice_Dates'].max().reset_index()
    customer_last_invoice.rename(columns={'Invoice_Dates': 'Last_Invoice_Date'}, inplace=True)
    clean_df = clean_df.merge(customer_last_invoice, on='Customer_Name', how='left')
    clean_df['Last_Invoice_Date'] = pd.to_datetime(clean_df['Last_Invoice_Date'], errors='coerce')
    clean_df['Dormancy_Days'] = (pd.Timestamp.today() - clean_df['Last_Invoice_Date']).dt.days

    clean_df['Avg_Order_Value'] = np.where(
        clean_df['PO_Qty'] > 0,
        clean_df['PO_Value'] / clean_df['PO_Qty'],
        0
    )

    clean_df['Order_Status'] = clean_df['Invoice_Dates'].apply(
        lambda x: 'Closed' if pd.notnull(x) else 'Open'
    )

    # ======================================================
    # KPI SECTION
    # ======================================================
    total_sales = round(df['PO_Value'].sum(), 2)   # FIXED (no duplication)
    total_orders = df['So_No'].nunique()
    avg_lead_time = round(clean_df['Lead_Time'].mean(), 1)
    on_time_perc = round(clean_df['Delivery_Status'].value_counts(normalize=True).get('On-Time', 0) * 100, 1)

    top_customers_20perc = clean_df.groupby('Customer_Name')['PO_Value'].sum().sort_values(ascending=False)
    top_20perc_customers_value = round(top_customers_20perc[:int(len(top_customers_20perc) * 0.2)].sum(), 2)
    total_value = clean_df['PO_Value'].sum()

    pending_value = df[df['Invoice_Dates'].isna()]['PO_Value'].sum()  # FIXED

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Sales", f"{total_sales:,.0f}")
    col2.metric("Total Orders", total_orders)
    col3.metric("Avg Lead Time", avg_lead_time)
    col4.metric("On-Time %", f"{on_time_perc}%")
    col5.metric("Pending Value", f"{pending_value:,.0f}")

    st.write("Top 20% Customer Contribution:",
             round(top_20perc_customers_value / total_value * 100, 1), "%")

    st.divider()

    # ======================================================
    # ALL ORIGINAL CHARTS
    # ======================================================

    monthly_sales = clean_df.groupby('Order_Month')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.line(monthly_sales, x='Order_Month', y='PO_Value',
                            title="Monthly Sales Value Trend", markers=True,
                            color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    quarterly_sales = clean_df.groupby('Order_Quarter')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.bar(quarterly_sales, x='Order_Quarter', y='PO_Value',
                           title="Quarterly Sales Value", text_auto=True,
                           color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    monthly_so_count = clean_df.groupby('Order_Month')['So_No'].nunique().reset_index()
    st.plotly_chart(px.line(monthly_so_count, x='Order_Month', y='So_No',
                            title="Monthly SO Count Trend", markers=True,
                            color_discrete_sequence=[COLOR_PURPLE]),
                    use_container_width=True)

    monthly_aov = clean_df.groupby('Order_Month')['Avg_Order_Value'].mean().reset_index()
    st.plotly_chart(px.line(monthly_aov, x='Order_Month', y='Avg_Order_Value',
                            title="Average Order Value Trend", markers=True,
                            color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    top_customers = clean_df.groupby('Customer_Name')['PO_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_customers, x='Customer_Name', y='PO_Value',
                           title="Top 10 Customers by Value", text_auto=True,
                           color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    top_customers_qty = clean_df.groupby('Customer_Name')['PO_Qty'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_customers_qty, x='Customer_Name', y='PO_Qty',
                           title="Top 10 Customers by Quantity", text_auto=True,
                           color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    top_products = clean_df.groupby('Product_Name')['PO_Qty'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_products, x='Product_Name', y='PO_Qty',
                           title="Top 10 Products by Quantity", text_auto=True,
                           color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    top_products_val = clean_df.groupby('Product_Name')['PO_Value'].sum().nlargest(10).reset_index()
    st.plotly_chart(px.bar(top_products_val, x='Product_Name', y='PO_Value',
                           title="Top 10 Products by Value", text_auto=True,
                           color_discrete_sequence=[COLOR_PURPLE]),
                    use_container_width=True)

    customer_contrib = clean_df.groupby('Customer_Name')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.pie(customer_contrib, values='PO_Value',
                           names='Customer_Name',
                           title="Customer Contribution to Revenue"),
                    use_container_width=True)

    region_contrib = clean_df.groupby('Site_Address')['PO_Value'].sum().reset_index()
    st.plotly_chart(px.pie(region_contrib, values='PO_Value',
                           names='Site_Address',
                           title="Region-wise Contribution to Revenue"),
                    use_container_width=True)

    monthly_sales['Sales_Growth_MoM'] = monthly_sales['PO_Value'].pct_change() * 100
    st.plotly_chart(px.line(monthly_sales, x='Order_Month',
                            y='Sales_Growth_MoM',
                            title="Sales Growth % (MoM)", markers=True,
                            color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    # Lead Time Trend
    lead_time_month = clean_df.groupby('Order_Month')['Lead_Time'].mean().reset_index()
    st.plotly_chart(px.line(lead_time_month,
                            x='Order_Month',
                            y='Lead_Time',
                            markers=True,
                            title="Average Lead Time by Month",
                            color_discrete_sequence=[COLOR_PRIMARY]),
                    use_container_width=True)

    st.plotly_chart(px.pie(clean_df, names='Delivery_Status',
                           title="On-Time vs Delayed Orders",
                           color_discrete_map={
                               'On-Time': COLOR_SECONDARY,
                               'Delayed': COLOR_ALERT
                           }),
                    use_container_width=True)

    # Repeat vs Occasional
    customer_order_count = clean_df.groupby('Customer_Name')['So_No'].nunique().reset_index()
    customer_order_count['Customer_Type_Freq'] = np.where(customer_order_count['So_No'] > 1, 'Repeat', 'Occasional')
    repeat_counts = customer_order_count['Customer_Type_Freq'].value_counts().reset_index()
    repeat_counts.columns = ['Customer_Type', 'Count']
    st.plotly_chart(px.bar(repeat_counts, x='Customer_Type', y='Count',
                           title="Repeat vs Occasional Customers", text_auto=True,
                           color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    dormant_customers = clean_df[clean_df['Dormancy_Days'] > 90]['Customer_Name'].value_counts().reset_index()
    dormant_customers.columns = ['Customer_Name', 'Count']
    st.plotly_chart(px.bar(dormant_customers, x='Customer_Name', y='Count',
                           title="Dormant Customers (>90 days)", text_auto=True,
                           color_discrete_sequence=[COLOR_ALERT]),
                    use_container_width=True)

    # ======================================================
    # ADVANCED ORDER ANALYTICS
    # ======================================================

    st.subheader("📦 Open vs Closed Orders")
    status_counts = clean_df.groupby('Order_Status')['So_No'].nunique().reset_index()
    st.plotly_chart(px.bar(status_counts, x='Order_Status', y='So_No',
                           text_auto=True,
                           title="Open vs Closed Sales Orders",
                           color_discrete_sequence=[COLOR_WARNING]),
                    use_container_width=True)

    st.subheader("⏳ Aging of Open Orders")
    open_orders = clean_df[clean_df['Order_Status'] == 'Open'].copy()
    if not open_orders.empty:
        open_orders['Aging_Days'] = (pd.Timestamp.today() - open_orders['Po_Date']).dt.days
        open_orders = open_orders[open_orders['Aging_Days'] >= 0]
        st.plotly_chart(px.histogram(open_orders,
                                     x='Aging_Days',
                                     nbins=20,
                                     title="Aging Distribution (Open Orders)",
                                     color_discrete_sequence=[COLOR_PRIMARY]),
                        use_container_width=True)

    st.subheader("🚚 Order-to-Delivery Lead Time (Closed Orders)")
    closed_orders = clean_df[clean_df['Order_Status'] == 'Closed'].copy()
    closed_orders['Order_to_Delivery_LT'] = (closed_orders['Invoice_Dates'] - closed_orders['Po_Date']).dt.days
    closed_orders = closed_orders[closed_orders['Order_to_Delivery_LT'] >= 0]
    st.plotly_chart(px.histogram(closed_orders,
                                 x='Order_to_Delivery_LT',
                                 nbins=20,
                                 title="Order-to-Delivery Lead Time",
                                 color_discrete_sequence=[COLOR_SECONDARY]),
                    use_container_width=True)

    st.subheader("📅 Delivery Delay Analysis")
    if 'Scheduled_Date' in clean_df.columns:
        closed_orders['Delivery_Delay'] = (closed_orders['Invoice_Dates'] - closed_orders['Scheduled_Date']).dt.days
        closed_orders['Delay_Type'] = closed_orders['Delivery_Delay'].apply(lambda x: 'Late' if x > 0 else 'Early/On-Time')
        delay_summary = closed_orders.groupby('Delay_Type')['So_No'].nunique().reset_index()

        st.plotly_chart(px.bar(delay_summary,
                               x='Delay_Type',
                               y='So_No',
                               text_auto=True,
                               title="Early vs Late Deliveries",
                               color_discrete_sequence=[COLOR_ALERT]),
                        use_container_width=True)

        st.plotly_chart(px.histogram(closed_orders,
                                     x='Delivery_Delay',
                                     nbins=20,
                                     title="Delivery Delay Distribution",
                                     color_discrete_sequence=[COLOR_WARNING]),
                        use_container_width=True)
            # ======================================================
    # 7️⃣ CAPACITY & PLANNING SUPPORT (OPERATIONS VIEW)
    # ======================================================

    st.subheader("🏭 Capacity & Planning Support (Operations View)")

    # ------------------------------------------------------
    # 1️⃣ Expected Load by Month (Total Quantity)
    # ------------------------------------------------------
    monthly_load = clean_df.groupby('Order_Month')['PO_Qty'].sum().reset_index()

    st.plotly_chart(
        px.bar(
            monthly_load,
            x='Order_Month',
            y='PO_Qty',
            title="Expected Production Load by Month (Total Ordered Quantity)",
            text_auto=True,
            color_discrete_sequence=["#1f77b4"]
        ),
        use_container_width=True
    )

    # ------------------------------------------------------
    # 2️⃣ Peak Demand Periods
    # ------------------------------------------------------
    peak_months = monthly_load.sort_values(by='PO_Qty', ascending=False).head(3)

    st.write("### 🔥 Top 3 Peak Demand Months")
    st.dataframe(peak_months, use_container_width=True)

    # ------------------------------------------------------
    # 3️⃣ Product-wise Capacity Requirement
    # ------------------------------------------------------
    product_capacity = clean_df.groupby('Product_Name')['PO_Qty'].sum().reset_index()
    product_capacity = product_capacity.sort_values(by='PO_Qty', ascending=False)

    st.plotly_chart(
        px.bar(
            product_capacity.head(15),
            x='Product_Name',
            y='PO_Qty',
            title="Product-wise Capacity Requirement (Top 15)",
            text_auto=True,
            color_discrete_sequence=["#2ca02c"]
        ),
        use_container_width=True
    )

    # ===============================
# PRODUCT DEMAND ANALYTICS SECTION
# ===============================

st.header("Product Demand & Seasonality Analysis")

# --------------------------------
# 1️⃣ Product-wise Demand Trend
# --------------------------------
st.subheader("Product-wise Demand Trend")

product_trend = clean_df.groupby(
    ['Order_Month', 'Product_Name']
)['Ordered_Qty'].sum().reset_index()

fig_product_trend = px.line(
    product_trend,
    x='Order_Month',
    y='Ordered_Qty',
    color='Product_Name',
    markers=True,
    title="Monthly Demand Trend by Product"
)

st.plotly_chart(fig_product_trend, use_container_width=True)


# --------------------------------
# 2️⃣ Seasonality Heatmap
# --------------------------------
st.subheader("Seasonality Analysis (Heatmap)")

seasonality = clean_df.groupby(
    ['Order_Month', 'Product_Name']
)['Ordered_Qty'].sum().reset_index()

season_pivot = seasonality.pivot(
    index='Product_Name',
    columns='Order_Month',
    values='Ordered_Qty'
).fillna(0)

fig_season = px.imshow(
    season_pivot,
    aspect="auto",
    title="Product Seasonality Heatmap"
)

st.plotly_chart(fig_season, use_container_width=True)


# --------------------------------
# 3️⃣ Slow-moving vs Fast-moving
# --------------------------------
st.subheader("Slow-moving vs Fast-moving Products")

product_speed = clean_df.groupby(
    'Product_Name'
)['Ordered_Qty'].sum().reset_index()

avg_demand = product_speed['Ordered_Qty'].mean()

product_speed['Category'] = product_speed['Ordered_Qty'].apply(
    lambda x: 'Fast Moving' if x > avg_demand else 'Slow Moving'
)

fig_speed = px.bar(
    product_speed.sort_values('Ordered_Qty', ascending=False),
    x='Product_Name',
    y='Ordered_Qty',
    color='Category',
    title="Fast vs Slow Moving Products"
)

st.plotly_chart(fig_speed, use_container_width=True)

# ===============================





