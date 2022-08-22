# Importing Libraries
import pandas as pd
import streamlit as st
import plotly.express as px

# Importing Data
df = pd.read_excel(
        io=r'/Users/dario/Desktop/excel/Dynamic Dashboard - Coca Cola.xlsx',
        engine='openpyxl',
        sheet_name='raw_data',
        usecols='A:L',
        nrows=3889
    )
# Cleaning data and creating month column
df = df.rename(columns={'Beverage Brand': 'Beverage_Brand'})
df['Month'] = pd.DatetimeIndex(df['Invoice Date']).month

# Page layout
st.set_page_config(page_title='Coca-Cola Sales Report', page_icon='🥤', layout='wide')
st.title('🥤 Fizzy Drinks Sales', )

# Select boxes for interactivity
left_column_multiselect, middle_column_multiselect, middle_column_2_multiselect, right_column_multiselect = st.columns(4)
with left_column_multiselect:
    retailer = st.multiselect('Select retailer', options=df['Retailer'].unique(), default=df['Retailer'].unique())
with middle_column_multiselect:
    region = st.multiselect('Select region', options=df['Region'].unique(), default=df['Region'].unique())
with middle_column_2_multiselect:
    brand = st.multiselect('Select brand', options=df['Beverage_Brand'].unique(), default=df['Beverage_Brand'].unique())
with right_column_multiselect:
    month = st.multiselect('Select Month of Invoice Date', options=df['Month'].unique(), default=df['Month'].unique())

# Df resulting from selection in boxes
df_selection = df.query(
    'Retailer == @retailer & Region == @region & Beverage_Brand == @brand & Month == @month'
)

st.markdown('---')

# KPIs
left_column, middle_column, middle_column_2, right_column = st.columns(4)
with left_column:
    st.header('Total Sales:')
    st.subheader(f'${int(round(df_selection["Total Sales"].sum()))}')
with middle_column:
    st.header('Total Units Sold:')
    st.subheader(f'{int(df_selection["Units Sold"].sum())}')
with middle_column_2:
    st.header('Total Operating Profit:')
    st.subheader(f'${df_selection["Operating Profit"].sum()}')
with right_column:
    st.header('Average Operating Margin:')
    st.subheader(f'{(df_selection["Operating Margin"].mean() * 100)}%')

# Monthly Sales Chart
monthly_sales = df_selection.groupby(by='Month').sum()['Total Sales']
fig_monthly_sales = px.bar(monthly_sales, x=monthly_sales.index, y='Total Sales', title='<b>Interactive Monthly Sales</b>')
fig_monthly_sales.update_layout(width=2000)
st.plotly_chart(fig_monthly_sales)

st.markdown('---')

# Resulting df
st.dataframe(df_selection)