import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# Set the title of the Streamlit app
st.title('Product Sales and Profit Analysis')

# File path (assuming the file is accessible by the Streamlit app)
file_path = "/content/drive/MyDrive/Herramientas Datos/Orders Final Limpio.xlsx"

# Read the data
@st.cache_data
def load_data(path):
    data = pd.read_excel(path)
    return data

df = load_data(file_path)

# Display data types
st.subheader('Data Types')
st.write(df.info())

# Calculate total sales per product
product_sales = df.groupby('Product Name')['Sales'].sum().reset_index()

# Get the top 5 best-selling products
top_5_products = product_sales.nlargest(5, 'Sales')

# Create a bar chart for top 5 best-selling products using Plotly Express
st.subheader('Top 5 Best-Selling Products')
fig_sales = px.bar(top_5_products, x='Sales', y='Product Name', title='Top 5 Best-Selling Products')
fig_sales.update_layout(yaxis={'categoryorder':'total ascending', 'tickangle': -45, 'tickfont': dict(size=10)}, margin=dict(l=150, r=20, t=40, b=30)) # Added layout updates
st.plotly_chart(fig_sales)

# Calculate total profit per product
product_profit = df.groupby('Product Name')['Profit'].sum().reset_index()

# Get the top 5 products by profit
top_5_profit_products = product_profit.nlargest(5, 'Profit')

# Create a bar chart for top 5 products by profit using Plotly Express
st.subheader('Top 5 Products by Profit')
fig_profit = px.bar(top_5_profit_products, x='Profit', y='Product Name', title='Top 5 Products by Profit')
fig_profit.update_layout(yaxis={'categoryorder':'total ascending', 'tickangle': -45, 'tickfont': dict(size=10)}, margin=dict(l=150, r=20, t=40, b=30)) # Added layout updates
st.plotly_chart(fig_profit)

# Display sales and profit details for the top 5 products
st.subheader('Details for Top 5 Best-Selling Products')
top_5_products_details = df[df['Product Name'].isin(top_5_products['Product Name'])]
st.write(top_5_products_details[['Product Name', 'Sales', 'Profit', 'Quantity']])
