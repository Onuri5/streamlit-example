import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# Set the title of the Streamlit app
st.title('Product Sales and Profit Analysis')

# File path (assuming the file is accessible by the Streamlit app)
file_path = "Orders Final Limpio.xlsx"

# Read the data
@st.cache_data
def load_data(path):
    data = pd.read_excel(path)
    return data

df = load_data(file_path)

# Display data types
st.subheader('Data Types')
st.write(df.info())

# Add a region filter
st.sidebar.subheader("Filtro por Región")
region_list = df['Region'].unique().tolist()
region_list.sort()
region_list.insert(0, 'Todas') # Add "Todas" option
selected_region = st.sidebar.selectbox('Selecciona una Región', region_list)

# Filter data by selected region
if selected_region != 'Todas':
    filtered_df = df[df['Region'] == selected_region].copy()
else:
    filtered_df = df.copy()

# Add a state filter based on the selected region
state_list = filtered_df['State'].unique().tolist()
state_list.sort()
state_list.insert(0, 'Todos') # Add "Todos" option
selected_state = st.sidebar.selectbox('Selecciona un Estado', state_list)

# Filter data by selected state
if selected_state != 'Todos':
    filtered_df = filtered_df[filtered_df['State'] == selected_state].copy()
else:
    filtered_df = filtered_df.copy()

# Calculate total sales per product for the filtered data
product_sales = filtered_df.groupby('Product Name')['Sales'].sum().reset_index()

# Get the top 5 best-selling products for the filtered data
top_5_products = product_sales.nlargest(5, 'Sales')

# Create a bar chart for top 5 best-selling products using Plotly Express
st.subheader(f'Top 5 Productos Más Vendidos en {selected_state if selected_state != "Todos" else selected_region}')
fig_sales = px.bar(top_5_products, x='Sales', y='Product Name', title=f'Top 5 Productos Más Vendidos en {selected_state if selected_state != "Todos" else selected_region}')
fig_sales.update_layout(yaxis={'categoryorder':'total ascending', 'tickangle': -45, 'tickfont': dict(size=10)}, margin=dict(l=150, r=20, t=40, b=30))
st.plotly_chart(fig_sales)

# Calculate total profit per product for the filtered data
product_profit = filtered_df.groupby('Product Name')['Profit'].sum().reset_index()

# Get the top 5 products by profit for the filtered data
top_5_profit_products = product_profit.nlargest(5, 'Profit')

# Create a bar chart for top 5 products by profit using Plotly Express
st.subheader(f'Top 5 Productos por Ganancia en {selected_state if selected_state != "Todos" else selected_region}')
fig_profit = px.bar(top_5_profit_products, x='Profit', y='Product Name', title=f'Top 5 Productos por Ganancia en {selected_state if selected_state != "Todos" else selected_region}')
fig_profit.update_layout(yaxis={'categoryorder':'total ascending', 'tickangle': -45, 'tickfont': dict(size=10)}, margin=dict(l=150, r=20, t=40, b=30))
st.plotly_chart(fig_profit)

# Display sales and profit details for the top 5 products in the filtered data
st.subheader(f'Detalles de los 5 Productos Más Vendidos en {selected_state if selected_state != "Todos" else selected_region}')
top_5_products_details = filtered_df[filtered_df['Product Name'].isin(top_5_products['Product Name'])].copy()
st.write(top_5_products_details[['Product Name', 'Sales', 'Profit', 'Quantity', 'Region', 'State']])
