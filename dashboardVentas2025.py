import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Product Sales and Profit Analysis", layout="wide")
st.title("Product Sales and Profit Analysis")

FILE_PATH = "Orders Final Limpio.xlsx"

@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Normaliza encabezados (espacios/casos)
    df.columns = df.columns.str.strip()

    # Unifica nombres típicos de fechas
    rename_map = {}
    for c in df.columns:
        lc = c.lower().strip()
        if lc == "order date":
            rename_map[c] = "Order Date"
        if lc in ("ship date", "shipdate", "ship date "):
            rename_map[c] = "Ship Date"
    if rename_map:
        df = df.rename(columns=rename_map)

    # Convierte fechas
    for col in ["Order Date", "Ship Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Convierte numéricos
    for col in ["Sales", "Profit", "Quantity"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

df = load_data(FILE_PATH)

# ---------------- Sidebar filtros ----------------
st.sidebar.subheader("Filtro por Región y Estado")
region_list = ["Todas"] + sorted(df["Region"].dropna().unique().tolist())
selected_region = st.sidebar.selectbox("Selecciona una Región", region_list)

df_region = df if selected_region == "Todas" else df[df["Region"] == selected_region].copy()

state_list = ["Todos"] + sorted(df_region["State"].dropna().unique().tolist())
selected_state = st.sidebar.selectbox("Selecciona un Estado", state_list)

filtered_df = df_region if selected_state == "Todos" else df_region[df_region["State"] == selected_state].copy()

if st.sidebar.checkbox("Mostrar datos filtrados"):
    st.subheader("Datos Filtrados")
    st.dataframe(filtered_df, use_container_width=True)

# ---------------- Tipos de datos (arregla el 'None') ----------------
st.subheader("Data Types")
st.dataframe(
    pd.DataFrame({"Column": df.columns, "Dtype": df.dtypes.astype(str)}),
    use_container_width=True,
)

# ---------------- Agregaciones ----------------
product_sales = (
    filtered_df.groupby("Product Name", as_index=False)["Sales"]
    .sum()
    .sort_values("Sales", ascending=False)
    .head(5)
)
product_profit = (
    filtered_df.groupby("Product Name", as_index=False)["Profit"]
    .sum()
    .sort_values("Profit", ascending=False)
    .head(5)
)

lugar = selected_state if selected_state != "Todos" else selected_region

# ---------------- Gráficas (sin etiquetas diagonales) ----------------
st.subheader(f"Top 5 Productos Más Vendidos en {lugar}")
fig_sales = px.bar(
    product_sales.sort_values("Sales", ascending=True),
    x="Sales",
    y="Product Name",
    orientation="h",
    labels={"Sales": "Ventas", "Product Name": "Producto"},
    title=f"Top 5 Productos Más Vendidos en {lugar}",
)
fig_sales.update_yaxes(automargin=True)
fig_sales.update_layout(margin=dict(l=280, r=20, t=40, b=40))
st.plotly_chart(fig_sales, use_container_width=True)

st.subheader(f"Top 5 Productos por Ganancia en {lugar}")
fig_profit = px.bar(
    product_profit.sort_values("Profit", ascending=True),
    x="Profit",
    y="Product Name",
    orientation="h",
    labels={"Profit": "Ganancia", "Product Name": "Producto"},
    title=f"Top 5 Productos por Ganancia en {lugar}",
)
fig_profit.update_yaxes(automargin=True)
fig_profit.update_layout(margin=dict(l=280, r=20, t=40, b=40))
st.plotly_chart(fig_profit, use_container_width=True)

# ---------------- Detalle (ahora incluye fechas) ----------------
st.subheader(f"Detalles de los 5 Productos Más Vendidos en {lugar}")
date_cols = [c for c in ["Order Date", "Ship Date"] if c in filtered_df.columns]
base_cols = ["Product Name", "Sales", "Profit", "Quantity", "Region", "State"]
cols_to_show = ["Product Name"] + date_cols + base_cols[1:]

details = filtered_df[filtered_df["Product Name"].isin(product_sales["Product Name"])]
st.dataframe(details[cols_to_show].sort_values(["Product Name", *date_cols]), use_container_width=True)
