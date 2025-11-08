import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Product Sales and Profit Analysis", layout="wide")
st.title("Product Sales and Profit Analysis")

FILE_PATH = "Orders Final Limpio.xlsx"

# ---------------- Carga y limpieza ----------------
@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    # Normaliza encabezados
    df.columns = df.columns.str.strip()
    df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')
    df['Ship Date'] = pd.to_datetime(df['Ship Date'], errors='coerce')

    # Si existen ambas variantes, nos quedamos con "Ship Date"
    if "Ship Date" in df.columns and "Ship date" in df.columns:
        # Si la "Ship Date" está vacía en alguna fila pero "Ship date" tiene dato, complétalo.
        mask_fill = df["Ship Date"].isna() & df["Ship date"].notna()
        if mask_fill.any():
            df.loc[mask_fill, "Ship Date"] = df.loc[mask_fill, "Ship date"]
        df = df.drop(columns=["Ship date"])

    # Convierte fechas
    for col in ["Order Date", "Ship Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Convierte numéricos (por si vienen como texto)
    for col in ["Sales", "Profit", "Quantity", "Discount"]:
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

# ---------------- Data types (evita mostrar None) ----------------
st.subheader("Data Types")
st.dataframe(
    pd.DataFrame({"Column": df.columns, "Dtype": df.dtypes.astype(str)}),
    use_container_width=True,
)

# ---------------- Mostrar datos filtrados (con formato de fechas) ----------------
if st.sidebar.checkbox("Mostrar datos filtrados"):
    st.subheader("Datos filtrados")
    # Fuerza formato YYYY-MM-DD para evitar el texto "115 years"
    df_to_show = filtered_df.copy()
    for col in ["Order Date", "Ship Date"]:
    if col in df_to_show.columns:
        df_to_show[col] = df_to_show[col].apply(
            lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else ""
        )
    st.dataframe(
        df_to_show,
        use_container_width=True,
        column_config={
            "Sales": st.column_config.NumberColumn(format="%.3f"),
            "Profit": st.column_config.NumberColumn(format="%.3f"),
            "Discount": st.column_config.NumberColumn(format="%.2f"),
        },
    )

# ---------------- Agregaciones ----------------
if not filtered_df.empty:
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
else:
    product_sales = pd.DataFrame(columns=["Product Name", "Sales"])
    product_profit = pd.DataFrame(columns=["Product Name", "Profit"])

lugar = selected_state if selected_state != "Todos" else selected_region

# ---------------- Gráficas sin etiquetas diagonales ----------------
st.subheader(f"Top 5 Productos Más Vendidos en {lugar}")
fig_sales = px.bar(
    product_sales.sort_values("Sales", ascending=True),
    x="Sales",
    y="Product Name",
    orientation="h",
    labels={"Sales": "Ventas", "Product Name": "Producto"},
    title=f"Top 5 Productos Más Vendidos en {lugar}",
)
fig_sales.update_yaxes(automargin=True)  # deja margen para nombres largos
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

# ---------------- Detalle (incluye Order/Ship Date) ----------------
st.subheader(f"Detalles de los 5 Productos Más Vendidos en {lugar}")
date_cols = [c for c in ["Order Date", "Ship Date"] if c in filtered_df.columns]
base_cols = ["Product Name", "Sales", "Profit", "Quantity", "Region", "State"]
cols_to_show = ["Product Name"] + date_cols + base_cols[1:]

details = filtered_df[filtered_df["Product Name"].isin(product_sales["Product Name"])]
details_to_show = details[cols_to_show].copy()

# Formato visual de fechas en la tabla de detalle
for col in date_cols:
    details_to_show[col] = details_to_show[col].apply(
        lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else ""
    )

st.dataframe(details_to_show.sort_values(["Product Name"] + date_cols), use_container_width=True)
