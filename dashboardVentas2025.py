import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Product Sales and Profit Analysis", layout="wide")
st.title("Product Sales and Profit Analysis")

FILE_PATH = "Orders Final Limpio.xlsx"

# ---------------- Carga, limpieza y normalización ----------------
@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Normaliza encabezados
    df.columns = df.columns.str.strip()

    # --- Consolidar y eliminar 'Ship date' (d minúscula) preservando datos en 'Ship Date' ---
    if "Ship date" in df.columns:
        if "Ship Date" not in df.columns:
            # si solo existe la minúscula, renómbrala a la estándar
            df = df.rename(columns={"Ship date": "Ship Date"})
        else:
            # si existen ambas, completa 'Ship Date' con los valores válidos de 'Ship date'
            tmp = pd.to_datetime(df["Ship date"], errors="coerce")
            df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
            fill_mask = df["Ship Date"].isna() & tmp.notna()
            if fill_mask.any():
                df.loc[fill_mask, "Ship Date"] = tmp.loc[fill_mask]
            # elimina la columna con d minúscula
            df = df.drop(columns=["Ship date"])

    # Convierte fechas (acepta serial de Excel o texto)
    for col in ["Order Date", "Ship Date"]:
        if col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                # serial de Excel -> fecha
                df[col] = pd.to_datetime(df[col], unit="d", origin="1899-12-30", errors="coerce")
            else:
                df[col] = pd.to_datetime(df[col], errors="coerce")

    # Convierte numéricos frecuentes
    for col in ["Sales", "Profit", "Quantity", "Discount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

df = load_data(FILE_PATH)

# ---------------- Filtros ----------------
st.sidebar.subheader("Filtro por Región y Estado")
region_list = ["Todas"] + sorted(df["Region"].dropna().unique().tolist())
selected_region = st.sidebar.selectbox("Selecciona una Región", region_list)

df_region = df if selected_region == "Todas" else df[df["Region"] == selected_region].copy()

state_list = ["Todos"] + sorted(df_region["State"].dropna().unique().tolist())
selected_state = st.sidebar.selectbox("Selecciona un Estado", state_list)

filtered_df = df_region if selected_state == "Todos" else df_region[df_region["State"] == selected_state].copy()

# ---------------- Data Types (sin df.info() -> None) ----------------
st.subheader("Data Types")
st.dataframe(
    pd.DataFrame({"Column": df.columns, "Dtype": df.dtypes.astype(str)}),
    use_container_width=True,
)

# ---------------- Mostrar datos filtrados ----------------
if st.sidebar.checkbox("Mostrar datos filtrados"):
    st.subheader("Datos Filtrados")
    df_show = filtered_df.copy()

    # Garantiza dtype datetime (evita 'None' por cadenas)
    for col in ["Order Date", "Ship Date"]:
        if col in df_show.columns and not pd.api.types.is_datetime64_any_dtype(df_show[col]):
            df_show[col] = pd.to_datetime(df_show[col], errors="coerce")

    st.dataframe(
        df_show,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Order Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "Ship Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
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

# ---------------- Gráficas (horizontales, sin etiquetas diagonales) ----------------
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

# ---------------- Detalles (incluye Order/Ship Date bien formateadas) ----------------
st.subheader(f"Detalles de los 5 Productos Más Vendidos en {lugar}")
date_cols = [c for c in ["Order Date", "Ship Date"] if c in filtered_df.columns]
base_cols = ["Product Name", "Sales", "Profit", "Quantity", "Region", "State"]
cols_to_show = ["Product Name"] + date_cols + base_cols[1:]

details = filtered_df[filtered_df["Product Name"].isin(product_sales["Product Name"])]
details_show = details[cols_to_show].copy()

# Asegura dtype datetime para DateColumn
for col in date_cols:
    if not pd.api.types.is_datetime64_any_dtype(details_show[col]):
        details_show[col] = pd.to_datetime(details_show[col], errors="coerce")

st.dataframe(
    details_show.sort_values(["Product Name"] + date_cols),
    use_container_width=True,
    hide_index=True,
    column_config={
        "Order Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Ship Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Sales": st.column_config.NumberColumn(format="%.3f"),
        "Profit": st.column_config.NumberColumn(format="%.3f"),
    },
)
