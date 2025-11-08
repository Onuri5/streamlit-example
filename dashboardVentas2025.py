import streamlit as st
import pandas as pd
import plotly.express as px
import textwrap
import numpy as np

# -------------------------
# Configuración de página
# -------------------------
st.set_page_config(page_title="Análisis de Ventas y Ganancias", layout="wide")
st.title("Análisis de Ventas y Ganancias de Productos")

# -------------------------
# Carga y limpieza de datos
# -------------------------
file_path = "Orders Limpio Final.xlsx"

@st.cache_data
def load_data(path):
    df = pd.read_excel(path)

    # Mostrar columnas detectadas
    st.write("Columnas detectadas:", df.columns.tolist())

    # Normalizar nombres de columnas (quita espacios y pone mayúscula inicial)
    df.columns = df.columns.str.strip().str.title()

    # Verificar si existen las columnas clave
    if "Order Date" not in df.columns:
        st.error("❌ No se encontró la columna 'Order Date' en el archivo.")
        st.stop()

    # Unificar columnas de Ship Date si existen variantes
    if "Ship Date" in df.columns and "Ship Date" in df.columns:
        sd_main = pd.to_datetime(df["Ship Date"], errors="coerce")
        df["Ship Date"] = sd_main
    elif "Ship Date" not in df.columns and "Ship date" in df.columns:
        df["Ship Date"] = pd.to_datetime(df["Ship date"], errors="coerce")
        df = df.drop(columns=["Ship date"])

    # Convertir fechas de Order Date
    col_fecha = "Order Date"
    if pd.api.types.is_datetime64_any_dtype(df[col_fecha]):
        pass
    elif pd.api.types.is_numeric_dtype(df[col_fecha]):
        origin_date = pd.Timestamp("1899-12-30")
        df[col_fecha] = pd.to_timedelta(df[col_fecha], unit="D") + origin_date
    elif pd.api.types.is_timedelta64_dtype(df[col_fecha]):
        origin_date = pd.Timestamp("1899-12-30")
        df[col_fecha] = origin_date + df[col_fecha]
    else:
        df[col_fecha] = pd.to_datetime(df[col_fecha], format='mixed', errors='coerce')

    # Verificar si las fechas se convirtieron correctamente
    if df[col_fecha].isna().all():
        st.error("❌ No se pudieron convertir las fechas de 'Order Date'. Revisa el formato en el archivo Excel.")
        st.stop()

    # Verificación visual
    st.write("Vista previa de fechas convertidas:")
    st.dataframe(df[["Order Date", "Ship Date"]].head(10))

    return df

df_orders = load_data(file_path)

# -------------------------
# Filtros laterales
# -------------------------
with st.sidebar:
    st.header("Filtros")

    # Fechas mínimas y máximas
    min_date = df_orders["Order Date"].min().date()
    max_date = df_orders["Order Date"].max().date()

    # Filtro por rango de fechas
    rango = st.date_input(
        "Rango de fechas",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="YYYY/MM/DD",
        help="Selecciona un rango dentro de las fechas disponibles en los datos.",
    )

    # Validar rango seleccionado
    if isinstance(rango, tuple) and len(rango) == 2:
        start_date, end_date = rango
    else:
        start_date, end_date = min_date, max_date

    if start_date > end_date:
        start_date, end_date = end_date, start_date
        st.warning("⚠️ Se invirtió el rango de fechas (inicio > fin).")

    # Filtros por región y estado
    region = None
    estado = None
    if "Region" in df_orders.columns:
        region = st.selectbox("Selecciona Región", ["Todas"] + sorted(df_orders["Region"].dropna().unique().tolist()))
    if "State" in df_orders.columns:
        estado = st.selectbox("Selecciona Estado", ["Todas"] + sorted(df_orders["State"].dropna().unique().tolist()))

    mostrar_tabla = st.checkbox("Mostrar datos filtrados", value=True)

# -------------------------
# Aplicar filtros
# -------------------------
mask = (df_orders["Order Date"] >= pd.Timestamp(start_date)) & (df_orders["Order Date"] <= pd.Timestamp(end_date))
if region and region != "Todas":
    mask &= (df_orders["Region"] == region)
if estado and estado != "Todas":
    mask &= (df_orders["State"] == estado)

df_filtered = df_orders.loc[mask].copy()

if df_filtered.empty:
    st.warning("⚠️ No hay datos disponibles para el rango o filtros seleccionados.")
    st.stop()

st.success("✅ Datos cargados y filtrados correctamente.")

# -------------------------
# Mostrar datos filtrados
# -------------------------
st.subheader("Datos filtrados")

if mostrar_tabla:
    cols_pref = [
        "Order Date", "Ship Date", "Discount", "Sales", "Quantity", "Profit",
        "Region", "State", "Order Id", "Product Name", "City"
    ]
    cols_show = [c for c in cols_pref if c in df_filtered.columns]
    if not cols_show:
        cols_show = df_filtered.columns.tolist()

    st.dataframe(
        df_filtered[cols_show].sort_values("Order Date"),
        use_container_width=True,
        hide_index=True,
    )

# -------------------------
# Función para ajustar etiquetas largas
# -------------------------
def wrap_text(txt: str, width: int = 22) -> str:
    return "<br>".join(textwrap.wrap(str(txt), width=width))

# -------------------------
# Gráficos: Top 5 Ventas y Ganancias
# -------------------------
ventas_por_producto = df_filtered.groupby("Product Name")["Sales"].sum()
ganancias_por_producto = df_filtered.groupby("Product Name")["Profit"].sum()

# Top 5 por ventas
top_5_v = ventas_por_producto.sort_values(ascending=False).head(5)
fig_ventas = px.bar(
    x=top_5_v.index,
    y=top_5_v.values,
    labels={"x": "Producto", "y": "Ventas Totales"},
    title="Top 5 Productos Más Vendidos"
)
fig_ventas.update_layout(xaxis_tickangle=0, margin=dict(b=160))
fig_ventas.update_xaxes(
    tickmode="array",
    tickvals=list(top_5_v.index),
    ticktext=[wrap_text(n) for n in top_5_v.index],
)
st.header("Top 5 Productos Más Vendidos")
st.plotly_chart(fig_ventas, use_container_width=True)

# Top 5 por ganancia
top_5_g = ganancias_por_producto.sort_values(ascending=False).head(5)
fig_ganancias = px.bar(
    x=top_5_g.index,
    y=top_5_g.values,
    labels={"x": "Producto", "y": "Ganancias Totales"},
    title="Top 5 Productos con Mayor Ganancia"
)
fig_ganancias.update_layout(xaxis_tickangle=0, margin=dict(b=160))
fig_ganancias.update_xaxes(
    tickmode="array",
    tickvals=list(top_5_g.index),
    ticktext=[wrap_text(n) for n in top_5_g.index],
)
st.header("Top 5 Productos con Mayor Ganancia")
st.plotly_chart(fig_ganancias, use_container_width=True)

# -------------------------
# Dispersión Ventas vs Ganancias
# -------------------------
df_summary = pd.concat([ventas_por_producto, ganancias_por_producto], axis=1)
df_summary.columns = ["Ventas Totales", "Ganancias Totales"]
fig_scatter = px.scatter(
    df_summary,
    x="Ventas Totales",
    y="Ganancias Totales",
    hover_name=df_summary.index,
    title="Relación entre Ventas y Ganancias por Producto",
)
st.header("Relación entre Ventas y Ganancias por Producto")
st.plotly_chart(fig_scatter, use_container_width=True)
