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
file_path = "Orders Final Limpio2.xlsx"
df_orders = pd.read_excel(file_path)

# 1) Eliminar duplicados exactos
df_orders = df_orders.drop_duplicates().reset_index(drop=True)

# 2) Corregir descuentos expresados como enteros (17 -> 0.17)
if "Discount" in df_orders.columns:
    mask_pct = (df_orders["Discount"] > 1) & (df_orders["Discount"] <= 100)
    df_orders.loc[mask_pct, "Discount"] = df_orders.loc[mask_pct, "Discount"] / 100.0

# 3) Unificar Ship Date y remover columna duplicada si existe
if "Ship Date" in df_orders.columns and "Ship date" in df_orders.columns:
    sd_main = pd.to_datetime(df_orders["Ship Date"], errors="coerce")
    sd_alt  = pd.to_datetime(df_orders["Ship date"], errors="coerce")
    df_orders["Ship Date"] = sd_main.fillna(sd_alt)
    df_orders = df_orders.drop(columns=["Ship date"])

# 4) Normalizar columna de fecha (Order Date)
col_fecha = "Order Date"
if pd.api.types.is_datetime64_any_dtype(df_orders[col_fecha]):
    pass
elif pd.api.types.is_numeric_dtype(df_orders[col_fecha]):
    origin_date = pd.Timestamp("1899-12-30")
    df_orders[col_fecha] = pd.to_timedelta(df_orders[col_fecha], unit="D") + origin_date
elif pd.api.types.is_timedelta64_dtype(df_orders[col_fecha]):
    origin_date = pd.Timestamp("1899-12-30")
    df_orders[col_fecha] = origin_date + df_orders[col_fecha]
else:
    df_orders[col_fecha] = pd.to_datetime(df_orders[col_fecha], errors="coerce")

if df_orders[col_fecha].isna().all():
    st.error("No se pudo convertir correctamente la columna 'Order Date' a fecha.")
    st.stop()

# -------------------------
# Filtros laterales
# -------------------------
with st.sidebar:
    st.header("Filtros")

    min_date = df_orders[col_fecha].min().date()
    max_date = df_orders[col_fecha].max().date()

    rango = st.date_input(
        "Rango de fechas",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="YYYY/MM/DD",
        help="Solo puedes elegir fechas dentro del rango disponible en los datos.",
    )

    if isinstance(rango, tuple) and len(rango) == 2:
        start_date, end_date = rango
    else:
        start_date, end_date = min_date, max_date

    clipped = False
    if start_date < min_date:
        start_date = min_date; clipped = True
    if end_date > max_date:
        end_date = max_date; clipped = True
    if clipped:
        st.info("Las fechas seleccionadas se ajustaron automáticamente al rango disponible en los datos.")

    if start_date > end_date:
        start_date, end_date = end_date, start_date
        st.warning("La fecha de inicio era mayor que la de fin. Se invirtieron para continuar.")

    region = None
    estado = None
    if "Region" in df_orders.columns:
        region = st.selectbox("Selecciona Región", ["Todas"] + sorted(df_orders["Region"].dropna().unique().tolist()))
    if "State" in df_orders.columns:
        estado = st.selectbox("Selecciona Estado", ["Todas"] + sorted(df_orders["State"].dropna().unique().tolist()))

    mostrar_tabla = st.checkbox("Mostrar datos filtrados", value=True)

# Aplicación de filtros
start_ts = pd.Timestamp(start_date)
end_ts = pd.Timestamp(end_date)

mask = (df_orders[col_fecha] >= start_ts) & (df_orders[col_fecha] <= end_ts)
if region and region != "Todas":
    mask &= (df_orders["Region"] == region)
if estado and estado != "Todas":
    mask &= (df_orders["State"] == estado)

df_filtered = df_orders.loc[mask].copy()

if df_filtered.empty:
    st.warning("No hay datos para el rango de fechas (y filtros) seleccionado.")
    st.stop()

st.success("Datos cargados y filtrados correctamente.")

# -------------------------
# Utilidad: envolver etiquetas largas
# -------------------------
def wrap_text(txt: str, width: int = 22) -> str:
    return "<br>".join(textwrap.wrap(str(txt), width=width))

# -------------------------
# Tabla de datos filtrados
# -------------------------
st.subheader("Datos filtrados")
if mostrar_tabla:
    cols_pref = [
        "Order Date","Discount","Sales","Quantity","Profit",
        "Region","State","Order ID","Ship Date","Product Name","City"
    ]
    cols_show = [c for c in cols_pref if c in df_filtered.columns]
    if not cols_show:
        cols_show = df_filtered.columns.tolist()

    st.dataframe(
        df_filtered[cols_show].sort_values(col_fecha),
        use_container_width=True,
        hide_index=True,
    )

# -------------------------
# Agregaciones para gráficas
# -------------------------
ventas_por_producto = df_filtered.groupby("Product Name")["Sales"].sum()
ganancias_por_producto = df_filtered.groupby("Product Name")["Profit"].sum()

# Top 5 por Ventas
top_5_v = ventas_por_producto.sort_values(ascending=False).head(5)
fig_ventas = px.bar(
    x=top_5_v.index,
    y=top_5_v.values,
    labels={"x": "Nombre del Producto", "y": "Ventas Totales"},
    title="Top 5 Productos Más Vendidos",
)
fig_ventas.update_layout(xaxis_tickangle=0, margin=dict(b=160))
fig_ventas.update_xaxes(
    tickmode="array",
    tickvals=list(top_5_v.index),
    ticktext=[wrap_text(n) for n in top_5_v.index],
)
st.header("Top 5 Productos Más Vendidos")
st.plotly_chart(fig_ventas, use_container_width=True)

# Top 5 por Ganancia
top_5_g = ganancias_por_producto.sort_values(ascending=False).head(5)
fig_ganancias = px.bar(
    x=top_5_g.index,
    y=top_5_g.values,
    labels={"x": "Nombre del Producto", "y": "Ganancias Totales"},
    title="Top 5 Productos con Mayor Ganancia",
)
fig_ganancias.update_layout(xaxis_tickangle=0, margin=dict(b=160))
fig_ganancias.update_xaxes(
    tickmode="array",
    tickvals=list(top_5_g.index),
    ticktext=[wrap_text(n) for n in top_5_g.index],
)
st.header("Top 5 Productos con Mayor Ganancia")
st.plotly_chart(fig_ganancias, use_container_width=True)

# Dispersión Ventas vs Ganancias
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

