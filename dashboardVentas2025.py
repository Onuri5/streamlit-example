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
# Carga de datos
# -------------------------
file_path = "Orders Final Limpio2.xlsx"

try:
    df_orders = pd.read_excel(file_path)
except FileNotFoundError:
    st.error(f"No se encontró el archivo '{file_path}'. Asegúrate de que esté en el mismo directorio del script.")
    st.stop()

# 1) Eliminar duplicados exactos
df_orders = df_orders.drop_duplicates().reset_index(drop=True)

# 2) Corregir descuentos expresados como enteros (17 -> 0.17)
if "Discount" in df_orders.columns:
    mask_pct = (df_orders["Discount"] > 1) & (df_orders["Discount"] <= 100)
    df_orders.loc[mask_pct, "Discount"] = df_orders.loc[mask_pct, "Discount"] / 100.0

# -------------------------
# Filtros laterales
# -------------------------
with st.sidebar:
    st.header("Filtros")

    # Fechas: se asume que Ship Date ya está en datetime
    if "Ship Date" in df_orders.columns:
        min_date = df_orders["Ship Date"].min().date()
        max_date = df_orders["Ship Date"].max().date()

        rango = st.date_input(
            "Rango de envío",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="YYYY/MM/DD",
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
    else:
        start_date = end_date = None

    region = None
    estado = None
    if "Region" in df_orders.columns:
        region = st.selectbox("Selecciona Región", ["Todas"] + sorted(df_orders["Region"].dropna().unique().tolist()))
    if "State" in df_orders.columns:
        estado = st.selectbox("Selecciona Estado", ["Todas"] + sorted(df_orders["State"].dropna().unique().tolist()))

    mostrar_tabla = st.checkbox("Mostrar datos filtrados", value=True)

# -------------------------
# Aplicación de filtros
# -------------------------
df_filtered = df_orders.copy()

if "Ship Date" in df_filtered.columns and start_date and end_date:
    mask = (df_filtered["Ship Date"] >= pd.Timestamp(start_date)) & (df_filtered["Ship Date"] <= pd.Timestamp(end_date))
    df_filtered = df_filtered.loc[mask]

if region and region != "Todas":
    df_filtered = df_filtered.loc[df_filtered["Region"] == region]

if estado and estado != "Todas":
    df_filtered = df_filtered.loc[df_filtered["State"] == estado]

if df_filtered.empty:
    st.warning("No hay datos para los filtros seleccionados.")
    st.stop()

st.success("Datos cargados y filtrados correctamente.")

# -------------------------
# Función para envolver texto
# -------------------------
def wrap_text(txt: str, width: int = 22) -> str:
    return "<br>".join(textwrap.wrap(str(txt), width=width))

# -------------------------
# Tabla de datos filtrados
# -------------------------
st.subheader("Datos filtrados")
if mostrar_tabla:
    cols_pref = [
        "Ship Date","Discount","Sales","Quantity","Profit",
        "Region","State","Order ID","Product Name","City"
    ]
    cols_show = [c for c in cols_pref if c in df_filtered.columns]
    if not cols_show:
        cols_show = df_filtered.columns.tolist()

    st.dataframe(
        df_filtered[cols_show].sort_values(cols_show[0]),
        use_container_width=True,
        hide_index=True,
    )

# -------------------------
# Agregaciones y gráficas
# -------------------------
if "Product Name" in df_filtered.columns:
    ventas_por_producto = df_filtered.groupby("Product Name")["Sales"].sum()
    ganancias_por_producto = df_filtered.groupby("Product Name")["Profit"].sum()

    # Top 5 por Ventas
    top_5_v = ventas_por_producto.sort_values(ascending=False).head(5)
    fig_ventas = px.bar(
        x=top_5_v.index,
        y=top_5_v.values,
        labels={"x": "Producto", "y": "Ventas Totales"},
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
        labels={"x": "Producto", "y": "Ganancias Totales"},
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
else:
    st.warning("No se encontró la columna 'Product Name' para generar las gráficas.")
