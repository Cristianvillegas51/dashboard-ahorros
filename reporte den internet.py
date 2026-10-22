import os
import json
import pandas as pd
import streamlit as st
import plotly.express as px

# 1. Configuración de la página (Fondo y estilo corporativo)
st.set_page_config(page_title="Dashboard de Ahorros HYC", layout="wide")

# Estilo CSS inyectado por código para fijar el fondo gris ejecutivo y tarjetas blancas
st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    div[data-testid="stMetricValue"] { color: #0A192F; font-size: 36px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Resumen Ejecutivo de Resultados de Ahorros")

# 2. RUTA INTERNA PARA LA NUBE (Busca el Excel en tu mismo repositorio de GitHub)
ruta_excel = "./08 Resultado de ahorros cris.xlsx"

@st.cache_data
def cargar_datos():
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el archivo de Excel en la ruta: {ruta_excel}")
    
    # Detectar automáticamente el nombre de la primera pestaña que tenga el Excel
    archivo_excel = pd.ExcelFile(ruta_excel, engine="openpyxl")
    primera_pestana = archivo_excel.sheet_names[0]
    
    df = pd.read_excel(ruta_excel, sheet_name=primera_pestana, engine="openpyxl")
    df.columns = df.columns.str.strip()  # Limpiar espacios vacíos en los títulos
    return df

try:
    df_base = cargar_datos()

    # 3. FILTROS DINÁMICOS AUTOMÁTICOS (Fila superior)
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        lista_meses = ["Todos"] + list(df_base['MES'].dropna().unique())
        mes_seleccionado = st.selectbox("📅 Selecciona el Mes:", lista_meses)
    with col_f2:
        lista_obras = ["Todas"] + list(df_base['OBRA'].dropna().unique())
        obra_seleccionada = st.selectbox("🏗️ Selecciona la Obra:", lista_obras)

    # Aplicar filtros dinámicos en tiempo real según la selección del usuario
    df_filtrado = df_base.copy()
    if mes_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado['MES'] == mes_seleccionado]
    if obra_seleccionada != "Todas":
        df_filtrado = df_filtrado[df_filtrado['OBRA'] == obra_seleccionada]

    # 4. KPI AUTOMÁTICO (Tarjeta de ahorro gigante)
    total_costo_venta = df_filtrado['AHORRO COSTO VENTA'].sum()
    st.metric(label="💰 Total Ahorro Costo Venta", value=f"${total_costo_venta:,.2f}")

    # 5. GRÁFICAS EJECUTIVAS AUTOMÁTICAS
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🎯 Distribución por Concepto de Ahorro")
        totales = {
            "Concepto": ["Costo Venta", "Negociación Cierre"],
            "Monto": [df_filtrado['AHORRO COSTO VENTA'].sum(), df_filtrado['AHORRO NEGOCIACIÓN CIERRE'].sum()]
        }
        df_dona = pd.DataFrame(totales)
        fig_pie = px.pie(df_dona, names="Concepto", values="Monto", hole=0.4,
                         color_discrete_sequence=["#0A192F", "#1E3A8A"])
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        st.subheader("📈 Acumulado de Ahorro Costo Venta por MES")
        df_mes = df_filtrado.groupby('MES', as_index=False)['AHORRO COSTO VENTA'].sum()
        fig_bar = px.bar(df_mes, x="MES", y="AHORRO COSTO VENTA", 
                         color_discrete_sequence=["#0A192F"])
        st.plotly_chart(fig_bar, use_container_width=True)

except Exception as e:
    st.error(f"Error al procesar el tablero analítico: {e}")
