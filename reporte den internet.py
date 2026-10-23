import os
import json
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# 1. ARQUITECTURA VISUAL (Fijar entorno ancho nativo y paleta institucional)
st.set_page_config(page_title="Tablero Control de Ahorros y Gestión", layout="wide")

# Inyección de estilos con los colores corporativos oficiales de la organización
st.markdown("""
    <style>
    .stApp { background-color: #53565A; color: #FFFFFF; }
    h1, h2, h3, h4, h5, h6, p, label { color: #FFFFFF !important; }
    .header-card { background-color: #0077BE; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header-title { color: #FFCD00; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
    .header-value { color: #FFFFFF; font-size: 24px; font-weight: 800; margin-top: 5px; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { background-color: #FFFFFF !important; color: #53565A !important; }
    div[data-testid="stNotification"] { background-color: #0077BE !important; color: #FFFFFF !important; border-left: 5px solid #FFCD00 !important; }
    </style>
    """, unsafe_allow_html=True)

# Encabezado corporativo con Logotipo Oficial
col_logo, col_titulo = st.columns([1, 4])
with col_logo:
    # URL directa de la imagen oficial del logotipo para renderizado dinámico en la nube
    st.image("https://grupohycsa.com", width=220)
with col_titulo:
    st.title("🦈 System Command - Monitoreo de Ahorros e Importes Fuera de Catálogo")
    st.markdown("<p style='color:#FFCD00; font-size:16px; font-weight:600; margin-top:-10px;'>Auditoría de Procesos, Eficiencia Comercial y Mitigación de Desviaciones Financieras</p>", unsafe_allow_html=True)

st.markdown("---")

# 2. CONEXIÓN DE DATOS ADAPTADA PARA LA NUBE (GitHub)
ruta_excel = "./08 Resultado de ahorros cris.xlsx"

if not os.path.exists(ruta_excel):
    st.error(f"Base de datos no localizada en la ruta configurada: {ruta_excel}")
    st.stop()

# Leer la primera pestaña física del archivo sin importar cómo se llame
archivo_excel = pd.ExcelFile(ruta_excel, engine="openpyxl")
nombre_hoja = archivo_excel.sheet_names[0]
df_raw = pd.read_excel(ruta_excel, sheet_name=nombre_hoja, engine="openpyxl")

# Limpiar espacios vacíos en los títulos de las columnas
df_raw.columns = df_raw.columns.str.strip()

# Mapeo dinámico basado en la posición física real de las columnas de tu Excel de GRUPO HYC
columnas_reales = list(df_raw.columns)

col_obra = columnas_reales[0]          # OBRA
col_mes = columnas_reales[1]           # MES
col_gestiones = columnas_reales[2]     # GESTIONES
col_importe_compra = columnas_reales[3] # IMPORTE DE COMPRA
col_sin_cv = columnas_reales[4]        # GESTIONES SIN COSTO DE VENTA
col_ahorro_cv = columnas_reales[5]     # AHORRO COSTO VENTA
col_sin_prof = columnas_reales[6]      # GESTIONES SIN PROFORMA
col_ahorro_prof = columnas_reales[7]    # AHORRO PROFORMA
col_ahorro_cierre = columnas_reales[8]  # AHORRO NEGOCIACIÓN CIERRE
col_fuera_mix = columnas_reales[9]      # IMPORTE COMPRADO FUERA DE PRECIOS MIX
col_resp = columnas_reales[10]         # Responsable

# Lavado y blindaje contra celdas de texto o vacías (Conversión a enteros puros sin decimales)
cols_monetarias = [col_importe_compra, col_ahorro_cv, col_ahorro_prof, col_ahorro_cierre, col_fuera_mix]
for c in cols_monetarias:
    if c in df_raw.columns:
        df_raw[c] = pd.to_numeric(df_raw[c], errors='coerce').fillna(0).round(0).astype(int)

cols_contadores = [col_gestiones, col_sin_cv, col_sin_prof]
for c in cols_contadores:
    if c in df_raw.columns:
        df_raw[c] = pd.to_numeric(df_raw[c], errors='coerce').fillna(0).round(0).astype(int)

# Limpiar espacios fantasmas en las columnas de texto para el correcto funcionamiento de los filtros
for c in [col_obra, col_mes, col_resp]:
    if c in df_raw.columns:
        df_raw[c] = df_raw[c].astype(str).str.strip()

# ==============================================================================
# 3. FILTROS DINÁMICOS SUPERIORES AVANZADOS
# ==============================================================================
f1, f2, f3 = st.columns(3)
with f1:
    opciones_obra = ["Todas las Obras"] + sorted([x for x in df_raw[col_obra].unique() if x != 'nan'])
    obra_sel = st.selectbox("🏗️ Frente de Trabajo / Proyecto:", opciones_obra)
with f2:
    opciones_mes = ["Todos los Meses"] + sorted([x for x in df_raw[col_mes].unique() if x != 'nan'])
    mes_sel = st.selectbox("📅 Ventana de Tiempo (Mes):", opciones_mes)
with f3:
    opciones_resp = ["Todos los Responsables"] + sorted([x for x in df_raw[col_resp].unique() if x != 'nan'])
    resp_sel = st.selectbox("👤 Comprador Asignado:", opciones_resp)

# Filtrado interactivo en tiempo real
df_f = df_raw.copy()
if obra_sel != "Todas las Obras":
    df_f = df_f[df_f[col_obra] == obra_sel]
if mes_sel != "Todos los Meses":
    df_f = df_f[df_f[col_mes] == mes_sel]
if resp_sel != "Todos los Responsables":
    df_f = df_f[df_f[col_resp] == resp_sel]

# Sumatorias financieras globales formateadas a enteros redondos
total_compra = int(df_f[col_importe_compra].sum())
total_ahorro_cv = int(df_f[col_ahorro_cv].sum())
total_ahorro_prof = int(df_f[col_ahorro_prof].sum())
total_ahorro_cierre = int(df_f[col_ahorro_cierre].sum())
total_ahorro_consolidado = total_ahorro_cv + total_ahorro_prof + total_ahorro_cierre
total_fuera_mix = int(df_f[col_fuera_mix].sum())

total_gestiones = int(df_f[col_gestiones].sum() if df_f[col_gestiones].sum() > 0 else 1)
total_sin_est_cv = int(df_f[col_sin_cv].sum())
total_sin_est_prof = int(df_f[col_sin_prof].sum())
total_sin_estrategia = total_sin_est_cv + total_sin_est_prof
eficiencia_gestion = float(((total_gestiones - total_sin_estrategia) / total_gestiones * 100))

# ==============================================================================
# 4. CONTROLES DE MEDICIÓN CORPORATIVOS (PALETA INSTITUCIONAL)
# ==============================================================================
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    st.markdown(f'<div class="header-card"><div class="header-title">💵 Importe Total Contratado</div><div class="header-value">${total_compra:,}</div></div>', unsafe_allow_html=True)
with kpi2:
    st.markdown(f'<div class="header-card"><div class="header-title">🛡️ Mitigación de Costo (Ahorros)</div><div class="header-value" style="color:#FFCD00;">${total_ahorro_consolidado:,}</div></div>', unsafe_allow_html=True)
with kpi3:
    st.markdown(f'<div class="header-card" style="background-color:#A94442;"><div class="header-title" style="color:#FFFFFF;">🚨 Desvíos Operativos Fuera de MIX</div><div class="header-value">${total_fuera_mix:,}</div></div>', unsafe_allow_html=True)
with kpi4:
    st.markdown(f'<div class="header-card"><div class="header-title">📊 Efectividad Contable</div><div class="header-value" style="color:#FFCD00;">{eficiencia_gestion:.1f}%</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# 5. BLOQUE VISUAL REESTRUCTURADO EN PAREJAS DE ALTA DENSIDAD INFORMATIVA
# ==============================================================================

# --- FILA GRÁFICA 1 ---
r1_c1, r1_c2 = st.columns(2)
with r1_c1:
    st.subheader("👤 Fuga Fuera de MIX por Comprador")
    df_resp_g = df_f.groupby(col_resp)[col_fuera_mix].sum().reset_index()
    fig_resp = px.pie(df_resp_g, names=col_resp, values=col_fuera_mix, hole=0.4,
                      color_discrete_sequence=['#0077BE', '#FFCD00', '#333333', '#CCCCCC'])
    fig_resp.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig_resp.update_traces(textfont_color='#FFFFFF')
    st.plotly_chart(fig_resp, use_container_width=True)
    st.info("🔍 Identifica qué compradores están rompiendo los acuerdos comerciales y autorizando importes fuera de catálogo.")

with r1_c2:
    st.subheader("🏗️ Eficiencia Comercial vs Volumen por Obra")
    df_obra_comp = df_f.groupby(col_obra).agg({col_importe_compra: 'sum', col_fuera_mix: 'sum'}).reset_index().sort_values(by=col_importe_compra, ascending=False).head(5)
    fig_comp_obra = go.Figure()
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_importe_compra], name='Importe Compra', marker_color='#0077BE'))
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_fuera_mix], name='Fuga Fuera MIX', marker_color='#FFCD00'))
    fig_comp_obra.update_layout(barmode='group', height=280, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(font=dict(color="#FFFFFF")))
    st.plotly_chart(fig_comp_obra, use_container_width=True)
    st.info("📊 Cruce analítico directo: compara la inyección de presupuesto contra la disciplina operativa por proyecto.")

# --- FILA GRÁFICA 2 ---
r2_c1, r2_c2 = st.columns(2)
with r2_c1:
    st.subheader("🎯 Efectividad de la Operación Contable")
    labels_gestion = ['Estrategia Correcta', 'Incidencia Ineficiente']
    valores_gestion = [max(0, total_gestiones - total_sin_estrategia), total_sin_estrategia]
    fig_gestion = px.pie(names=labels_gestion, values=valores_gestion, hole=0.5,
                         color_discrete_sequence=['#0077BE', '#A94442'])
    fig_gestion.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig_gestion.update_traces(textfont_color='#FFFFFF')
    st.plotly_chart(fig_gestion, use_container_width=True)
    st.info("💡 Mide el porcentaje de procesos que cumplieron con la validación de tarifas institucionales.")

with r2_c2:
    st.subheader("📅 Velocidad Histórica de Fugas de Capital")
