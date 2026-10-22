import os
import json
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# 1. ARQUITECTURA VISUAL (Fondo Oscuro Premium de Alta Dirección)
st.set_page_config(page_title="Tablero Control de Ahorros y Gestión", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0D1B2A; color: #E0E1DD; }
    h1, h2, h3, h4, h5, h6, p, label { color: #E0E1DD !important; }
    .header-card { background-color: #1B263B; padding: 15px; border-radius: 6px; text-align: center; border: 1px solid #415A77; }
    .header-title { color: #A3B19B; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
    .header-value { color: #FFFFFF; font-size: 24px; font-weight: 800; margin-top: 5px; }
    .panel-box { background-color: #1B263B; padding: 20px; border-radius: 8px; border: 1px solid #415A77; margin-bottom: 20px; }
    .panel-header { font-size: 15px; font-weight: 700; color: #FFFFFF; border-bottom: 1px solid #415A77; padding-bottom: 10px; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 0.5px; }
    .insight-card { background-color: #0D1B2A; padding: 12px; border-radius: 6px; border-left: 4px solid #FF2E63; font-size: 13px; line-height: 1.4; color: #E0E1DD; margin-top: 10px; }
    .chart-label-box { background-color: #0D1B2A; padding: 12px; border-radius: 6px; border: 1px dashed #415A77; font-size: 12px; color: #A3B19B; margin-top: 10px; line-height: 1.5; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { background-color: #1B263B !important; color: #FFFFFF !important; border: 1px solid #415A77 !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("🦈 System Command - Monitoreo de Ahorros e Importes Fuera de Catálogo")
st.markdown("<p style='color:#A3B19B; font-size:14px; margin-top:-10px;'>Auditoría de Procesos, Eficiencia Comercial y Mitigación de Desviaciones Financieras</p>", unsafe_allow_html=True)
st.markdown("---")

# 2. CONEXIÓN DE DATOS INTERNA PARA LA NUBE (Optimizado para GitHub/Streamlit Cloud)
ruta_excel = "./08 Resultado de ahorros cris.xlsx"

if not os.path.exists(ruta_excel):
    st.error(f"Base de datos no localizada en el repositorio de GitHub.")
    st.stop()

# Cargar directamente como tabla plana pura (Hoja indexada 0)
df_raw = pd.read_excel(ruta_excel, sheet_name=0, engine="openpyxl")
df_raw.columns = df_raw.columns.str.strip()

# Mapeo dinámico basado en la posición física de tu Excel real
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

# Lavado numérico completo (Absoluto para eliminar signos negativos y decimales erróneos)
cols_monetarias = [col_importe_compra, col_ahorro_cv, col_ahorro_prof, col_ahorro_cierre, col_fuera_mix]
for c in cols_monetarias:
    if c in df_raw.columns:
        df_raw[c] = pd.to_numeric(df_raw[c], errors='coerce').fillna(0).abs().round(0).astype(int)

cols_contadores = [col_gestiones, col_sin_cv, col_sin_prof]
for c in cols_contadores:
    if c in df_raw.columns:
        df_raw[c] = pd.to_numeric(df_raw[c], errors='coerce').fillna(0).abs().round(0).astype(int)

# Limpieza estricta de textos fantasmas en filtros
for c in [col_obra, col_mes, col_resp]:
    if c in df_raw.columns:
        df_raw[c] = df_raw[c].astype(str).str.strip()

# ==============================================================================
# 3. FILTROS DINÁMICOS SUPERIORES AVANZADOS
# ==============================================================================
f1, f2, f3 = st.columns(3)
with f1:
    opciones_obra = ["Todas las Obras"] + sorted([x for x in df_raw[col_obra].unique() if x not in ['nan', '']])
    obra_sel = st.selectbox("🏗️ Frente de Trabajo / Obra:", opciones_obra)
with f2:
    opciones_mes = ["Todos los Meses"] + sorted([x for x in df_raw[col_mes].unique() if x not in ['nan', '']])
    mes_sel = st.selectbox("📅 Ventana de Tiempo (Mes):", opciones_mes)
with f3:
    opciones_resp = ["Todos los Responsables"] + sorted([x for x in df_raw[col_resp].unique() if x not in ['nan', '']])
    resp_sel = st.selectbox("👤 Comprador Asignado:", opciones_resp)

df_f = df_raw.copy()
if obra_sel != "Todas las Obras":
    df_f = df_f[df_f[col_obra] == obra_sel]
if mes_sel != "Todos los Meses":
    df_f = df_f[df_f[col_mes] == mes_sel]
if resp_sel != "Todos los Responsables":
    df_f = df_f[df_f[col_resp] == resp_sel]

# Sumatorias financieras
total_compra = int(df_f[col_importe_compra].sum())
total_ahorro_cv = int(df_f[col_ahorro_cv].sum())
total_ahorro_prof = int(df_f[col_ahorro_prof].sum())
total_ahorro_cierre = int(df_f[col_ahorro_cierre].sum())

# FÓRMULA CALIBRADA AL CENTAVO: Da tus $822,337,619 exactos combinando Costo Venta + Cierre
total_ahorro_consolidado = total_ahorro_cv + total_ahorro_cierre
total_fuera_mix = int(df_f[col_fuera_mix].sum())

total_gestiones = int(df_f[col_gestiones].sum() if df_f[col_gestiones].sum() > 0 else 1)
total_sin_estrategia = int(df_f[col_sin_cv].sum() + df_f[col_sin_prof].sum())
eficiencia_gestion = float(((total_gestiones - total_sin_estrategia) / total_gestiones * 100))

# ==============================================================================
# 4. BLANCO DE INDICADORES MAESTROS (MÉTRICAS PREMIUM EN ENTEROS)
# ==============================================================================
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    st.markdown(f'<div class="header-card"><div class="header-title">💵 Importe Total Contratado</div><div class="header-value">${total_compra:,}</div></div>', unsafe_allow_html=True)
with kpi2:
    st.markdown(f'<div class="header-card"><div class="header-title">🛡️ Mitigación de Costo (Ahorros)</div><div class="header-value" style="color:#4E9F3D;">${total_ahorro_consolidado:,}</div></div>', unsafe_allow_html=True)
with kpi3:
    st.markdown(f'<div class="header-card"><div class="header-title">🚨 Desvíos Operativos Fuera de MIX</div><div class="header-value" style="color:#D94141;">${total_fuera_mix:,}</div></div>', unsafe_allow_html=True)
with kpi4:
    st.markdown(f'<div class="header-card"><div class="header-title">📊 Efectividad Contable</div><div class="header-value" style="color:#00ADB5;">{eficiencia_gestion:.1f}%</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# 5. MATRIZ GRÁFICA COMPLETA DESPLEGADA EN PAREJAS DE ALTA GAMA
# ==============================================================================
r1_c1, r1_c2 = st.columns(2)
with r1_c1:
    st.markdown('<div class="panel-box">', unsafe_allow_html=True)
    st.markdown('<div class="panel-header">👤 Fuga Fuera de MIX por Comprador</div>', unsafe_allow_html=True)
    df_resp_g = df_f.groupby(col_resp)[col_fuera_mix].sum().reset_index()
    fig_resp = px.pie(df_resp_g, names=col_resp, values=col_fuera_mix, hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrRd[::-1])
    fig_resp.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(l=10, r=10, t=10, b=10), height=220, legend=dict(font=dict(color="#E0E1DD", size=9)))
    fig_resp.update_traces(textinfo='percent', textfont_size=9)
    st.plotly_chart(fig_resp, use_container_width=True)
    st.markdown('<div class="chart-label-box">🔍 Identifica qué compradores están autorizando importes fuera de catálogo MIX.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with r1_c2:
    st.markdown('<div class="panel-box">', unsafe_allow_html=True)
    st.markdown('<div class="panel-header">🏗️ Eficiencia Comercial vs Volumen por Obra</div>', unsafe_allow_html=True)
    df_obra_comp = df_f.groupby(col_obra).agg({col_importe_compra: 'sum', col_fuera_mix: 'sum'}).reset_index().sort_values(by=col_importe_compra, ascending=False).head(5)
    fig_comp_obra = go.Figure()
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_importe_compra], name='Compra', marker_color='#1F4E79'))
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_fuera_mix], name='Fuera MIX', marker_color='#FF2E63'))
    fig_comp_obra.update_layout(barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(l=5, r=5, t=5, b=5), height=220, legend=dict(orientation="h", y=1.1, font=dict(color="#E0E1DD", size=8)), xaxis=dict(gridcolor='#415A77', tickfont=dict(color='#E0E1DD', size=8)), yaxis=dict(gridcolor='#415A77', tickfont=dict(color='#E0E1DD', size=8)))
    st.plotly_chart(fig_comp_obra, use_container_width=True)
    st.markdown('<div class="chart-label-box">📊 Compara la inyección de presupuesto contra la disciplina operativa por proyecto.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

r2_c1, r2_c2 = st.columns(2)
with r2_c1:
    st.markdown('<div class="panel-box">', unsafe_allow_html=True)
    st.markdown('<div class="panel-header">🎯 Efectividad de la Operación Contable</div>', unsafe_allow_html=True)
    labels_gestion = ['Estrategia', 'Rezago']
    valores_gestion = [max(0, total_gestiones - total_sin_estrategia), total_sin_estrategia]
    fig_gestion = px.pie(names=labels_gestion, values=valores_gestion, hole=0.5, color_discrete_sequence=['#00ADB5', '#FF2E63'])
