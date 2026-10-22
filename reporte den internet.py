import os
import json
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# 1. ARQUITECTURA VISUAL (Fijar entorno ancho nativo)
st.set_page_config(page_title="Tablero Control de Ahorros y Gestión", layout="wide")

st.title("📊 System Command - Monitoreo de Ahorros e Importes Fuera de Catálogo")
st.markdown("### Auditoría de Procesos, Eficiencia Comercial y Mitigación de Desviaciones")
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
    obra_sel = st.selectbox("🏗️ Frente de Trabajo:", opciones_obra)
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
# 4. BLANCO DE INDICADORES MAESTROS (MÉTRICAS SIN DECIMALES)
# ==============================================================================
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    st.metric(label="💵 Importe Total Contratado", value=f"${total_compra:,}")
with kpi2:
    st.metric(label="🛡️ Mitigación de Costo (Ahorros)", value=f"${total_ahorro_consolidado:,}")
with kpi3:
    st.metric(label="🚨 Desvíos Operativos Fuera de MIX", value=f"${total_fuera_mix:,}")
with kpi4:
    st.metric(label="📊 Efectividad Contable", value=f"{eficiencia_gestion:.1f}%")

st.markdown("---")

# ==============================================================================
# 5. BLOQUE VISUAL REESTRUCTURADO EN PAREJAS DE ALTA DENSIDAD INFORMATIVA
# ==============================================================================

# --- FILA GRÁFICA 1 ---
r1_c1, r1_c2 = st.columns(2)
with r1_c1:
    st.subheader("👤 Fuga Fuera de MIX por Comprador")
    df_resp_g = df_f.groupby(col_resp)[col_fuera_mix].sum().reset_index()
    fig_resp = px.pie(df_resp_g, names=col_resp, values=col_fuera_mix, hole=0.4)
    fig_resp.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_resp, use_container_width=True)
    st.info("🔍 Identifica qué compradores están rompiendo los acuerdos comerciales y autorizando importes fuera de catálogo.")

with r1_c2:
    st.subheader("🏗️ Eficiencia Comercial vs Volumen por Obra")
    df_obra_comp = df_f.groupby(col_obra).agg({col_importe_compra: 'sum', col_fuera_mix: 'sum'}).reset_index().sort_values(by=col_importe_compra, ascending=False).head(5)
    fig_comp_obra = go.Figure()
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_importe_compra], name='Importe Compra'))
    fig_comp_obra.add_trace(go.Bar(x=df_obra_comp[col_obra], y=df_obra_comp[col_fuera_mix], name='Fuga Fuera MIX'))
    fig_comp_obra.update_layout(barmode='group', height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_comp_obra, use_container_width=True)
    st.info("📊 Cruce analítico directo: compara la inyección de presupuesto contra la disciplina operativa por proyecto.")

# --- FILA GRÁFICA 2 ---
r2_c1, r2_c2 = st.columns(2)
with r2_c1:
    st.subheader("🎯 Efectividad de la Operación Contable")
    labels_gestion = ['Estrategia Correcta', 'Incidencia Ineficiente']
    valores_gestion = [max(0, total_gestiones - total_sin_estrategia), total_sin_estrategia]
    fig_gestion = px.pie(names=labels_gestion, values=valores_gestion, hole=0.5)
    fig_gestion.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_gestion, use_container_width=True)
    st.info("💡 Mide el porcentaje de procesos que cumplieron con la validación de tarifas institucionales.")

with r2_c2:
    st.subheader("📅 Velocidad Histórica de Fugas de Capital")
    df_fuga_mes = df_f.groupby(col_mes)[col_fuera_mix].sum().reset_index()
    fig_fuga_mes = px.bar(df_fuga_mes, x=col_mes, y=col_fuera_mix)
    fig_fuga_mes.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_fuga_mes, use_container_width=True)
    st.info("📅 Rastrea de forma temporal en qué periodos mensuales se disparó la indisciplina presupuestal.")

# --- FILA GRÁFICA 3 ---
r3_c1, r3_c2 = st.columns(2)
with r3_c1:
    st.subheader("📉 Captura y Rendimiento de Ahorros")
    df_mes_data = df_f.groupby(col_mes)[[col_ahorro_cv, col_ahorro_prof, col_ahorro_cierre]].sum().reset_index()
    fig_lineas = go.Figure()
    fig_lineas.add_trace(go.Scatter(x=df_mes_data[col_mes], y=df_mes_data[col_ahorro_cv], name='Costo Venta', mode='lines+markers'))
    fig_lineas.add_trace(go.Scatter(x=df_mes_data[col_mes], y=df_mes_data[col_ahorro_prof], name='Proforma', mode='lines+markers'))
    fig_lineas.add_trace(go.Scatter(x=df_mes_data[col_mes], y=df_mes_data[col_ahorro_cierre], name='Cierre', mode='lines+markers'))
    fig_lineas.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_lineas, use_container_width=True)
    st.info("📈 Analiza qué canales de negociación (Costo Venta, Proforma o Cierre) están rindiendo más dinero.")

with r3_c2:
    st.subheader("⚠️ Top 5 Proyectos Críticos con Mayor Desvío")
    df_fugas = df_f.groupby(col_obra)[col_fuera_mix].sum().reset_index().sort_values(by=col_fuera_mix, ascending=True).tail(5)
    fig_fugas = px.bar(df_fugas, y=col_obra, x=col_fuera_mix, orientation='h')
    fig_fugas.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_fugas, use_container_width=True)
    st.info("🚨 Semáforo de alarma: detecta al instante qué frentes de trabajo necesitan auditoría interna inmediata.")
