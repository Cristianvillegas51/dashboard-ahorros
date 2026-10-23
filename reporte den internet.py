# ==============================================================================
# TABLERO DE CONTROL DE AHORROS Y GESTIÓN - GRUPO HYCSA
# ==============================================================================
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image, ImageOps

st.set_page_config(page_title="Tablero Control de Ahorros y Gestión", layout="wide")

try:
    CARPETA = Path(__file__).parent
except NameError:
    CARPETA = Path.cwd()

NOMBRE_EXCEL = "08 Resultado de ahorros cris.xlsx"

# Responsables escritos distinto que son la misma persona (en MAYÚSCULAS).
# Si SANDY y SANDRA fueran personas distintas, borra esa línea.
UNIFICAR_RESPONSABLES = {"SANDY": "SANDRA"}

AMARILLO = "#FFCD00"
AZUL = "#00A3E0"
BLANCO = "#FFFFFF"
ROJO = "#FF6B6B"
PALETA = [AMARILLO, AZUL, BLANCO, "#9FD8F2", "#C8CACC", "#F28C28", "#B58AF5", "#7BD88F"]
MM = 1_000_000  # las gráficas se muestran en millones de pesos

ORDEN_MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
               "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

OPC_EF_A = "Promedio de cobertura (costo de venta y proforma)"
OPC_EF_B = "Gestión inefectiva si le falta cualquiera de las dos"
OPC_AH_A = "Suma de las tres columnas de ahorro"
OPC_AH_B = "Sin duplicar (si costo de venta = proforma, cuenta una sola vez)"
OPC_MIX_A = "Magnitud (valor absoluto)"
OPC_MIX_B = "Solo valores positivos"
OPC_MIX_C = "Solo valores negativos (en positivo)"

st.markdown("""
    <style>
    .stApp { background-color: #53565A; color: #FFFFFF; }
    h1, h2, h3, h4, h5, h6, p, label { color: #FFFFFF !important; }
    .stMetric { background-color: #0077BE; padding: 15px; border-radius: 6px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-bottom: 4px solid #FFCD00; }
    div[data-testid="stMetricLabel"] p { color: #FFCD00 !important; font-weight: 700 !important; text-transform: uppercase; font-size: 11px !important; letter-spacing: 0.5px; }
    div[data-testid="stMetricValue"] div { color: #FFFFFF !important; font-weight: 800 !important; font-size: 22px !important; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { background-color: #FFFFFF !important; color: #53565A !important; }
    div[data-testid="stNotification"] { background-color: #0077BE !important; color: #FFFFFF !important; border-left: 5px solid #FFCD00 !important; }
    </style>
    """, unsafe_allow_html=True)


# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================
def estilo_grafica(fig, alto=320):
    """Panel oscuro, texto blanco y rejillas suaves para que los colores resalten."""
    fig.update_layout(
        height=alto,
        margin=dict(l=20, r=20, t=30, b=20),
        paper_bgcolor="rgba(0,0,0,0.25)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=BLANCO, size=12),
        legend=dict(font=dict(color=BLANCO)),
    )
    fig.update_xaxes(color=BLANCO, gridcolor="rgba(255,255,255,0.12)", linecolor="rgba(255,255,255,0.4)")
    fig.update_yaxes(color=BLANCO, gridcolor="rgba(255,255,255,0.12)", linecolor="rgba(255,255,255,0.4)")
    return fig


def cargar_logo():
    """Busca una imagen con 'logo', 'hycsa' o 'captura' en el nombre y le pone fondo blanco."""
    exts = (".png", ".jpg", ".jpeg", ".webp")
    archivos = [p for p in CARPETA.iterdir() if p.is_file() and p.suffix.lower() in exts]

    def prioridad(p):
        n = p.stem.lower()
        if n.startswith("logo"):
            return 0
        if "hycsa" in n:
            return 1
        if "captura" in n:
            return 2
        return 9

    candidatos = sorted([p for p in archivos if prioridad(p) < 9], key=prioridad)
    if not candidatos:
        return None, archivos
    img = Image.open(candidatos[0]).convert("RGBA")
    fondo = Image.new("RGBA", img.size, (255, 255, 255, 255))
    fondo.alpha_composite(img)
    return ImageOps.expand(fondo.convert("RGB"), border=16, fill="white"), archivos


def normalizar(texto):
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return " ".join(t.upper().split())


def a_numero(serie):
    """Convierte a número aunque venga como texto: '$1,234.50', '-$500', '$-', '(300)', '0.00'."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").fillna(0.0), 0
    s = serie.astype(str).str.strip()
    negativo = s.str.startswith("(") & s.str.endswith(")")
    s = s.str.replace(r"[\$,\s\(\)]", "", regex=True)
    s = s.replace({"-": "0", "": "0", "nan": "0", "NaN": "0", "None": "0", "NaT": "0"})
    num = pd.to_numeric(s, errors="coerce")
    no_leidos = int(num.isna().sum())
    num = num.fillna(0.0)
    num = num.where(~negativo, -num.abs())
    return num, no_leidos


def mes_a_texto(serie):
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie.dt.month.map(lambda x: ORDEN_MESES[int(x) - 1] if pd.notna(x) else "")
    if pd.api.types.is_numeric_dtype(serie):
        return serie.map(lambda x: ORDEN_MESES[int(x) - 1] if pd.notna(x) and 1 <= int(x) <= 12 else "")
    return serie.astype(str).str.strip().str.upper()


def m(x):
    return ("-" if x < 0 else "") + f"${abs(x):,.0f}"


def orden_mes(mes):
    return ORDEN_MESES.index(mes) if mes in ORDEN_MESES else 99


# ==============================================================================
# ENCABEZADO CORPORATIVO
# ==============================================================================
col_logo, col_titulo = st.columns([1, 3])
with col_logo:
    logo, imagenes = cargar_logo()
    if logo is not None:
        st.image(logo, width=240)
    else:
        st.warning(f"No encontré el logo. Guarda 'logo_hycsa.png' en: {CARPETA}")
        if imagenes:
            st.caption("Imágenes que sí veo en la carpeta: " + ", ".join(p.name for p in imagenes))
with col_titulo:
    st.title("🦈 System Command - Monitoreo de Ahorros")
    st.markdown(
        "<h4 style='color:#FFCD00; margin-top:-15px;'>Auditoría de Procesos, Eficiencia Comercial y Mitigación de Desviaciones</h4>",
        unsafe_allow_html=True,
    )
st.markdown("---")

# ==============================================================================
# 1. CONEXIÓN DE DATOS
# ==============================================================================
candidatas = [
    CARPETA / NOMBRE_EXCEL,
    Path(r"C:\Users\CristianVillegasArri\OneDrive - GRUPO HYC, S.A. de C.V\Descargas\MiDashboard") / NOMBRE_EXCEL,
]
ruta_excel = next((r for r in candidatas if r.exists()), None)
if ruta_excel is None:
    st.error(f"No encontré '{NOMBRE_EXCEL}'. Colócalo en la misma carpeta que este archivo: {CARPETA}")
    st.stop()

archivo_excel = pd.ExcelFile(ruta_excel, engine="openpyxl")

cfg = st.expander("⚙️ Configuración de columnas, criterios de cálculo y auditoría")
with cfg:
    hoja = st.selectbox("Hoja del Excel", archivo_excel.sheet_names, index=0)

df_raw = pd.read_excel(ruta_excel, sheet_name=hoja, engine="openpyxl")
df_raw = df_raw.dropna(how="all")
df_raw.columns = [str(c).strip() for c in df_raw.columns]
opciones = list(df_raw.columns)


def detectar(claves, excluir=(), pos=None):
    """Busca la columna por palabras completas; si no la halla, usa la posición."""
    for c in opciones:
        palabras = set(normalizar(c).split())
        if all(k in palabras for k in claves) and not any(e in palabras for e in excluir):
            return c
    if pos is not None and pos < len(opciones):
        return opciones[pos]
    return opciones[0]


CAMPOS = [
    ("Obra", ["OBRA"], [], 0),
    ("Mes original", ["MES"], ["CORRECTO"], 1),
    ("Gestiones", ["GESTIONES"], ["SIN"], 2),
    ("Importe de compra", ["IMPORTE", "COMPRA"], ["FUERA"], 3),
    ("Gestiones sin costo de venta", ["SIN", "COSTO"], ["AHORRO"], 4),
    ("Ahorro costo de venta", ["AHORRO", "COSTO"], ["SIN"], 5),
    ("Gestiones sin proforma", ["SIN", "PROFORMA"], ["AHORRO"], 6),
    ("Ahorro proforma", ["AHORRO", "PROFORMA"], ["SIN"], 7),
    ("Ahorro negociación cierre", ["NEGOCIACION"], [], 8),
    ("Importe fuera de MIX", ["FUERA"], [], 9),
    ("Responsable", ["RESPONSABLE"], [], 10),
]

with cfg:
    st.caption("Verifica que cada concepto apunte a la columna correcta de tu Excel; si no, corrígelo aquí.")
    ui = st.columns(3)
    sel = {}
    for i, (etiqueta, claves, excluir, pos) in enumerate(CAMPOS):
        det = detectar(claves, excluir, pos)
        with ui[i % 3]:
            sel[etiqueta] = st.selectbox(etiqueta, opciones, index=opciones.index(det), key=f"col_{i}")

    mc_det = next((c for c in opciones if {"MES", "CORRECTO"} <= set(normalizar(c).split())), None)
    opciones_mes = ["(usar Mes original)"] + opciones
    with ui[len(CAMPOS) % 3]:
        sel_mes = st.selectbox("Mes para el filtro (MES CORRECTO)", opciones_mes,
                               index=opciones_mes.index(mc_det) if mc_det else 0, key="col_mes")

    st.markdown("**Criterios de cálculo**")
    criterio = st.radio("Efectividad contable", [OPC_EF_A, OPC_EF_B], index=0)
    modo_ahorro = st.radio("Ahorro consolidado", [OPC_AH_A, OPC_AH_B], index=0)
    modo_mix = st.radio("Importe fuera de MIX", [OPC_MIX_A, OPC_MIX_B, OPC_MIX_C], index=0)

col_mes_fuente = sel["Mes original"] if sel_mes.startswith("(") else sel_mes

# ------------------------------------------------------------------------------
# Limpieza: los importes conservan su SIGNO (un ahorro negativo es un sobrecosto)
# ------------------------------------------------------------------------------
mapa_num = {
    "GESTIONES": sel["Gestiones"],
    "IMPORTE": sel["Importe de compra"],
    "SIN_CV": sel["Gestiones sin costo de venta"],
    "AHORRO_CV": sel["Ahorro costo de venta"],
    "SIN_PROF": sel["Gestiones sin proforma"],
    "AHORRO_PROF": sel["Ahorro proforma"],
    "AHORRO_CIERRE": sel["Ahorro negociación cierre"],
    "FUERA_RAW": sel["Importe fuera de MIX"],
}
limpio, no_leidos = {}, {}
for nombre, col in mapa_num.items():
    serie, malos = a_numero(df_raw[col])
    limpio[nombre] = serie.round(2)
    no_leidos[nombre] = malos

df = pd.DataFrame({
    "FILA_EXCEL": df_raw.index + 2,
    "OBRA": df_raw[sel["Obra"]].astype(str).str.replace(r"\s+", " ", regex=True).str.strip(),
    "MES": mes_a_texto(df_raw[col_mes_fuente]),
    "RESPONSABLE": (df_raw[sel["Responsable"]].astype(str).str.strip().str.upper()
                    .replace(UNIFICAR_RESPONSABLES).str.title()),
    **limpio,
})

invalidas = (df["OBRA"].str.lower().isin(["nan", "none", ""])
             | df["OBRA"].str.upper().isin(["TOTAL", "TOTALES", "GRAN TOTAL"]))
filas_excluidas = int(invalidas.sum())
df = df[~invalidas].copy()

# Columnas derivadas según los criterios elegidos
if modo_mix == OPC_MIX_A:
    df["FUERA_MIX"] = df["FUERA_RAW"].abs()
elif modo_mix == OPC_MIX_B:
    df["FUERA_MIX"] = df["FUERA_RAW"].clip(lower=0)
else:
    df["FUERA_MIX"] = (-df["FUERA_RAW"]).clip(lower=0)

if modo_ahorro == OPC_AH_A:
    ahorro_prof_ef = df["AHORRO_PROF"]
else:
    ahorro_prof_ef = df["AHORRO_PROF"].where(df["AHORRO_PROF"] != df["AHORRO_CV"], 0.0)
df["AHORRO_TOTAL"] = df["AHORRO_CV"] + ahorro_prof_ef + df["AHORRO_CIERRE"]

# Una gestión no puede incumplir más veces que gestiones tiene el renglón
df["SIN_CV_C"] = np.minimum(np.maximum(df["SIN_CV"], 0), df["GESTIONES"])
df["SIN_PROF_C"] = np.minimum(np.maximum(df["SIN_PROF"], 0), df["GESTIONES"])
if criterio == OPC_EF_A:
    df["BASE_EF"] = df["GESTIONES"] * 2
    df["INCUMP"] = df["SIN_CV_C"] + df["SIN_PROF_C"]
    unidad = "validaciones"
else:
    df["BASE_EF"] = df["GESTIONES"]
    df["INCUMP"] = np.maximum(df["SIN_CV_C"], df["SIN_PROF_C"])
    unidad = "gestiones"

# ==============================================================================
# 2. FILTROS SUPERIORES
# ==============================================================================


def validos(serie):
    return [x for x in serie.unique() if str(x).strip().lower() not in ("nan", "none", "nat", "")]


f1, f2, f3 = st.columns(3)
with f1:
    obra_sel = st.selectbox("🏗️ Frente de Trabajo:", ["Todas las Obras"] + sorted(validos(df["OBRA"])))
with f2:
    mes_sel = st.selectbox("📅 Ventana de Tiempo (Mes Correcto):",
                           ["Todos los Meses"] + sorted(validos(df["MES"]), key=orden_mes))
with f3:
    resp_sel = st.selectbox("👤 Comprador Asignado:",
                            ["Todos los Responsables"] + sorted(validos(df["RESPONSABLE"])))

df_f = df.copy()
if obra_sel != "Todas las Obras":
    df_f = df_f[df_f["OBRA"] == obra_sel]
if mes_sel != "Todos los Meses":
    df_f = df_f[df_f["MES"] == mes_sel]
if resp_sel != "Todos los Responsables":
    df_f = df_f[df_f["RESPONSABLE"] == resp_sel]

# ==============================================================================
# 3. INDICADORES
# ==============================================================================
total_compra = float(df_f["IMPORTE"].sum())
total_ahorro = float(df_f["AHORRO_TOTAL"].sum())
total_fuera = float(df_f["FUERA_MIX"].sum())
total_gestiones = float(df_f["GESTIONES"].sum())
base_ef = float(df_f["BASE_EF"].sum())
incumplidas = float(df_f["INCUMP"].sum())
cumplidas = base_ef - incumplidas
eficiencia = (cumplidas / base_ef * 100) if base_ef > 0 else 0.0
cov_cv = (1 - df_f["SIN_CV_C"].sum() / total_gestiones) * 100 if total_gestiones > 0 else 0.0
cov_pr = (1 - df_f["SIN_PROF_C"].sum() / total_gestiones) * 100 if total_gestiones > 0 else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("💵 Importe Total Contratado", m(total_compra))
k2.metric("🛡️ Mitigación de Costo (Ahorros)", m(total_ahorro))
k3.metric("🚨 Desvíos Operativos Fuera de MIX", m(total_fuera))
k4.metric("📊 Efectividad Contable", f"{eficiencia:.1f}%")
st.caption(f"Cobertura con costo de venta: {cov_cv:.1f}%  ·  Cobertura con proforma: {cov_pr:.1f}%  ·  "
           f"Gestiones analizadas: {int(total_gestiones):,}")
st.markdown("---")

# ==============================================================================
# 4. GRÁFICAS
# ==============================================================================
# ---------- FILA 1 ----------
r1a, r1b = st.columns(2)

with r1a:
    st.subheader("👤 Fuga Fuera de MIX por Comprador")
    d = df_f.groupby("RESPONSABLE", as_index=False)["FUERA_MIX"].sum()
    d = d[d["FUERA_MIX"] > 0]
    if d.empty:
        st.info("Sin importes fuera de MIX con el filtro actual.")
    else:
        fig = px.pie(d, names="RESPONSABLE", values="FUERA_MIX", hole=0.4, color_discrete_sequence=PALETA)
        fig.update_traces(textposition="outside", textinfo="percent+label",
                          marker=dict(line=dict(color=BLANCO, width=1.5)))
        fig.update_layout(showlegend=False)
        estilo_grafica(fig)
        st.plotly_chart(fig, width="stretch")
    st.info("🔍 Identifica qué compradores concentran los importes fuera de catálogo.")

with r1b:
    st.subheader("🏗️ Importe de Compra vs Fuera de MIX por Obra")
    d = (df_f.groupby("OBRA", as_index=False)[["IMPORTE", "FUERA_MIX"]].sum()
         .sort_values("IMPORTE", ascending=False).head(8))
    fig = go.Figure()
    fig.add_trace(go.Bar(x=d["OBRA"], y=(d["IMPORTE"] / MM).round(2), name="Importe de compra (eje izq.)",
                         marker_color=AZUL))
    fig.add_trace(go.Scatter(x=d["OBRA"], y=(d["FUERA_MIX"] / MM).round(2), name="Fuera de MIX (eje der.)",
                             mode="lines+markers", yaxis="y2",
                             line=dict(color=AMARILLO, width=3), marker=dict(size=9, color=AMARILLO)))
    fig.update_layout(
        yaxis=dict(title="Millones de $"),
        yaxis2=dict(title="Millones de $", overlaying="y", side="right", showgrid=False, color=BLANCO),
        legend=dict(orientation="h", y=-0.35),
    )
    estilo_grafica(fig)
    st.plotly_chart(fig, width="stretch")
    st.info("📊 Dos escalas distintas: así se aprecia el desvío aunque sea pequeño frente al importe contratado.")

# ---------- FILA 2 ----------
r2a, r2b = st.columns(2)

with r2a:
    st.subheader("🎯 Efectividad de la Operación Contable")
    if base_ef <= 0:
        st.info("Sin gestiones con el filtro actual.")
    else:
        fig = px.pie(names=[f"Cumplido ({unidad})", f"Incumplido ({unidad})"],
                     values=[cumplidas, incumplidas], hole=0.6, color_discrete_sequence=[AZUL, ROJO])
        fig.update_traces(textposition="outside", textinfo="percent",
                          marker=dict(line=dict(color=BLANCO, width=1.5)))
        fig.add_annotation(text=f"<b>{eficiencia:.1f}%</b>", x=0.5, y=0.5, showarrow=False,
                           font=dict(size=26, color=AMARILLO))
        estilo_grafica(fig)
        st.plotly_chart(fig, width="stretch")
    st.info("💡 Porcentaje de validaciones (costo de venta y proforma) que sí se cumplieron.")

with r2b:
    st.subheader("👥 Efectividad por Comprador")
    d = df_f.groupby("RESPONSABLE", as_index=False)[["BASE_EF", "INCUMP"]].sum()
    d = d[d["BASE_EF"] > 0].copy()
    if d.empty:
        st.info("Sin gestiones con el filtro actual.")
    else:
        d["EFECTIVIDAD"] = ((1 - d["INCUMP"] / d["BASE_EF"]) * 100).round(1)
        d = d.sort_values("EFECTIVIDAD")
        fig = px.bar(d, x="EFECTIVIDAD", y="RESPONSABLE", orientation="h", text="EFECTIVIDAD",
                     color_discrete_sequence=[AMARILLO])
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
        fig.update_xaxes(range=[0, 115], title_text="% de efectividad")
        fig.update_yaxes(title_text="")
        estilo_grafica(fig)
        st.plotly_chart(fig, width="stretch")
    st.info("📌 Compara la disciplina de validación de cada comprador con el mismo criterio.")

# ---------- FILA 3 ----------
r3a, r3b = st.columns(2)

with r3a:
    st.subheader("🛡️ Ahorro Consolidado por Obra")
    d = df_f.groupby("OBRA", as_index=False)["AHORRO_TOTAL"].sum()
    d = d[d["AHORRO_TOTAL"] != 0].sort_values("AHORRO_TOTAL")
    if d.empty:
        st.info("Sin ahorros con el filtro actual.")
    else:
        d["MILLONES"] = (d["AHORRO_TOTAL"] / MM).round(2)
        d["TIPO"] = np.where(d["AHORRO_TOTAL"] >= 0, "Ahorro", "Sobrecosto")
        fig = px.bar(d, x="MILLONES", y="OBRA", orientation="h", color="TIPO", text="MILLONES",
                     color_discrete_map={"Ahorro": AZUL, "Sobrecosto": ROJO})
        fig.update_traces(texttemplate="%{text:,.1f}", textposition="outside", cliponaxis=False)
        fig.update_xaxes(title_text="Millones de $")
        fig.update_yaxes(title_text="")
        fig.update_layout(legend_title_text="")
        estilo_grafica(fig, alto=max(320, 26 * len(d) + 80))
        st.plotly_chart(fig, width="stretch")
    st.info("🧾 Las barras rojas son obras donde el resultado neto fue sobrecosto, no ahorro.")

with r3b:
    st.subheader("📅 Ahorro vs Fuga Fuera de MIX por Mes")
    d = df_f.groupby("MES", as_index=False).agg(AHORRO=("AHORRO_TOTAL", "sum"), FUGA=("FUERA_MIX", "sum"))
    d = d[d["MES"].isin(validos(d["MES"]))].copy()
    d["ORDEN"] = d["MES"].map(orden_mes)
    d = d.sort_values("ORDEN")
    if d.empty:
        st.info("Sin datos con el filtro actual.")
    else:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=d["MES"], y=(d["AHORRO"] / MM).round(2), name="Ahorro consolidado", marker_color=AZUL))
        fig.add_trace(go.Bar(x=d["MES"], y=(d["FUGA"] / MM).round(2), name="Fuera de MIX", marker_color=AMARILLO))
        fig.update_layout(barmode="group", yaxis_title="Millones de $", legend=dict(orientation="h", y=-0.2))
        estilo_grafica(fig)
        st.plotly_chart(fig, width="stretch")
    st.info("📈 Evolución mensual: ahorro logrado frente a compras fuera de catálogo.")

# ==============================================================================
# 5. AUDITORÍA DE DATOS
# ==============================================================================
with st.expander("🔍 Auditoría de datos (compara contra tu Excel)"):
    st.markdown("**Totales con el filtro actual**")
    resumen = pd.DataFrame({
        "Concepto": [
            "Filas en el filtro", "Gestiones", "Gestiones sin costo de venta (recortadas)",
            "Gestiones sin proforma (recortadas)", "Importe de compra", "Ahorro costo de venta",
            "Ahorro proforma", "Ahorro negociación cierre", "Ahorro consolidado (criterio elegido)",
            "Fuera de MIX con signo (neto)", "Fuera de MIX (criterio elegido)",
            f"Base de efectividad ({unidad})", f"Incumplidas ({unidad})",
        ],
        "Total": [
            f"{len(df_f):,}", f"{int(total_gestiones):,}", f"{int(df_f['SIN_CV_C'].sum()):,}",
            f"{int(df_f['SIN_PROF_C'].sum()):,}", m(total_compra), m(df_f["AHORRO_CV"].sum()),
            m(df_f["AHORRO_PROF"].sum()), m(df_f["AHORRO_CIERRE"].sum()), m(total_ahorro),
            m(df_f["FUERA_RAW"].sum()), m(total_fuera), f"{int(base_ef):,}", f"{int(incumplidas):,}",
        ],
    })
    st.dataframe(resumen, hide_index=True, width="stretch")

    st.markdown("**Calidad de la base (todas las filas, sin filtro)**")
    st.write(f"Filas leídas: {len(df_raw):,} · usadas: {len(df):,} · excluidas (vacías o TOTAL): {filas_excluidas}")
    st.write(f"Celdas de texto que no se pudieron convertir a número: {sum(no_leidos.values())}")
    st.write("Responsables detectados: " + ", ".join(sorted(validos(df["RESPONSABLE"]))))

    neg = {
        "Ahorro costo de venta": int((df["AHORRO_CV"] < 0).sum()),
        "Ahorro proforma": int((df["AHORRO_PROF"] < 0).sum()),
        "Ahorro negociación cierre": int((df["AHORRO_CIERRE"] < 0).sum()),
        "Fuera de MIX": int((df["FUERA_RAW"] < 0).sum()),
    }
    st.write("Filas con valores NEGATIVOS → " + " · ".join(f"{k}: {v}" for k, v in neg.items()))

    igual = (df["AHORRO_CV"] == df["AHORRO_PROF"]) & (df["AHORRO_CV"] != 0)
    st.write(f"Filas donde ahorro costo de venta = ahorro proforma (posible duplicado al sumar): "
             f"{int(igual.sum())} filas, {m(df.loc[igual, 'AHORRO_PROF'].sum())}")

    excede = df[(df["SIN_CV"] > df["GESTIONES"]) | (df["SIN_PROF"] > df["GESTIONES"])]
    st.write(f"Filas con más gestiones 'sin...' que gestiones totales: {len(excede)}")
    if len(excede):
        st.dataframe(excede[["FILA_EXCEL", "OBRA", "MES", "GESTIONES", "SIN_CV", "SIN_PROF"]],
                     hide_index=True, width="stretch")

    repetidas = df[(df["IMPORTE"] > 0) & df.duplicated(["GESTIONES", "IMPORTE"], keep=False)]
    st.write(f"Filas con gestiones e importe idénticos a otra fila (posible copiado): {len(repetidas)}")
    if len(repetidas):
        st.dataframe(repetidas.sort_values(["IMPORTE", "FILA_EXCEL"])[
            ["FILA_EXCEL", "OBRA", "MES", "GESTIONES", "IMPORTE"]], hide_index=True, width="stretch")

    dup_clave = df[df.duplicated(["OBRA", "MES"], keep=False)]
    st.write(f"Filas con obra y mes repetidos (se sumarían dos veces): {len(dup_clave)}")
    if len(dup_clave):
        st.dataframe(dup_clave[["FILA_EXCEL", "OBRA", "MES", "RESPONSABLE", "IMPORTE"]],
                     hide_index=True, width="stretch")
