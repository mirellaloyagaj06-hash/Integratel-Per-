
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------
st.set_page_config(page_title="Integratel | Inteligencia concursal", page_icon="📡", layout="wide")
ROOT = Path(__file__).parent
DEFAULT_XLSX = ROOT / "Movistar_Proceso_Concursal_Reestructuracion.xlsx"
PDF_PATH = ROOT / "Informe_Concursal_Integratel.pdf"

NAVY = "#102F50"
BLUE = "#176EAF"
CYAN = "#42B4E6"
LIME = "#A8CF38"
PALE = "#EEF8FC"
INK = "#18324D"
GREEN = "#16836A"
AMBER = "#D68A28"
RED = "#C94C55"

st.markdown("""
<style>
:root {--navy:#12345A;--blue:#2476B8;--ink:#203047;--muted:#63748A;--line:#DCE5EF;}
.stApp {background:#F3F6FA;color:var(--ink)!important;}
.block-container {max-width:1580px;padding-top:1.15rem;padding-bottom:2.5rem;}
header[data-testid="stHeader"] {background:rgba(243,246,250,.96);}
/* Evita texto blanco heredado del tema oscuro del navegador/Streamlit */
.stApp, .stApp p, .stApp label, .stApp [data-testid="stMarkdownContainer"],
.stApp [data-testid="stMarkdownContainer"] p, .stApp [data-testid="stMarkdownContainer"] li,
.stApp [data-testid="stMarkdownContainer"] h1, .stApp [data-testid="stMarkdownContainer"] h2,
.stApp [data-testid="stMarkdownContainer"] h3, .stApp [data-testid="stMarkdownContainer"] h4,
.stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stDataFrame"],
.stApp [data-testid="stTable"] {color:var(--ink)!important;}
.hero {background:linear-gradient(112deg,#102F50 0%,#155786 58%,#2479AE 100%);padding:22px 28px;border-radius:18px;color:#fff!important;margin-bottom:16px;box-shadow:0 9px 26px rgba(16,47,80,.14);}
.hero h1 {color:#fff!important;font-size:1.9rem;margin:0 0 6px 0;font-weight:780;}
.hero p {color:#E5F0FA!important;margin:0;font-size:.98rem;}
.hero .eyebrow {text-transform:uppercase;letter-spacing:.13em;font-size:.72rem;color:#C8E2F5!important;font-weight:750;margin-bottom:8px;}
.brandmark {width:190px;height:68px;flex:0 0 190px;border-radius:14px;background:#fff;display:flex;align-items:center;justify-content:center;padding:8px 12px;box-shadow:0 4px 12px rgba(0,0,0,.12);}
.brandmark img {display:block;width:100%;height:100%;object-fit:contain;}
.eyebrow {text-transform:uppercase;letter-spacing:.13em;font-size:.72rem;color:#BBD8EF;font-weight:700;margin-bottom:8px;}
.section-kicker {color:#58718A;text-transform:uppercase;letter-spacing:.12em;font-size:.72rem;font-weight:700;margin:6px 0;}
div[data-testid="stMetric"] {background:#fff;border:1px solid #DCE6F0;padding:17px 18px;border-radius:15px;box-shadow:0 4px 14px rgba(24,50,77,.055);min-height:112px;}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] p {color:#52677E!important;font-weight:700!important;}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] div {color:#123B65!important;font-weight:800!important;}
div[data-testid="stMetricDelta"] {font-weight:650;}
div[data-testid="stTabs"] button {font-weight:700;color:#29445F!important;}
.card {background:#fff;border:1px solid #DFE8F1;border-radius:15px;padding:18px 20px;margin-bottom:10px;color:#203047!important;box-shadow:0 3px 12px rgba(24,50,77,.035);}
.card h2,.card h3,.card b,.card p,.card li {color:#203047!important;}
.note {background:#EAF2FA;border-left:4px solid #347EB4;padding:12px 15px;border-radius:8px;color:#25445F!important;}
.warning {background:#FFF6E8;border-left:4px solid #D68A28;padding:12px 15px;border-radius:8px;color:#684A1C!important;}
.status {display:inline-block;background:#E5F4EF;color:#126A53!important;border-radius:30px;padding:7px 12px;font-size:.82rem;font-weight:750;}
.smallmuted {font-size:.82rem;color:#667F96!important;}
section[data-testid="stSidebar"] {background:#142B43;}
section[data-testid="stSidebar"] * {color:#F3F7FB!important;}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] {background:#1C3855;border-radius:12px;padding:8px;}
/* tablas legibles */
div[data-testid="stDataFrame"] [role="grid"], div[data-testid="stDataFrame"] canvas {color:#203047!important;}
/* Identidad Movistar: celeste, azul profundo y verde limón */
:root {--mov-blue:#0066A6;--mov-cyan:#35B6E8;--mov-lime:#A8CF38;}
.hero {background:linear-gradient(112deg,#005B96 0%,#087DB5 62%,#28A9D6 100%);}
.hero, .hero * {color:#FFFFFF!important;}
.hero p {color:#F1FAFF!important;}
.hero .eyebrow {color:#DDF5FF!important;}
.stApp [data-testid="stDataFrame"] {border:1px solid #D5E5EF!important;border-radius:12px!important;overflow:hidden;}
.stApp [data-testid="stDataFrame"] [role="columnheader"] {background:#E8F6FC!important;font-weight:750!important;}
.stApp [data-testid="stDataFrame"] [role="row"]:nth-child(even) {background:#F5FAFD!important;}
.stApp [data-testid="stDataFrame"] [role="row"]:hover {background:#EAF7FB!important;}
.stApp [data-testid="stSelectbox"] label {font-weight:700!important;color:#14547A!important;}
div[data-testid="stMetric"] {border-top:4px solid #35B6E8;}
/* espacio entre paneles y gráficos */
div[data-testid="stPlotlyChart"] {background:#fff;border:1px solid #D5E5EF;border-radius:14px;padding:8px 8px 2px;box-shadow:0 3px 12px rgba(24,50,77,.035);}
.timeline-wrap {position:relative;margin:10px 0 22px 8px;padding:4px 0 4px 30px;border-left:3px solid #35B6E8;}
.timeline-item {position:relative;background:#FFFFFF;border:1px solid #DCEAF2;border-radius:13px;padding:14px 17px;margin:0 0 13px 0;box-shadow:0 3px 12px rgba(13,73,105,.045);}
.timeline-item:before {content:"";position:absolute;left:-39px;top:19px;width:15px;height:15px;border-radius:50%;background:#35B6E8;border:4px solid #E8F6FC;box-shadow:0 0 0 2px #35B6E8;}
.timeline-item.done:before {background:#A8CF38;box-shadow:0 0 0 2px #A8CF38;}
.timeline-item.pending:before {background:#F0B64C;box-shadow:0 0 0 2px #F0B64C;}
.timeline-item.projected:before {background:#1685BD;box-shadow:0 0 0 2px #1685BD;}
.timeline-item.risk:before {background:#D65B68;box-shadow:0 0 0 2px #D65B68;}
.timeline-top {display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:5px;}
.timeline-date {font-size:.82rem;font-weight:750;color:#0879AD;}
.timeline-status {font-size:.72rem;font-weight:750;padding:4px 9px;border-radius:20px;background:#EAF5FA;color:#17658B;}
.timeline-title {font-weight:750;color:#163B59;font-size:1rem;margin:2px 0 5px;}
.timeline-desc {color:#526B7E;font-size:.88rem;line-height:1.45;}
.table-note {background:#F0F9FC;border-left:4px solid #35B6E8;padding:10px 13px;border-radius:8px;color:#28536B;font-size:.86rem;}

hr {border-color:#DCE5EF!important;}

/* Contraste reforzado: evita texto oscuro sobre superficies oscuras y viceversa. */
.stApp, .stApp [data-testid="stAppViewContainer"] {background:#F3F6FA!important;color:#203047!important;}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp p,.stApp li,.stApp label,
.stApp [data-testid="stMarkdownContainer"],.stApp [data-testid="stMarkdownContainer"] * {color:#203047;}
.hero h1,.hero p,.hero .eyebrow {color:inherit!important;}
.hero h1 {color:#FFFFFF!important;}
.stApp .hero h1,.stApp .hero p,.stApp .hero .eyebrow {color:#FFFFFF!important;}
.hero p {color:#E5F0FA!important;}
.hero .eyebrow {color:#C8E2F5!important;}
/* Sidebar: fondo azul oscuro con todos los textos y controles en alto contraste. */
section[data-testid="stSidebar"],section[data-testid="stSidebar"] > div {background:#142B43!important;}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] * {color:#F4F8FC!important;}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {color:#C6D5E4!important;}
section[data-testid="stSidebar"] [role="radiogroup"] label p {color:#F4F8FC!important;}
section[data-testid="stSidebar"] [role="radio"] div {color:#F4F8FC!important;}
/* Controles principales siempre claros, incluso si el navegador usa modo oscuro. */
.stApp input,.stApp textarea,.stApp [data-baseweb="select"] > div,
.stApp [data-baseweb="input"] > div,.stApp [data-testid="stSelectbox"] div {background:#FFFFFF!important;color:#203047!important;}
.stApp [data-baseweb="select"] *, .stApp [data-baseweb="input"] * {color:#203047!important;}
.stApp [role="listbox"],.stApp [role="option"] {background:#FFFFFF!important;color:#203047!important;}
.stApp [data-testid="stDataFrame"],.stApp [data-testid="stTable"] {background:#FFFFFF!important;color:#203047!important;}
.stApp [data-testid="stDataFrame"] *, .stApp [data-testid="stTable"] * {color:#203047!important;}
.stApp [data-testid="stMetricLabel"] *, .stApp [data-testid="stMetricValue"] * {color:#123B65!important;}
.stApp [data-testid="stPlotlyChart"] {background:#FFFFFF!important;}

</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# CARGA DEL EXCEL
# ------------------------------------------------------------
@st.cache_data
def load_table(path_or_buffer, sheet_name, header):
    return pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=header)

@st.cache_data
def get_sheet_names(path_or_buffer):
    x = pd.ExcelFile(path_or_buffer)
    names = x.sheet_names
    return names

with st.sidebar:
    st.markdown("## ◈ INTEGRATEL PERÚ")
    st.caption("Panel de inteligencia financiera y concursal")
    uploaded = st.file_uploader("Actualizar libro Excel", type=["xlsx"])
    source = uploaded if uploaded is not None else str(DEFAULT_XLSX)
    st.divider()
    page = st.radio("NAVEGACIÓN", [
        "Vista ejecutiva", "Crisis financiera", "Acreedores y poder de voto",
        "Junta y escenarios", "Ruta del procedimiento", "Plan propuesto",
        "Estados financieros", "Fuentes y metodología"
    ])
    st.divider()
    st.caption("Corte del informe: 01/11/2026")
    st.caption("Trabajo universitario · LGSC")

if not DEFAULT_XLSX.exists() and uploaded is None:
    st.error("No se encuentra el archivo Excel del proyecto.")
    st.stop()

try:
    sheets = get_sheet_names(source)
except Exception as e:
    st.error(f"No se pudo leer el Excel: {e}")
    st.stop()

def read(sheet, header=0):
    if sheet not in sheets:
        return pd.DataFrame()
    try:
        return load_table(source, sheet, header)
    except Exception:
        return pd.DataFrame()

def numeric(series):
    return pd.to_numeric(series, errors="coerce")

def money_m(value):
    if pd.isna(value): return "—"
    return f"S/ {value:,.1f} M"

def style_fig(fig, height=390):
    """Tema claro consistente: títulos, ejes, leyendas y etiquetas legibles."""
    fig.update_layout(
        template="plotly_white", height=height,
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
        font=dict(family="Arial, sans-serif", color="#203047", size=13),
        title=dict(font=dict(size=18, color="#12345A"), x=0.02, xanchor="left"),
        margin=dict(l=28, r=28, t=66, b=70),
        legend=dict(orientation="h", y=-0.24, x=0.02,
                    font=dict(color="#334B63", size=12),
                    bgcolor="rgba(255,255,255,0)"),
        hoverlabel=dict(bgcolor="#FFFFFF", font=dict(color="#203047", size=13)),
        uniformtext_minsize=11, uniformtext_mode="hide"
    )
    fig.update_xaxes(showgrid=False, linecolor="#DCE5EF", tickfont=dict(color="#40566D", size=11), title_font=dict(color="#40566D"))
    fig.update_yaxes(gridcolor="#E8EEF4", zerolinecolor="#CCD8E4", tickfont=dict(color="#40566D", size=11), title_font=dict(color="#40566D"))
    return fig

def section(title, subtitle=None):
    st.markdown(f"### {title}")
    if subtitle: st.caption(subtitle)

def callout(text, kind="note"):
    st.markdown(f'<div class="{kind}">{text}</div>', unsafe_allow_html=True)

def clean_table(data, keep_empty_columns=False):
    """Limpia marcadores vacíos sin alterar cifras reales ni eliminar ceros."""
    out=data.copy()
    out.columns=[str(c).strip() for c in out.columns]
    out=out.replace(to_replace=r"(?i)^\s*(none|null|nan|nat)\s*$",value=pd.NA,regex=True)
    out=out.replace([float("inf"),float("-inf")],pd.NA)
    if not keep_empty_columns:
        out=out.dropna(axis=1,how="all")
    return out

def show_table(data, height=None, hide_index=True, column_config=None):
    out=clean_table(data)
    # Las celdas sin dato se muestran como raya; los ceros se conservan.
    out=out.fillna("—")
    styler=out.style.set_properties(**{"text-align":"left","padding":"9px 10px"})
    styler=styler.set_table_styles([
        {"selector":"th","props":[("background-color","#E8F6FC"),("color","#075B88"),("font-weight","700"),("border-bottom","2px solid #35B6E8")]},
        {"selector":"td","props":[("border-bottom","1px solid #E3EDF3"),("color","#173B56")]},
        {"selector":"tbody tr:nth-child(even)","props":[("background-color","#F5FAFD")]},
    ])
    kwargs={"use_container_width":True,"hide_index":hide_index}
    if height is not None: kwargs["height"]=height
    if column_config is not None: kwargs["column_config"]=column_config
    st.dataframe(styler,**kwargs)

# Excel layout: header rows based on the supplied workbook.
creditors = read("Acreedores", 4)
creditors.columns = [str(c).strip() for c in creditors.columns]
creditors = creditors[creditors.get("Acreedor", pd.Series(dtype=str)).notna()].copy()
creditors = creditors[creditors["Acreedor"].astype(str).str.lower() != "total"]
if not creditors.empty:
    creditors["Monto (S/ M)"] = numeric(creditors["Monto (S/ M)"])
    creditors["% del total"] = numeric(creditors["% del total"])

vote = read("Junta_Votacion", 14)
proc = read("Procedimiento", 3)
plan = read("Plan_Reestructuracion", 11)
ind = read("Indicadores", 2)
if not ind.empty:
    ind = ind.rename(columns={ind.columns[0]:"Indicador"})
    year_cols = [c for c in ind.columns if str(c).replace(".0","").isdigit()]
    for c in year_cols: ind[c] = numeric(ind[c])
esf = read("ESF", 2)
er = read("ER", 2)
efe = read("EFE", 2)
cron = read("Cronologia", 2)
sup = read("Supuestos", 3)

# Logotipo oficial proporcionado por la usuaria (fondo eliminado)
import base64
_logo_path = ROOT / "movistar_logo.png"
_logo_b64 = base64.b64encode(_logo_path.read_bytes()).decode("ascii") if _logo_path.exists() else ""

# Header
st.markdown(f"""
<div class="hero">
  <div style="display:flex;align-items:center;gap:16px;">
    <div class="brandmark" aria-label="Logotipo Movistar">
      <img src="data:image/png;base64,{_logo_b64}" alt="Movistar" />
    </div>
    <div>
      <div class="eyebrow">INTELIGENCIA FINANCIERA Y CONCURSAL · PERÚ · LEY 27809</div>
      <h1>Telefónica del Perú S.A.A. · Integratel Perú</h1>
      <p>Procedimiento concursal ordinario · Diagnóstico financiero · Análisis de acreedores y escenarios</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# Global source / status strip
c1,c2,c3 = st.columns([1.1,1.5,1.2])
with c1: st.markdown('<span class="status">En reconocimiento de créditos</span>', unsafe_allow_html=True)
with c2: st.markdown("**Junta de Acreedores:** no instalada a la fecha de corte del informe", unsafe_allow_html=True)
with c3: st.markdown("**Corte documental:** 01 de noviembre de 2026", unsafe_allow_html=True)

# ------------------------------------------------------------
# VISTA EJECUTIVA
# ------------------------------------------------------------
if page == "Vista ejecutiva":
    section("Panorama del procedimiento", "Indicadores clave del caso según el Excel y el informe proporcionados.")
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Créditos reconocidos", "S/ 5.154 M", help="Total reportado en el informe; sujeto a reconsideraciones pendientes.")
    m2.metric("Crédito SUNAT", "S/ 1.209,6 M", "23,47 % referencial")
    m3.metric("Ingresos 2024", "S/ 6.020 M", "−7,0 % vs. 2023")
    m4.metric("Junta de Acreedores", "No instalada", "Sin fecha definida en el informe")
    callout("La situación real del expediente debe distinguirse del escenario académico: al corte del informe, el procedimiento seguía en reconocimiento de créditos. La reestructuración y el plan del Capítulo VII son un supuesto de trabajo, no acuerdos aprobados.", "warning")

    if not ind.empty and year_cols:
        section("Evolución de la situación financiera", "Indicadores históricos del libro; importes expresados en miles de S/.")
        executive_metrics = ["Ingresos ordinarios", "Ganancia (pérdida) neta", "Total activos", "Total pasivos", "Patrimonio total"]
        ed = ind[ind["Indicador"].astype(str).isin(executive_metrics)].copy()
        if not ed.empty:
            years_exec = sorted(year_cols, key=lambda x: int(float(x)))
            ed = ed.melt(id_vars=["Indicador"], value_vars=years_exec, var_name="Año", value_name="Valor")
            ed["Año"] = ed["Año"].astype(str).str.replace(".0", "", regex=False)
            ef = px.line(ed, x="Año", y="Valor", color="Indicador", markers=True,
                         title="Tendencia de indicadores principales", color_discrete_sequence=["#2476B8", "#E35D6A", "#20A080", "#E0A13A", "#6B7FA6"])
            ef.update_traces(line=dict(width=3), marker=dict(size=7))
            st.plotly_chart(style_fig(ef, 390), use_container_width=True)

    left,right = st.columns([1.05,1.35])
    with left:
        section("Etapa actual")
        st.markdown("""
        <div class="card">
        <div class="smallmuted">PROCEDIMIENTO CONCURSAL ORDINARIO</div>
        <h2 style="color:#176EAF;margin:.3rem 0">Reconocimiento de créditos</h2>
        <p style="margin:0">Reconsideraciones y apelaciones pendientes; padrón definitivo de acreedores aún por cerrar.</p>
        <hr style="border:0;border-top:1px solid #E5ECF3">
        <b>Lo que todavía no ha ocurrido</b>
        <ul>
        <li>Instalación formal de la Junta.</li>
        <li>Decisión sobre reestructuración o liquidación.</li>
        <li>Aprobación de un plan definitivo.</li>
        <li>Designación de liquidador.</li>
        </ul></div>
        """, unsafe_allow_html=True)
    with right:
        section("Ruta general del caso")
        steps = pd.DataFrame({
            "Etapa":["Solicitud","Inicio PCO","Publicación","Reconocimiento","Junta","Decisión","Plan"],
            "Estado":["Cumplido","Cumplido","Cumplido","En trámite","Pendiente","Pendiente","Supuesto"]
        })
        colors={"Cumplido":GREEN,"En trámite":AMBER,"Pendiente":"#A9B7C5","Supuesto":BLUE}
        fig=go.Figure()
        for _,r in steps.iterrows():
            fig.add_trace(go.Bar(x=[1],y=[r["Etapa"]],orientation="h",marker_color=colors[r["Estado"]],name=r["Estado"],showlegend=False,
                                 text=r["Estado"],textposition="inside",hovertemplate="%{y}<br>"+r["Estado"]+"<extra></extra>"))
        fig.update_layout(barmode="overlay",height=340,margin=dict(l=8,r=8,t=15,b=8),paper_bgcolor="white",plot_bgcolor="white",
                          xaxis=dict(visible=False,range=[0,1]),yaxis=dict(autorange="reversed",tickfont=dict(color=INK)))
        st.plotly_chart(style_fig(fig),use_container_width=True)
    st.divider()
    section("Lectura financiera rápida")
    if not ind.empty and year_cols:
        years=sorted(year_cols,key=lambda x:int(float(x)))
        latest=years[-1]
        wanted=["Ingresos ordinarios","Ganancia (pérdida) neta","Total activos","Total pasivos","Patrimonio total"]
        rows=ind[ind["Indicador"].astype(str).isin(wanted)]
        show=rows[["Indicador",latest]].copy()
        show.columns=["Indicador",f"{latest} (miles de S/)"]
        show_table(show)
    st.caption("Las cifras financieras de 2025 están identificadas en el libro como consolidadas; los años previos corresponden a estados individuales. Evita compararlas sin advertirlo.")

# ------------------------------------------------------------
# CRISIS FINANCIERA
# ------------------------------------------------------------
elif page == "Crisis financiera":
    section("Diagnóstico financiero", "Series históricas tomadas de la hoja Indicadores (2020–2025).")
    if ind.empty or not year_cols:
        st.warning("No se encontró una tabla Indicadores con años.")
    else:
        years=sorted(year_cols,key=lambda x:int(float(x)))
        metric_options=["Ingresos ordinarios","Ganancia (pérdida) operativa","Ganancia (pérdida) neta","Total activos","Total pasivos","Patrimonio total","Flujo de operación","Capital de trabajo (miles S/)"]
        selected=st.multiselect("Selecciona indicadores",metric_options,default=metric_options[:3])
        subset=ind[ind["Indicador"].astype(str).isin(selected)]
        melted=subset.melt(id_vars=["Indicador"],value_vars=years,var_name="Año",value_name="Valor")
        melted["Año"]=melted["Año"].astype(str).str.replace(".0","",regex=False)
        fig=px.line(melted,x="Año",y="Valor",color="Indicador",markers=True,title="Evolución de indicadores (miles de S/)")
        st.plotly_chart(style_fig(fig,420),use_container_width=True)
        st.divider()
        left,right=st.columns(2)
        with left:
            section("Ingresos y pérdidas")
            wanted=["Ingresos ordinarios","Ganancia (pérdida) neta"]
            d=ind[ind["Indicador"].astype(str).isin(wanted)].melt(id_vars=["Indicador"],value_vars=years,var_name="Año",value_name="Valor")
            fig=px.bar(d,x="Año",y="Valor",color="Indicador",barmode="group",title="Ingresos frente a resultado neto")
            st.plotly_chart(style_fig(fig,350),use_container_width=True)
        with right:
            section("Solvencia patrimonial")
            wanted=["Total activos","Total pasivos","Patrimonio total"]
            d=ind[ind["Indicador"].astype(str).isin(wanted)].melt(id_vars=["Indicador"],value_vars=years,var_name="Año",value_name="Valor")
            fig=px.line(d,x="Año",y="Valor",color="Indicador",markers=True,title="Activos, pasivos y patrimonio")
            st.plotly_chart(style_fig(fig,350),use_container_width=True)
        st.caption("Unidades: miles de S/. La hoja Indicadores contiene ratios calculados y controles de integridad; revisa la metodología antes de interpretar cada ratio.")

# ------------------------------------------------------------
# ACREEDORES
# ------------------------------------------------------------
elif page == "Acreedores y poder de voto":
    section("Quiénes concentran el poder de decisión", "Distribución referencial calculada sobre S/ 5.154 millones. Las cifras estimadas se identifican expresamente.")
    if creditors.empty:
        st.warning("No se encontró la tabla de acreedores.")
    else:
        creditors["% voto"] = creditors["% del total"]*100
        order_map={"1":1,"2":2,"3":3,"4":4,"5":5}
        creditors["Orden_num"] = pd.to_numeric(creditors["Orden de prelación (art. 42)"],errors="coerce")
        creditors=creditors.sort_values(["Orden_num","Monto (S/ M)"],ascending=[True,False])
        a,b,c=st.columns(3)
        a.metric("Total reconocido", "S/ 5.154 M")
        b.metric("Mayor bloque", "Integra Tec", "40,00 % referencial")
        c.metric("Mayoría calificada", "> 66,6 %", "Art. 53.1 LGSC")
        left,right=st.columns([1.1,1])
        with left:
            section("Poder de voto por acreedor")
            vote_order = creditors.sort_values("% voto", ascending=True)
            fig=px.bar(vote_order,x="% voto",y="Acreedor",orientation="h",color="Acreedor",
                       text=vote_order["% voto"].map(lambda x:f"{x:.2f}%"),
                       hover_data=["Monto (S/ M)","Tipo de dato","Vinculado al deudor"],
                       title="Participación referencial en los votos",
                       color_discrete_sequence=["#2476B8", "#159B91", "#E7A23B", "#D85B67", "#7186A6", "#69A6C8"])
            fig.update_traces(textposition="outside", textfont=dict(color="#203047",size=12), cliponaxis=False)
            fig.update_layout(showlegend=False, xaxis_title="Participación (%)", yaxis_title="", margin=dict(l=20,r=45,t=65,b=45))
            st.plotly_chart(style_fig(fig,440),use_container_width=True)
        with right:
            section("Composición del pasivo")
            fig=px.pie(creditors,names="Acreedor",values="Monto (S/ M)",hole=.58,
                       hover_data=["% voto","Tipo de dato"],title="Composición de los créditos por acreedor",
                       color_discrete_sequence=["#2476B8", "#159B91", "#E7A23B", "#D85B67", "#7186A6", "#69A6C8"])
            fig.update_traces(textinfo="percent",textposition="inside",insidetextfont=dict(color="#FFFFFF",size=13),
                              marker=dict(line=dict(color="#FFFFFF",width=2)),
                              hovertemplate="<b>%{label}</b><br>Crédito: S/ %{value:,.2f} M<br>Participación: %{percent}<extra></extra>")
            fig.update_layout(showlegend=True,legend=dict(orientation="v",y=0.5,x=1.02,font=dict(color="#334B63",size=11)),
                              margin=dict(l=10,r=125,t=65,b=20))
            st.plotly_chart(style_fig(fig,440),use_container_width=True)
        section("Orden de prelación y naturaleza de la cifra", "Detalle depurado de acreedores; los campos sin información no ocupan espacio.")
        preferred=["Orden de prelación (art. 42)","Acreedor","Monto (S/ M)","% voto","Bloque de voto","Tipo de dato","Vinculado al deudor"]
        view=creditors[[c for c in preferred if c in creditors.columns]].copy()
        if "% voto" in view: view["% voto"]=pd.to_numeric(view["% voto"],errors="coerce").map(lambda x:f"{x:.2f}%" if pd.notna(x) else "—")
        if "Monto (S/ M)" in view: view["Monto (S/ M)"]=pd.to_numeric(view["Monto (S/ M)"],errors="coerce").map(lambda x:f"{x:,.2f}" if pd.notna(x) else "—")
        if "Bloque de voto" in view: view["Bloque de voto"]=pd.to_numeric(view["Bloque de voto"],errors="coerce").map(lambda x:f"{x:,.2f}" if pd.notna(x) else "—")
        show_table(view)
        callout("Advertencia: SUNAT figura con crédito reconocido reportado; Integra Tec está estimado y bonistas está marcado como supuesto en el Excel. Los créditos laborales y proveedores se presentan por separado; el monto laboral es un supuesto. El padrón puede variar por reconsideraciones, apelaciones y créditos contingentes.", "warning")
        st.markdown("**Lectura:** Integra Tec es el mayor bloque individual, pero no alcanza por sí sola la mayoría calificada. El poder efectivo depende también de asistencia, reconocimiento definitivo y reglas de convocatoria.")

# ------------------------------------------------------------
# JUNTA / ESCENARIOS
# ------------------------------------------------------------
elif page == "Junta y escenarios":
    section("Simulador de votación de la Junta", "Escenarios del Capítulo V del informe. No representa una votación que ya haya ocurrido.")
    total=5154.0
    majority=0.666
    if not vote.empty:
        scenario_rows=vote.iloc[0:4].copy()
        # use fixed known columns from sheet header at row 14
        scenario_rows.columns=[str(c).strip() for c in scenario_rows.columns]
        scenario_names=["A","B","C","D"]
        chosen=st.selectbox("Selecciona escenario",scenario_names,index=0,format_func=lambda x: {
            "A":"A · Integra Tec + bonistas",
            "B":"B · Integra Tec + SUNAT",
            "C":"C · SUNAT + bonistas",
            "D":"D · Integra Tec + SUNAT + laborales + proveedores"}[x])
        row=scenario_rows.iloc[scenario_names.index(chosen)]
        yes=float(row.get("% a favor de reestructuración",0))*100
        no=float(row.get("% a favor de liquidación",0))*100
        x,y,z=st.columns(3)
        x.metric("A favor de reestructurar",f"{yes:.2f}%")
        y.metric("A favor de liquidar",f"{no:.2f}%")
        z.metric("Umbral legal","> 66,6 %")
        fig=go.Figure()
        fig.add_trace(go.Bar(name="A favor",x=["Votación"],y=[yes],marker_color=GREEN))
        fig.add_trace(go.Bar(name="En contra",x=["Votación"],y=[no],marker_color=RED))
        fig.add_hline(y=66.6,line_dash="dash",line_color=AMBER,annotation_text="Mayoría calificada > 66,6%")
        fig.update_layout(barmode="stack",yaxis=dict(range=[0,100],ticksuffix="%"),height=360,
                          paper_bgcolor="white",plot_bgcolor="white",margin=dict(l=25,r=25,t=55,b=35),
                          font=dict(color="#203047",size=13),legend=dict(font=dict(color="#334B63")))
        fig.update_xaxes(tickfont=dict(color="#40566D")); fig.update_yaxes(tickfont=dict(color="#40566D"))
        st.plotly_chart(style_fig(fig),use_container_width=True)
        st.markdown(f"**Resultado que consigna el Excel:** {row.get('Resultado jurídico','No consignado')}")
        st.caption(str(row.get("Descripción","")))
        st.divider()
        section("Prueba interactiva de combinaciones")
        st.write("Marca los bloques que votarían a favor. Los porcentajes se calculan directamente con los montos de la hoja Acreedores.")
        total_creditos=float(creditors["Monto (S/ M)"].sum()) if not creditors.empty else 0.0
        weights={}
        for block in ["Integra Tec","SUNAT","Bonistas","Créditos laborales","Proveedores y otros"]:
            monto=float(creditors.loc[creditors["Acreedor"].astype(str).str.strip()==block,"Monto (S/ M)"].sum()) if not creditors.empty else 0.0
            weights[block]=(monto/total_creditos*100) if total_creditos else 0.0
        cols=st.columns(5)
        selected_blocks=[]
        default_favorables=["Integra Tec","Bonistas"] if chosen=="A" else []
        for col,(name,pct) in zip(cols,weights.items()):
            with col:
                if st.checkbox(f"{name} · {pct:.2f}%",value=(name in default_favorables),key="v_"+name):
                    selected_blocks.append(name)
        yes_custom=sum(weights[n] for n in selected_blocks)
        st.metric("Voto favorable simulado",f"{yes_custom:.2f}%",f"{'Alcanza' if yes_custom>66.6 else 'No alcanza'} la mayoría")
        st.progress(min(yes_custom/100,1.0))
        st.caption("La simulación aplica el umbral de primera convocatoria mostrado en el informe. La segunda convocatoria y la composición final de créditos pueden cambiar la base de cálculo.")
    else:
        st.warning("No se encontró la hoja Junta_Votacion.")

# ------------------------------------------------------------
# PROCEDIMIENTO
# ------------------------------------------------------------
elif page == "Ruta del procedimiento":
    section("Línea de tiempo del procedimiento concursal", "Secuencia cronológica basada en las fechas y descripciones del informe y del libro Excel.")
    if proc.empty:
        st.warning("No se encontró la hoja Procedimiento.")
    else:
        import html
        timeline=proc.dropna(subset=["Etapa"]).copy()
        timeline["N.º"]=numeric(timeline["N.º"])
        timeline=timeline.sort_values("N.º")
        status_filter=st.multiselect("Filtrar etapas",sorted(timeline["Estado"].dropna().astype(str).unique()),
                                     default=sorted(timeline["Estado"].dropna().astype(str).unique()))
        timeline=timeline[timeline["Estado"].astype(str).isin(status_filter)]
        def fmt_date(value):
            if pd.isna(value): return "Fecha no indicada"
            if hasattr(value,"strftime"): return value.strftime("%d/%m/%Y")
            return str(value)
        def status_class(value):
            v=str(value).lower()
            if "cumplido" in v: return "done"
            if "trámite" in v: return "pending"
            if "contingente" in v: return "risk"
            return "projected"
        cards=[]
        for _,r in timeline.iterrows():
            status=html.escape(str(r.get("Estado","Sin estado")))
            title=html.escape(str(r.get("Etapa","")))
            date=html.escape(fmt_date(r.get("Fecha (real o proyectada)")))
            desc=html.escape(str(r.get("Qué ocurre","")))
            legal=html.escape(str(r.get("Base legal","")))
            legal_html=f'<div class="timeline-desc"><b>Base legal:</b> {legal}</div>' if legal and legal!="—" else ""
            cards.append(f'<div class="timeline-item {status_class(status)}"><div class="timeline-top"><span class="timeline-date">{date}</span><span class="timeline-status">{status}</span></div><div class="timeline-title">{title}</div><div class="timeline-desc">{desc}</div>{legal_html}</div>')
        st.markdown('<div class="timeline-wrap">'+"".join(cards)+'</div>',unsafe_allow_html=True)
        callout("Las etapas con fechas futuras son proyecciones académicas del archivo, no fechas oficiales confirmadas. El informe señala que, al corte documental, el procedimiento seguía en reconocimiento de créditos y la Junta aún no estaba instalada.","warning")
        with st.expander("Consultar detalle de etapas y responsables"):
            cols=["N.º","Etapa","Estado","Fecha (real o proyectada)","Plazo legal","Base legal","Responsable"]
            show_table(timeline[[c for c in cols if c in timeline.columns]])

# ------------------------------------------------------------
# PLAN
# ------------------------------------------------------------
elif page == "Plan propuesto":
    section("Plan de reestructuración patrimonial · escenario académico", "Capítulo VII: supuesto de que la Junta se instala y aprueba la continuidad.")
    callout("Este plan no está aprobado ni vigente. Sus importes, porcentajes de capitalización y cronograma son ilustrativos y dependen de decisiones futuras de la Junta.", "warning")
    plan_df=read("Plan_Reestructuracion",11)
    if not plan_df.empty:
        # table begins row 11, rows 12-16 creditors, row 17 total
        plan_df=plan_df.dropna(subset=["Acreedor"],how="any")
        plan_df=plan_df[plan_df["Acreedor"].astype(str).str.lower()!="total"].copy()
        for c in ["Crédito (S/ M)","Capitalizado en acciones (S/ M)","Efectivo nominal (S/ M)","Cuota anual (S/ M)"]:
            if c in plan_df: plan_df[c]=numeric(plan_df[c])
        a,b,c=st.columns(3)
        a.metric("Plazo supuesto","15 años","3 de gracia + 12 de pago")
        b.metric("Capitalización propuesta","60%","Créditos de 5.º orden")
        c.metric("Efectivo nominal","S/ 2.847,36 M","Sin intereses en el modelo")
        left,right=st.columns([1.1,1])
        with left:
            fig=px.bar(plan_df,x="Acreedor",y=["Capitalizado en acciones (S/ M)","Efectivo nominal (S/ M)"],
                       barmode="stack",title="Tratamiento propuesto por acreedor")
            st.plotly_chart(style_fig(fig,390),use_container_width=True)
        with right:
            fig=px.pie(plan_df,names="Acreedor",values="Efectivo nominal (S/ M)",hole=.6,title="Distribución del pago en efectivo")
            st.plotly_chart(style_fig(fig,390),use_container_width=True)
        plan_view=clean_table(plan_df)
        for col in ["Crédito (S/ M)","Capitalizado en acciones (S/ M)","Efectivo nominal (S/ M)","Cuota anual (S/ M)"]:
            if col in plan_view:
                plan_view[col]=pd.to_numeric(plan_view[col],errors="coerce").map(lambda x:f"{x:,.2f}" if pd.notna(x) else "—")
        if "% capitalizado" in plan_view:
            plan_view["% capitalizado"]=pd.to_numeric(plan_view["% capitalizado"],errors="coerce").map(lambda x:f"{x:.1%}" if pd.notna(x) and x<=1 else (f"{x:.1f}%" if pd.notna(x) else "—"))
        show_table(plan_view)
        schedule=read("Plan_Reestructuracion",27)
        if not schedule.empty and "Año" in schedule:
            schedule=schedule.dropna(subset=["Año"])
            schedule["Año"]=numeric(schedule["Año"])
            cols=[c for c in ["Laborales","SUNAT","Bonistas","Integra Tec","Proveedores y otros"] if c in schedule.columns]
            for col in cols: schedule[col]=numeric(schedule[col])
            fig=px.bar(schedule,x="Año",y=cols,title="Cronograma indicativo de pagos por año")
            st.plotly_chart(style_fig(fig,420),use_container_width=True)
            show_table(schedule)
        cap=read("Plan_Reestructuracion",52)
        if not cap.empty:
            st.markdown("**Capacidad de pago referencial**")
            show_table(cap)
    else:
        st.warning("No se encontró la tabla del plan propuesto.")

# ------------------------------------------------------------
# ESTADOS FINANCIEROS
# ------------------------------------------------------------
elif page == "Estados financieros":
    section("Estados financieros comparativos", "Consulta por estado y año con formato financiero, jerarquía visual y cifras tomadas directamente del Excel.")
    tab=st.selectbox("Estado financiero",["ESF","ER","EFE"],format_func=lambda x:{"ESF":"Estado de situación financiera","ER":"Estado de resultados","EFE":"Estado de flujos de efectivo"}[x])
    df=read(tab,2)
      if not df.empty:
    account_col=next((c for c in df.columns if str(c).strip().lower()=="cuenta"),df.columns)
    years=[c for c in df.columns if str(c).replace(".0","").isdigit()]
    if years:
        years=sorted(years,key=lambda x:int(float(x)))
        chosen_year=st.selectbox("Año del estado financiero",years,index=len(years)-1,format_func=lambda x:str(int(float(x))))
        view=df[[account_col,chosen_year]].copy()
        view.columns=["Cuenta",f"Importe · {int(float(chosen_year))} (miles de S/)"]
        view=view.dropna(subset=["Cuenta"])
        view=view[~view["Cuenta"].astype(str).str.strip().str.lower().isin(["none","nan","total control"])]
        view[f"Importe · {int(float(chosen_year))} (miles de S/)"]=pd.to_numeric(view[f"Importe · {int(float(chosen_year))} (miles de S/)"])
        view=view.dropna(subset=[f"Importe · {int(float(chosen_year))} (miles de S/)"])
        positive=int((view.iloc[:,1]>0).sum()); negative=int((view.iloc[:,1]<0).sum())
        
        a,b,c=st.columns(3)
        a.metric("Cuentas mostradas",f"{len(view):,}")
        b.metric("Saldos positivos",f"{positive:,}")
        c.metric("Saldos negativos",f"{negative:,}")
        
        col_name = view.columns
        view[col_name] = view[col_name].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "-").astype(str)
        
        st.markdown('<div class="table-note">Importes expresados miles de S/. Los valores se conservan según el libro Excel; las c')
        st.dataframe(view.style.set_properties(**{"padding":"9px 12px"}).set_table_styles([
            {"selector":"th","props":[("background-color","#E8F6FC"),("color","#075B88"),("font-weight","750"),("border-bottom","3px solid #075B88")]},
            {"selector":"td","props":[("border-bottom","1px solid #E3EDF3"),("color","#173B56")]},
            {"selector":"tbody tr:nth-child(even)","props":[("background-color","#F5FAFD")]},
        ]),use_container_width=True,hide_index=True,height=620)
        st.download_button("Descargar vista CSV",view.to_csv(index=False).encode("utf-8-sig"),f"{tab}_{chosen_year}.csv","text/csv")
    else:
        show_table(df)

# ------------------------------------------------------------
# FUENTES / METODOLOGÍA
# ------------------------------------------------------------
elif page == "Fuentes y metodología":
    section("Trazabilidad y límites de la información")
    st.markdown("""
    **Fuentes incorporadas al proyecto**
    - Libro Excel proporcionado: hojas Resumen, Supuestos, Empresa, Cronologia, Procedimiento, Acreedores, Junta_Votacion, Plan_Reestructuracion, Base_Datos, Indicadores, ESF, ER y EFE.
    - Informe académico PDF: *Análisis del proceso concursal y la reestructuración empresarial de Telefónica del Perú S.A.A. bajo la Ley General del Sistema Concursal*.
    - Referencias normativas y periodísticas citadas en el informe: deben consultarse en sus publicaciones originales antes de una entrega final.

    **Criterios de presentación**
    - “Dato”: cifra reportada en la fuente del trabajo.
    - “Estimado”, “Supuesto” o “Cálculo”: no equivale a crédito firme reconocido por Indecopi.
    - “Proyectado”: etapa o fecha futura modelada, no hecho ocurrido.
    - El poder de voto se calcula en función del monto de créditos reconocidos, sujeto al padrón definitivo y a las reglas de asistencia y convocatoria.
    - La Junta no se encontraba instalada al corte del informe; por tanto, no existe decisión definitiva sobre reestructuración o liquidación ni nombramiento de liquidador.
    - Las cifras 2025 están marcadas como consolidadas; 2020–2024 como individuales en el libro de trabajo.
    """)
    if PDF_PATH.exists():
        with open(PDF_PATH,"rb") as f:
            st.download_button("Descargar informe PDF de referencia",f.read(),"Informe_Concursal_Integratel.pdf","application/pdf")
    st.markdown("**Hojas del libro disponibles:**")
    st.write(" · ".join(sheets))

st.divider()
st.caption("Aplicación académica · Elaborada con el Excel y el PDF entregados. Verificar siempre el expediente y las fuentes oficiales antes de afirmar cambios en el estado del procedimiento.")
