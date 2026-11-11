
from pathlib import Path
import html
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Integratel | Panel financiero y concursal", page_icon="📡", layout="wide")
ROOT = Path(__file__).parent
XLSX_PATH = ROOT / "Movistar_Plan_de_Reestructuracion_corregido.xlsx"
LOGO_PATH = ROOT / "assets" / "movistar_logo.png"
NAVY, BLUE, CYAN, LIME = "#005B96", "#087DB5", "#35B6E8", "#A8CF38"
INK, MUTED, PALE, GREEN, AMBER, RED = "#18324D", "#60758A", "#EEF8FC", "#16836A", "#D68A28", "#C94C55"

st.markdown("""
<style>
.stApp{background:#F3F7FA;color:#18324D}
.block-container{max-width:1600px;padding-top:1.2rem;padding-bottom:2.5rem}
header[data-testid="stHeader"]{background:rgba(243,247,250,.96)}
.stApp p,.stApp label,.stApp [data-testid="stMarkdownContainer"],.stApp h1,.stApp h2,.stApp h3,.stApp h4,
.stApp [data-testid="stCaptionContainer"],.stApp [data-testid="stDataFrame"]{color:#18324D}
.hero{background:linear-gradient(112deg,#005B96 0%,#087DB5 62%,#28A9D6 100%);padding:22px 28px;border-radius:18px;color:#fff!important;margin-bottom:18px;box-shadow:0 9px 26px rgba(16,47,80,.14)}
.hero h1,.hero p,.hero span{color:#fff!important;margin:0}
.hero h1{font-size:1.85rem;font-weight:800;margin-bottom:5px}
.hero p{font-size:.96rem;color:#EAF8FF!important}
.hero-grid{display:flex;align-items:center;gap:22px}
.hero-logo{width:175px;height:66px;background:#fff;border-radius:13px;display:flex;align-items:center;justify-content:center;padding:8px;flex:0 0 175px}
.hero-logo img{max-width:100%;max-height:100%;object-fit:contain}
.kicker{font-size:.72rem;text-transform:uppercase;letter-spacing:.14em;font-weight:800;color:#DDF5FF!important;margin-bottom:7px!important}
div[data-testid="stMetric"]{background:#fff;border:1px solid #DCE8F0;padding:15px 17px;border-radius:15px;box-shadow:0 4px 14px rgba(24,50,77,.05);min-height:104px}
div[data-testid="stMetricLabel"],div[data-testid="stMetricLabel"] p{color:#526B80!important;font-weight:700!important}
div[data-testid="stMetricValue"],div[data-testid="stMetricValue"] div{color:#075B91!important;font-weight:800!important}
.card{background:#fff;border:1px solid #DCE8F0;border-radius:15px;padding:17px 19px;margin:8px 0;color:#18324D}
.note{background:#EAF6FB;border-left:4px solid #35B6E8;padding:12px 15px;border-radius:8px;color:#23465F}
.warning{background:#FFF6E8;border-left:4px solid #D68A28;padding:12px 15px;border-radius:8px;color:#684A1C}
section[data-testid="stSidebar"]{background:#142F49}
section[data-testid="stSidebar"] *{color:#F4F8FB!important}
section[data-testid="stSidebar"] [data-testid="stRadio"] label{padding:4px 0}
.stApp [data-testid="stDataFrame"]{border:1px solid #D5E5EF;border-radius:12px;overflow:hidden}
.stApp [data-testid="stSelectbox"] label,.stApp [data-testid="stMultiSelect"] label{font-weight:750;color:#14547A!important}
div[data-testid="stTabs"] button{font-weight:700;color:#214A67!important}
.timeline{position:relative;margin:12px 0 8px 14px;padding-left:28px;border-left:3px solid #35B6E8}
.timeline-item{position:relative;background:#fff;border:1px solid #DCE8F0;border-radius:13px;padding:14px 17px;margin:0 0 13px 0;box-shadow:0 3px 10px rgba(24,50,77,.04)}
.timeline-item:before{content:"";position:absolute;left:-38px;top:20px;width:16px;height:16px;border-radius:50%;background:#35B6E8;border:3px solid white;box-shadow:0 0 0 2px #35B6E8}
.timeline-item.done:before{background:#A8CF38;box-shadow:0 0 0 2px #A8CF38}
.timeline-item.pending:before{background:#D68A28;box-shadow:0 0 0 2px #D68A28}
.timeline-item.risk:before{background:#C94C55;box-shadow:0 0 0 2px #C94C55}
.t-date{font-size:.82rem;font-weight:800;color:#087DB5}
.t-status{float:right;border-radius:30px;padding:4px 9px;font-size:.75rem;font-weight:800;background:#EAF6FB;color:#075B91}
.t-title{font-size:1.02rem;font-weight:800;margin:5px 0;color:#18324D}
.t-desc{font-size:.9rem;color:#526B80}
</style>
""", unsafe_allow_html=True)

def fmt_num(v, decimals=2):
    try:
        if pd.isna(v): return "—"
        if isinstance(v, str) and v.strip() in ("—","n.a.","None","nan"): return v
        return f"{float(v):,.{decimals}f}"
    except Exception:
        return str(v)

def fmt_pct(v, decimals=2):
    try:
        if pd.isna(v): return "—"
        return f"{float(v)*100:.{decimals}f}%"
    except Exception: return str(v)

def load_sheet(name, header=0):
    try:
        df=pd.read_excel(XLSX_PATH, sheet_name=name, header=header)
        df=df.loc[:,~df.columns.astype(str).str.match(r"^Unnamed")]
        df.columns=[str(c).strip() for c in df.columns]
        df=df.dropna(how="all")
        return df
    except Exception:
        return pd.DataFrame()

def clean_for_display(df):
    if df.empty: return df
    out=df.copy()
    out=out.loc[:,[c for c in out.columns if not out[c].isna().all()]]
    out=out.dropna(how="all")
    for c in out.columns:
        if out[c].dtype=="object":
            out[c]=out[c].map(lambda x:"—" if pd.isna(x) or str(x).strip().lower() in ("none","nan") else x)
    return out

def table(df, height=420, key=None):
    df=clean_for_display(df)
    if df.empty:
        st.info("No hay registros para mostrar en esta sección.")
        return
    st.dataframe(df,use_container_width=True,hide_index=True,height=height)

def section(title, subtitle=None):
    st.markdown(f"### {title}")
    if subtitle: st.caption(subtitle)

def note(text, warning=False):
    cls="warning" if warning else "note"
    st.markdown(f'<div class="{cls}">{html.escape(str(text))}</div>',unsafe_allow_html=True)

def chart_layout(fig,height=380):
    fig.update_layout(height=height,paper_bgcolor="white",plot_bgcolor="white",
                      font=dict(color=INK,size=12),margin=dict(l=25,r=25,t=55,b=35),
                      legend=dict(font=dict(color=INK)))
    fig.update_xaxes(tickfont=dict(color=INK),gridcolor="#E7EEF3")
    fig.update_yaxes(tickfont=dict(color=INK),gridcolor="#E7EEF3")
    return fig

def get_row_value(df, key_col, value_col, key):
    if df.empty or key_col not in df or value_col not in df: return None
    found=df[df[key_col].astype(str).str.strip().str.lower()==key.lower()]
    return found.iloc[0][value_col] if not found.empty else None

# Data tabs are read directly from the newly uploaded workbook.
if not XLSX_PATH.exists():
    st.error("No se encontró el Excel Movistar_Plan_de_Reestructuracion_corregido.xlsx. Colócalo junto a app.py.")
    st.stop()

try:
    xls=pd.ExcelFile(XLSX_PATH)
    available=[s for s in xls.sheet_names if s!="Liquidacion_Prelacion"]
except Exception as e:
    st.error(f"No se pudo abrir el Excel: {e}")
    st.stop()

with st.sidebar:
    if LOGO_PATH.exists(): st.image(str(LOGO_PATH),width=170)
    st.markdown("## INTEGRATEL PERÚ")
    st.caption("Panel de inteligencia financiera y concursal")
    st.divider()
    page=st.radio("NAVEGACIÓN",[
        "Vista ejecutiva","Crisis financiera","Acreedores y poder de voto",
        "Junta y escenarios","Ruta del procedimiento","Plan propuesto",
        "Capacidad de pago","Sensibilidad","Estados financieros",
        "Indicadores","Cronología","Datos y trazabilidad"
    ])
    st.divider()
    st.caption("Fuente activa")
    st.markdown("**Movistar_Plan_de_Reestructuracion_corregido.xlsx**")
    st.caption("La hoja de liquidación se excluye de la aplicación.")

logo_html=""
if LOGO_PATH.exists():
    import base64
    logo_html=f'<div class="hero-logo"><img src="data:image/png;base64,{base64.b64encode(LOGO_PATH.read_bytes()).decode()}"></div>'
st.markdown(f'<div class="hero"><div class="hero-grid">{logo_html}<div><p class="kicker">TELEFÓNICA DEL PERÚ S.A.A. · INTEGRATEL PERÚ</p><h1>Panel de inteligencia financiera y concursal</h1><p>Diagnóstico financiero · Junta de acreedores · Plan de reestructuración · Proyecciones</p></div></div></div>',unsafe_allow_html=True)

# ---------------------- VISTA EJECUTIVA ----------------------
if page=="Vista ejecutiva":
    section("Resumen ejecutivo","Indicadores y conclusiones principales tomados del nuevo libro de trabajo.")
    summary=load_sheet("Resumen",6)
    if not summary.empty and {"Indicador clave","Valor"}.issubset(summary.columns):
        kv=dict(zip(summary["Indicador clave"].astype(str),summary["Valor"]))
        def pick(*terms):
            for k,v in kv.items():
                if any(t.lower() in k.lower() for t in terms): return v
            return None
        a,b,c,d=st.columns(4)
        a.metric("Créditos reconocidos",f"S/ {fmt_num(pick('Total de créditos reconocidos'),1)} M")
        b.metric("Voto escenario A",fmt_pct(pick("Votos a favor en el escenario A")))
        c.metric("Mayoría requerida",fmt_pct(pick("Mayoría calificada requerida")))
        d.metric("DSCR mínimo · escenario base",f"{fmt_num(pick('DSCR mínimo años 1–15'),2)}x")
        # Se excluyen referencias que proceden de la hoja de liquidación, que no forma parte del aplicativo.
        if "Hoja de origen" in summary.columns:
            summary=summary[summary["Hoja de origen"].astype(str)!="Liquidacion_Prelacion"]
        summary=summary[~summary["Indicador clave"].astype(str).str.contains("liquidación",case=False,na=False)]
        table(summary,height=390)
    else:
        st.warning("No se encontró el resumen estructurado.")
    company=load_sheet("Empresa",2)
    if not company.empty:
        section("Indicadores de la empresa")
        table(company,height=350)
    note("Las cifras de acreedores y las proyecciones deben leerse según su etiqueta: dato, estimado, supuesto o cálculo. El panel no representa una decisión oficial de la Junta ni de Indecopi.",True)

# ---------------------- CRISIS FINANCIERA ----------------------
elif page=="Crisis financiera":
    section("Diagnóstico de la crisis financiera","Evolución 2020–2025 según la hoja Indicadores.")
    ind=load_sheet("Indicadores",2)
    if not ind.empty:
        years=[c for c in ind.columns if str(c).replace(".0","").isdigit()]
        metric_options=[c for c in ind["Indicador"].dropna().astype(str).unique() if not c[:1].isdigit()]
        if years:
            chosen=st.selectbox("Indicador para analizar",metric_options,index=0 if metric_options else None)
            row=ind[ind["Indicador"].astype(str)==chosen]
            if not row.empty:
                vals=pd.to_numeric(row.iloc[0][years],errors="coerce")
                plot=pd.DataFrame({"Año":[str(int(float(y))) for y in years],"Valor":vals.values})
                fig=px.line(plot,x="Año",y="Valor",markers=True,title=chosen)
                st.plotly_chart(chart_layout(fig),use_container_width=True)
                table(row,height=130)
        # KPI financial situation from last year
        def indicator_value(name):
            r=ind[ind["Indicador"].astype(str).str.strip().str.lower()==name.lower()]
            if r.empty or not years: return None
            return pd.to_numeric(r.iloc[0][years[-1]],errors="coerce")
        x,y,z=st.columns(3)
        x.metric("Ingresos 2025",f"S/ {fmt_num(indicator_value('Ingresos ordinarios'),0)} miles")
        y.metric("Pérdida neta 2025",f"S/ {fmt_num(indicator_value('Ganancia (pérdida) neta'),0)} miles")
        z.metric("Endeudamiento 2025",fmt_pct(indicator_value("Endeudamiento total")))
    else: st.warning("No se encontró la hoja Indicadores.")

# ---------------------- ACREEDORES ----------------------
elif page=="Acreedores y poder de voto":
    section("Acreedores y poder de voto","Detalle individual de acreencias y resumen de bloques según el Excel.")
    creditors=load_sheet("Acreedores",4)
    if not creditors.empty:
        detail=creditors[creditors.get("Acreedor",pd.Series(dtype=str)).notna()].copy()
        if "Acreedor" in detail:
            detail=detail[~detail["Acreedor"].astype(str).str.lower().isin(["total"])].copy()
        total=pd.to_numeric(detail.get("Monto (S/ M)",pd.Series(dtype=float)),errors="coerce").sum()
        a,b,c=st.columns(3)
        a.metric("Total de acreencias detalladas",f"S/ {fmt_num(total,1)} M")
        b.metric("Acreedores / conceptos",f"{len(detail)}")
        c.metric("Mayoría calificada","> 66,6 %")
        section("Detalle de acreedores por naturaleza")
        table(detail,height=300)
    # The workbook groups labor and suppliers in the voting block "Otros".
    block=load_sheet("Acreedores",14)
    if not block.empty:
        section("Bloques de votación del modelo")
        table(block,height=230)
        note("El Excel conserva «Otros» como bloque agregado para la votación, porque así está definido el modelo de pesos. En el detalle superior, los créditos laborales y los proveedores aparecen separados por naturaleza. No se deben atribuir votos independientes a esos dos componentes sin modificar la base de votación.",True)
    rank=load_sheet("Acreedores",24)
    if not rank.empty:
        section("Orden de prelación y tratamiento concursal")
        table(rank,height=310)
    note("Las cifras estimadas, supuestas o calculadas no equivalen necesariamente a créditos firmes reconocidos por Indecopi.",True)

# ---------------------- JUNTA ----------------------
elif page=="Junta y escenarios":
    section("Simulador de escenarios de la Junta","Escenarios consignados en la hoja Junta_Votacion; son simulaciones y no una votación real.")
    scenarios=load_sheet("Junta_Votacion",14)
    if not scenarios.empty and "Escenario" in scenarios.columns:
        scenarios=scenarios[scenarios["Escenario"].astype(str).str.fullmatch("[A-D]",na=False)].copy()
        labels={str(r["Escenario"]):str(r.get("Descripción","")) for _,r in scenarios.iterrows()}
        selected=st.selectbox("Selecciona escenario",scenarios["Escenario"].astype(str).tolist(),
                              format_func=lambda v:f"{v} · {labels.get(v,'')}")
        row=scenarios[scenarios["Escenario"].astype(str)==selected].iloc[0]
        yes=float(pd.to_numeric(row.get("% a favor de reestructuración",0),errors="coerce") or 0)*100
        no=float(pd.to_numeric(row.get("% a favor de liquidación",0),errors="coerce") or 0)*100
        a,b,c=st.columns(3)
        a.metric("A favor de reestructurar",f"{yes:.2f}%")
        b.metric("A favor de liquidar",f"{no:.2f}%")
        c.metric("Umbral legal","> 66,6 %")
        fig=go.Figure()
        fig.add_trace(go.Bar(name="A favor",x=["Votación"],y=[yes],marker_color=LIME))
        fig.add_trace(go.Bar(name="En contra",x=["Votación"],y=[no],marker_color=RED))
        fig.add_hline(y=66.6,line_dash="dash",line_color=AMBER,annotation_text="Mayoría calificada > 66,6 %")
        fig.update_layout(barmode="stack",yaxis=dict(range=[0,100],ticksuffix="%"))
        st.plotly_chart(chart_layout(fig,360),use_container_width=True)
        st.markdown(f"**Descripción:** {row.get('Descripción','—')}")
        table(scenarios,height=250)
    else: st.warning("No se encontró la tabla de escenarios.")
    votes=load_sheet("Junta_Votacion",8)
    if not votes.empty:
        with st.expander("Ver pesos de voto y reglas de instalación"):
            table(votes,height=220)
    note("El resultado depende de los créditos reconocidos, asistencia, convocatoria y reglas legales aplicables. Verifica cualquier cambio con el expediente oficial.",True)

# ---------------------- PROCEDIMIENTO ----------------------
elif page=="Ruta del procedimiento":
    section("Línea de tiempo del procedimiento concursal","Fechas, estado, plazo y base legal según la hoja Procedimiento.")
    proc=load_sheet("Procedimiento",3)
    if not proc.empty:
        proc=proc[proc["Etapa"].notna()].copy()
        status_values=proc["Estado"].dropna().astype(str).unique().tolist()
        chosen=st.multiselect("Filtrar por estado",status_values,default=status_values)
        proc=proc[proc["Estado"].astype(str).isin(chosen)]
        cards=[]
        for _,r in proc.iterrows():
            status=str(r.get("Estado","Sin estado"))
            date=r.get("Fecha (real o proyectada)","No indicada")
            if hasattr(date,"strftime"): date=date.strftime("%d/%m/%Y")
            status_l=status.lower()
            cls="done" if "cumplido" in status_l else ("pending" if "trámite" in status_l else ("risk" if "contingente" in status_l else ""))
            cards.append(f'<div class="timeline-item {cls}"><span class="t-date">{html.escape(str(date))}</span><span class="t-status">{html.escape(status)}</span><div class="t-title">{html.escape(str(r.get("Etapa","")))}</div><div class="t-desc">{html.escape(str(r.get("Qué ocurre","")))}</div><div class="t-desc"><b>Base legal:</b> {html.escape(str(r.get("Base legal","—")))}</div><div class="t-desc"><b>Responsable:</b> {html.escape(str(r.get("Responsable","—")))}</div></div>')
        st.markdown('<div class="timeline">'+"".join(cards)+"</div>",unsafe_allow_html=True)
        with st.expander("Ver tabla completa del procedimiento"):
            table(proc,height=440)
    else: st.warning("No se encontró la hoja Procedimiento.")
    chrono=load_sheet("Cronologia",2)
    if not chrono.empty:
        with st.expander("Antecedentes y cronología de la empresa"):
            table(chrono,height=380)

# ---------------------- PLAN ----------------------
elif page=="Plan propuesto":
    section("Plan de reestructuración propuesto","Tratamiento por acreedor, parámetros, cronograma y controles del modelo.")
    plan_params=load_sheet("Plan_Reestructuracion",4)
    if not plan_params.empty:
        section("Parámetros principales")
        table(plan_params,height=230)
    creditors_plan=load_sheet("Plan_Reestructuracion",16)
    if not creditors_plan.empty:
        section("Tratamiento por acreedor")
        table(creditors_plan,height=300)
        amount_col=next((c for c in creditors_plan.columns if "A pagar en efectivo" in c),None)
        cap_col=next((c for c in creditors_plan.columns if "Capitalizado en acciones" in c),None)
        if amount_col and cap_col and "Acreedor" in creditors_plan:
            chart=creditors_plan[creditors_plan["Acreedor"].astype(str)!="Total"].copy()
            for col in [amount_col,cap_col]: chart[col]=pd.to_numeric(chart[col],errors="coerce")
            fig=px.bar(chart,x="Acreedor",y=[cap_col,amount_col],barmode="group",
                       title="Capitalización frente a pago en efectivo",
                       color_discrete_sequence=[LIME,CYAN])
            st.plotly_chart(chart_layout(fig,390),use_container_width=True)
    plan_sched=load_sheet("Plan_Reestructuracion",39)
    if not plan_sched.empty:
        section("Cronograma anual del servicio de deuda")
        yearcol=next((c for c in plan_sched.columns if c.lower()=="año"),None)
        if yearcol:
            cols=[c for c in ["Interés","Amortización programada","Barrido de caja (prepago)","Saldo final"] if c in plan_sched.columns]
            for col in cols: plan_sched[col]=pd.to_numeric(plan_sched[col],errors="coerce")
            fig=px.line(plan_sched,x=yearcol,y=cols,markers=True,title="Evolución del cronograma")
            st.plotly_chart(chart_layout(fig,390),use_container_width=True)
        table(plan_sched,height=390)
    controls=load_sheet("Plan_Reestructuracion",60)
    if not controls.empty:
        section("Controles y lectura del plan")
        table(controls,height=330)
    note("El plan es una propuesta modelada en el Excel; no debe presentarse como aprobado ni vigente mientras no exista acuerdo formal de la Junta.",True)

# ---------------------- CAPACIDAD DE PAGO ----------------------
elif page=="Capacidad de pago":
    section("Capacidad de pago","Flujo de caja disponible para el servicio de la deuda concursal (FCSD).")
    cap=load_sheet("Capacidad_Pago",4)
    if not cap.empty:
        if "Año" in cap:
            cap=cap[pd.to_numeric(cap["Año"],errors="coerce").notna()].copy()
            for c in ["Ingresos","EBITDA","FCSD"]:
                if c in cap: cap[c]=pd.to_numeric(cap[c],errors="coerce")
            if "FCSD" in cap and "Año" in cap:
                fig=px.bar(cap,x="Año",y="FCSD",color="Año del plan" if "Año del plan" in cap else None,
                           title="Flujo de caja disponible para el servicio de deuda",
                           color_discrete_sequence=[CYAN,LIME,BLUE])
                st.plotly_chart(chart_layout(fig,390),use_container_width=True)
        table(cap,height=420)
    raw_cap=pd.read_excel(XLSX_PATH,sheet_name="Capacidad_Pago",header=None)
    read=raw_cap.iloc[25:34,:3].copy()
    read.columns=["Indicador","Valor","Detalle"]
    read=read.dropna(how="all")
    if not read.empty:
        section("Lectura de capacidad y cobertura")
        table(read,height=260)
    note("2025 se presenta como año real referencial; 2026–2042 son proyecciones del modelo. Las proyecciones dependen de supuestos operativos y financieros.",True)

# ---------------------- SENSIBILIDAD ----------------------
elif page=="Sensibilidad":
    section("Análisis de sensibilidad","Comparación del cronograma frente a escenarios pesimista, base y optimista.")
    sens=load_sheet("Sensibilidad",3)
    if not sens.empty:
        sens=sens[pd.to_numeric(sens.get("Año",pd.Series(dtype=float)),errors="coerce").notna()].copy()
        cols=[c for c in ["FCSD Pesimista","FCSD Base","FCSD Optimista","Servicio programado (sin barrido)"] if c in sens.columns]
        for c in cols: sens[c]=pd.to_numeric(sens[c],errors="coerce")
        if cols:
            fig=px.line(sens,x="Año",y=cols,markers=True,title="FCSD y servicio programado por escenario")
            st.plotly_chart(chart_layout(fig,410),use_container_width=True)
        table(sens,height=430)
    summary=load_sheet("Sensibilidad",25)
    if not summary.empty:
        section("Resumen comparativo de escenarios")
        table(summary,height=250)
    note("El DSCR compara el flujo disponible con el servicio de deuda. Un valor inferior a 1,0x indica que el flujo del periodo no cubre por sí solo el servicio programado.",False)

# ---------------------- ESTADOS FINANCIEROS ----------------------
elif page=="Estados financieros":
    section("Estados financieros comparativos","Información de las hojas ESF, ER y EFE del nuevo Excel.")
    choices={"ESF":"Estado de Situación Financiera","ER":"Estado de Resultados","EFE":"Estado de Flujo de Efectivo"}
    tab=st.selectbox("Estado financiero",list(choices),format_func=lambda x:choices[x])
    df=load_sheet(tab,2)
    if not df.empty:
        years=[c for c in df.columns if str(c).replace(".0","").isdigit()]
        if years:
            years=sorted(years,key=lambda v:float(str(v).replace(".0","")))
            selected=st.selectbox("Año",years,index=len(years)-1,format_func=lambda x:str(int(float(x))))
            account=next((c for c in df.columns if c.lower()=="cuenta"),df.columns[1])
            keep=[c for c in [account,"Nota",selected] if c in df.columns]
            view=df[keep].copy()
            view=view.dropna(subset=[account])
            view=view[~view[account].astype(str).str.lower().isin(["none","nan","total control"])]
            amount=pd.to_numeric(view[selected],errors="coerce")
            a,b,c=st.columns(3)
            a.metric("Cuentas mostradas",f"{len(view)}")
            b.metric("Saldos positivos",f"{int((amount>0).sum())}")
            c.metric("Saldos negativos",f"{int((amount<0).sum())}")
            view[selected]=amount.map(lambda v:fmt_num(v,0) if pd.notna(v) else "—")
            view=view.rename(columns={selected:f"Importe · {int(float(selected))} (miles de S/)"})
            table(view,height=650)
            st.download_button("Descargar vista CSV",view.to_csv(index=False).encode("utf-8-sig"),f"{tab}_{int(float(selected))}.csv","text/csv")
        else: table(df,height=650)
    else: st.warning("No se encontró la hoja seleccionada.")

# ---------------------- INDICADORES ----------------------
elif page=="Indicadores":
    section("Indicadores financieros y de solvencia","Series comparativas 2020–2025; los indicadores se toman del Excel.")
    ind=load_sheet("Indicadores",2)
    if not ind.empty:
        groups=[]
        for v in ind["Indicador"].dropna().astype(str):
            if v[:1].isdigit(): groups.append(v)
        group=st.selectbox("Grupo de indicadores",groups)
        # Se muestran los indicadores posteriores al rótulo de grupo hasta el siguiente rótulo.
        idx=ind.index[ind["Indicador"].astype(str)==group][0]
        next_idx=next((i for i in ind.index if i>idx and str(ind.loc[i,"Indicador"])[:1].isdigit()),None)
        sub=ind.loc[idx+1:next_idx-1 if next_idx is not None else ind.index.max()].copy()
        sub=sub[sub["Indicador"].notna()]
        table(sub,height=470)
    else: st.warning("No se encontró la hoja Indicadores.")

# ---------------------- CRONOLOGÍA ----------------------
elif page=="Cronología":
    section("Antecedentes relevantes","Hitos corporativos y financieros que contextualizan el procedimiento.")
    chrono=load_sheet("Cronologia",2)
    table(chrono,height=560)

# ---------------------- DATOS / TRAZABILIDAD ----------------------
elif page=="Datos y trazabilidad":
    section("Supuestos, información empresarial y base de datos")
    tabs=st.tabs(["Supuestos","Empresa","Base_Datos","Hojas incluidas"])
    with tabs[0]:
        sup=load_sheet("Supuestos",3)
        table(sup,height=520)
    with tabs[1]:
        emp=load_sheet("Empresa",2)
        table(emp,height=420)
    with tabs[2]:
        base=load_sheet("Base_Datos",0)
        if not base.empty:
            q=st.text_input("Buscar cuenta o estado financiero")
            if q:
                mask=base.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)
                base=base[mask]
            table(base.head(600),height=520)
            st.caption("Se muestran hasta 600 registros para mantener una navegación fluida. La base completa permanece en el Excel descargable.")
    with tabs[3]:
        visible=[s for s in available if s!="Liquidacion_Prelacion"]
        st.write("Hojas utilizadas por el aplicativo:")
        st.write(" · ".join(visible))
        st.caption("La hoja Liquidacion_Prelacion se excluye de la navegación y de los análisis del aplicativo.")

st.divider()
st.caption("Aplicación académica de análisis financiero y concursal. Datos, supuestos y proyecciones se presentan diferenciados; validar siempre la información con documentos y fuentes oficiales.")
