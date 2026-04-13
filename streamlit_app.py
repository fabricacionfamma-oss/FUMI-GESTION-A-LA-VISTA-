import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
import calendar
from fpdf import FPDF
from datetime import timedelta

# ==========================================
# 0. CONFIGURACIÓN Y MAPEO
# ==========================================
st.set_page_config(page_title="Reportes Fumiscor", layout="wide", page_icon="📊")

MAQUINAS_MAP = {
    "P-023": "PRENSAS PROGRESIVAS", "P-024": "PRENSAS PROGRESIVAS", "P-025": "PRENSAS PROGRESIVAS", "P-026": "PRENSAS PROGRESIVAS",
    "P-027": "PRENSAS PROGRESIVAS GRANDES", "BAL-002": "BALANCIN", "BAL-003": "BALANCIN", "BAL-005": "BALANCIN", "BAL-006": "BALANCIN",
    "BAL-007": "BALANCIN", "BAL-008": "BALANCIN", "BAL-009": "BALANCIN", "BAL-010": "BALANCIN", "P-011": "HIDRAULICAS", 
    "P-016": "HIDRAULICAS", "P-017": "HIDRAULICAS", "P-018": "HIDRAULICAS", "P-015": "MECANICAS", "P-019": "MECANICAS", 
    "P-020": "MECANICAS", "P-021": "MECANICAS", "P-022": "MECANICAS", "GOF01": "Gofradora",
    "SOP-003": "PRP", "SOP-005": "PRP", "SOP-008": "PRP", "SOP-009": "PRP", "SOP-010": "PRP",
    "SOP-017": "PRP", "SOP-018": "PRP", "SOP-019": "PRP", "SOP-020": "PRP", "SOP-022": "PRP",
    "SOP-023": "PRP", "SOP-024": "PRP", "SOP-025": "PRP", "DOB-001": "DOBLADORA", "DOB-002": "DOBLADORA", 
    "DOB-003": "DOBLADORA", "DOB-004": "DOBLADORA", "DOB-005": "DOBLADORA", "DOB-006": "DOBLADORA",
    "Celda 01 Fumis": "CELDA SOLDADURA RENAULT", "Celda 02 Fumis": "CELDA SOLDADURA RENAULT"
}

GRUPOS_ESTAMPADO = ['PRENSAS PROGRESIVAS', 'PRENSAS PROGRESIVAS GRANDES', 'BALANCIN', 'HIDRAULICAS', 'MECANICAS', 'Gofradora']
GRUPOS_SOLDADURA = ['PRP', 'DOBLADORA', 'CELDA SOLDADURA', 'CELDA SOLDADURA RENAULT']

class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area; self.fecha_str = fecha_str; self.theme_color = theme_color

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

# ==========================================
# 1. CARGA DE DATOS (SQL)
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_full(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_s = fecha_ini.strftime('%Y-%m-%d'); fin_s = fecha_fin.strftime('%Y-%m-%d')
        
        # Métricas OEE
        q_metrics = f"SELECT c.Name as Máquina, SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas, SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada, (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE, (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD, (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD, (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE FROM PROD_D_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Date BETWEEN '{ini_s}' AND '{fin_s}' GROUP BY c.Name"
        
        # Tendencia Mensual
        q_trend = f"SELECT p.Month, c.Name as Máquina, SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas, SUM(p.Good + p.Rework + p.Scrap) as Totales FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} GROUP BY p.Month, c.Name"
        
        # Fallas
        q_event = f"SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2] FROM EVENT_01 e LEFT JOIN CELL c ON e.CellId = c.CellId LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId WHERE e.Date BETWEEN '{ini_s}' AND '{fin_s}'"
        
        # Detalle por Pieza
        q_piezas = f"SELECT c.Name as Máquina, pr.Code as Pieza, SUM(p.Scrap) as Scrap, SUM(p.Rework) as RT FROM PROD_D_01 p JOIN CELL c ON p.CellId = c.CellId JOIN PRODUCT pr ON p.ProductId = pr.ProductId WHERE p.Date BETWEEN '{ini_s}' AND '{fin_s}' GROUP BY c.Name, pr.Code"

        return conn.query(q_metrics), conn.query(q_event), conn.query(q_trend), conn.query(q_piezas)
    except Exception as e:
        st.error(f"Error SQL: {e}"); return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 2. MOTOR INFORME PRODUCTIVO
# ==========================================
def crear_pdf_productivo(area, label, df_trend, df_piezas, mes_sel, anio_sel):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    grupos = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa = {k.upper(): v for k, v in MAQUINAS_MAP.items()}
    
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label, theme_color)
    pdf.add_page(orientation='L')
    
    # Filtrar
    df_t = df_trend.copy(); df_t['G'] = df_t['Máquina'].str.upper().map(mapa)
    df_t = df_t[df_t['G'].isin(grupos)]
    df_p = df_piezas.copy(); df_p['G'] = df_p['Máquina'].str.upper().map(mapa)
    df_p = df_p[df_p['G'].isin(grupos)]

    # Encabezado
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(20, 6, "MES", 1, 0, 'C'); pdf.cell(20, 6, "AÑO", 1, 0, 'C')
    pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C'); pdf.cell(40, 6, "AREA", 1, 1, 'C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(20, 6, str(mes_sel), 1, 0, 'C'); pdf.cell(20, 6, str(anio_sel), 1, 0, 'C')
    pdf.cell(197, 6, f"PLANTA: {area.upper()}", 1, 0, 'C'); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C')

    # Gráficos
    df_ev = df_t.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
    df_ev['% Scrap'] = (df_ev['Observadas'] / df_ev['Totales'].replace(0, 1)) * 100
    df_ev['% RT'] = (df_ev['Retrabajo'] / df_ev['Totales'].replace(0, 1)) * 100

    def get_img(fig):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig.write_image(tmp.name); return tmp.name

    # Izquierda: Evolución
    f1 = px.bar(df_ev, x='Month', y='Totales', title="PIEZAS PRODUCIDAS MES A MES")
    f2 = px.bar(df_ev, x='Month', y='% Scrap', title="% DE SCRAP MES A MES")
    f3 = px.bar(df_ev, x='Month', y='% RT', title="% DE RT MES A MES")
    for f in [f1, f2, f3]: f.update_layout(margin=dict(l=5,r=5,t=30,b=5), height=200)

    pdf.image(get_img(f1), 10, 48, 130); pdf.image(get_img(f2), 10, 98, 130); pdf.image(get_img(f3), 10, 148, 130)

    # Derecha: Top Piezas
    t_s = df_p.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index()
    t_rt = df_p.groupby('Pieza')['RT'].sum().nlargest(5).reset_index()
    f4 = px.bar(t_s, x='Scrap', y='Pieza', orientation='h', title="TOP 5 SCRAP POR PIEZA")
    f5 = px.bar(t_rt, x='RT', y='Pieza', orientation='h', title="TOP 5 RT POR PIEZA")
    for f in [f4, f5]: f.update_layout(margin=dict(l=5,r=5,t=30,b=5), height=300)

    pdf.image(get_img(f4), 145, 48, 140); pdf.image(get_img(f5), 145, 110, 140)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 3. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Sistema de Reportes Fumiscor")
st.divider()

col1, col2, col3 = st.columns(3)
with col1: tipo = st.selectbox("Periodo", ["Mensual", "Diario"])
with col2: mes = st.selectbox("Mes", range(1, 13), index=pd.Timestamp.now().month-1)
with col3: anio = st.selectbox("Año", [2024, 2025, 2026], index=0)

# Simulación de fechas para la query
f_ini = pd.Timestamp(anio, mes, 1)
f_fin = f_ini + pd.offsets.MonthEnd(0)

# Cargar Datos
df_metrics, df_event, df_trend, df_piezas = fetch_data_full(f_ini, f_fin, mes, anio)

st.subheader("Descargar Informes Productivos")
b1, b2 = st.columns(2)

with b1:
    if st.button("📥 Informe Productivo ESTAMPADO"):
        with st.spinner("Generando..."):
            pdf = crear_pdf_productivo("Estampado", f"{mes}-{anio}", df_trend, df_piezas, mes, anio)
            st.download_button("Click para Descargar PDF", pdf, f"Productivo_Estampado_{mes}_{anio}.pdf")

with b2:
    if st.button("📥 Informe Productivo SOLDADURA"):
        with st.spinner("Generando..."):
            pdf = crear_pdf_productivo("Soldadura", f"{mes}-{anio}", df_trend, df_piezas, mes, anio)
            st.download_button("Click para Descargar PDF", pdf, f"Productivo_Soldadura_{mes}_{anio}.pdf")
