import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
import calendar
from fpdf import FPDF
from datetime import timedelta

# ==========================================
# 0. CONFIGURACIÓN Y CONSTANTES
# ==========================================
st.set_page_config(page_title="Reportes Fumiscor", layout="wide", page_icon="📊")

MAQUINAS_MAP = {
    # (Se mantiene tu diccionario actual de máquinas...)
    "P-023": "PRENSAS PROGRESIVAS", "P-024": "PRENSAS PROGRESIVAS", "P-025": "PRENSAS PROGRESIVAS", "P-026": "PRENSAS PROGRESIVAS",
    "P-027": "PRENSAS PROGRESIVAS GRANDES", "BAL-002": "BALANCIN", "BAL-003": "BALANCIN", "BAL-005": "BALANCIN", "BAL-006": "BALANCIN",
    "BAL-007": "BALANCIN", "BAL-008": "BALANCIN", "BAL-009": "BALANCIN", "BAL-010": "BALANCIN",
    "SOP-003": "PRP", "SOP-005": "PRP", "SOP-008": "PRP", "SOP-009": "PRP", "SOP-010": "PRP",
    "DOB-001": "DOBLADORA", "DOB-002": "DOBLADORA", "DOB-003": "DOBLADORA", "DOB-004": "DOBLADORA",
    "Celda 01 Fumis": "CELDA SOLDADURA RENAULT", "Celda 02 Fumis": "CELDA SOLDADURA RENAULT"
}

GRUPOS_ESTAMPADO = ['PRENSAS PROGRESIVAS', 'PRENSAS PROGRESIVAS GRANDES', 'BALANCIN', 'HIDRAULICAS', 'MECANICAS', 'Gofradora']
GRUPOS_SOLDADURA = ['PRP', 'DOBLADORA', 'CELDA SOLDADURA', 'CELDA SOLDADURA RENAULT']

# ==========================================
# 1. FUNCIONES AUXILIARES
# ==========================================
class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area
        self.fecha_str = fecha_str
        self.theme_color = theme_color

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

# ==========================================
# 2. CARGA DE DATOS
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, tipo_periodo, mes=None, anio=None):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d')
        fin_str = fecha_fin.strftime('%Y-%m-%d')

        # Consulta de Métricas OEE
        q_metrics = f"""
            SELECT c.Name as Máquina, 
                   SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                   SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada,
                   (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE,
                   (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD,
                   (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD,
                   (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE
            FROM PROD_D_03 p JOIN CELL c ON p.CellId = c.CellId
            WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}'
            GROUP BY c.Name
        """
        
        # Consulta para Evolución Mensual (Productividad/Calidad)
        q_trend = f"""
            SELECT p.Month, c.Name as Máquina,
                   SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                   SUM(p.Good + p.Rework + p.Scrap) as Totales
            FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId
            WHERE p.Year = {anio if anio else 2024}
            GROUP BY p.Month, c.Name
        """

        # Consulta de Eventos (Fallas)
        q_event = f"""
            SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2]
            FROM EVENT_01 e
            LEFT JOIN CELL c ON e.CellId = c.CellId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        
        # Consulta por Pieza (Para Top Scrap/RT)
        q_piezas = f"""
            SELECT c.Name as Máquina, pr.Code as Pieza,
                   SUM(p.Scrap) as Scrap, SUM(p.Rework) as RT
            FROM PROD_D_01 p 
            JOIN CELL c ON p.CellId = c.CellId
            JOIN PRODUCT pr ON p.ProductId = pr.ProductId
            WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}'
            GROUP BY c.Name, pr.Code
        """

        return conn.query(q_metrics), conn.query(q_event), conn.query(q_trend), conn.query(q_piezas)
    except Exception as e:
        st.error(f"Error DB: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR INFORME PRODUCTIVO (NUEVO WIREFRAME)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, df_metrics, mes_sel, anio_sel):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    grupos_area = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa_limpio = {str(k).strip().upper(): v for k, v in MAQUINAS_MAP.items()}

    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    pdf.add_page(orientation='L')

    # Filtro de datos por área
    df_t = df_trend.copy()
    df_t['Grupo'] = df_t['Máquina'].str.strip().str.upper().map(mapa_limpio)
    df_t = df_t[df_t['Grupo'].isin(grupos_area)]
    
    df_p = df_piezas.copy()
    df_p['Grupo'] = df_p['Máquina'].str.strip().str.upper().map(mapa_limpio)
    df_p = df_p[df_p['Grupo'].isin(grupos_area)]

    # --- ENCABEZADO ---
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(20, 6, "MES", 1, 0, 'C'); pdf.cell(20, 6, "AÑO", 1, 0, 'C')
    pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C')
    pdf.cell(40, 6, "AREA", 1, 1, 'C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(20, 6, str(mes_sel), 1, 0, 'C'); pdf.cell(20, 6, str(anio_sel), 1, 0, 'C')
    pdf.cell(197, 6, f"PLANTA: {area.upper()}", 1, 0, 'C')
    pdf.cell(40, 6, "GENERAL", 1, 1, 'C')

    # --- KPIs ---
    pdf.set_fill_color(100, 150, 200) # Azul Fumiscor
    y_kpi = 25
    kpis = ["OEE", "PERFORMANCE", "DISPONIBILIDAD", "CALIDAD"]
    for i, txt in enumerate(kpis):
        pdf.rect(10 + (i*69), y_kpi, 65, 18, 'F')
        pdf.set_xy(10 + (i*69), y_kpi + 2)
        pdf.set_font("Arial", 'B', 9); pdf.set_text_color(255)
        pdf.cell(65, 5, txt, 0, 1, 'L')
    pdf.set_text_color(0)

    # --- COLUMNA IZQUIERDA: EVOLUCIÓN MENSUAL ---
    df_ev = df_t.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
    df_ev['% Scrap'] = (df_ev['Observadas'] / df_ev['Totales']) * 100
    df_ev['% RT'] = (df_ev['Retrabajo'] / df_ev['Totales']) * 100

    def save_chart(fig, w, h):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig.write_image(tmp.name)
            return tmp.name

    # 1. Piezas mes a mes
    fig1 = px.bar(df_ev, x='Month', y='Totales', title="PIEZAS PRODUCIDAS MES A MES")
    fig1.update_layout(margin=dict(l=5, r=5, t=30, b=5), height=200)
    
    # 2. % Scrap mes a mes
    fig2 = px.bar(df_ev, x='Month', y='% Scrap', title="GRAFICO DE BARRAS % DE SCRAP MES A MES")
    fig2.update_layout(margin=dict(l=5, r=5, t=30, b=5), height=200)

    # 3. % RT mes a mes
    fig3 = px.bar(df_ev, x='Month', y='% RT', title="GRAFICO DE BARRAS % DE RT MES A MES")
    fig3.update_layout(margin=dict(l=5, r=5, t=30, b=5), height=200)

    pdf.image(save_chart(fig1, 400, 200), 10, 48, 130)
    pdf.image(save_chart(fig2, 400, 200), 10, 98, 130)
    pdf.image(save_chart(fig3, 400, 200), 10, 148, 130)

    # --- COLUMNA DERECHA: TOP POR PIEZA ---
    top_scrap = df_p.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index()
    top_rt = df_p.groupby('Pieza')['RT'].sum().nlargest(5).reset_index()

    fig4 = px.bar(top_scrap, x='Scrap', y='Pieza', orientation='h', title="top scrap por pieza")
    fig4.update_layout(margin=dict(l=5, r=5, t=30, b=5), height=300)
    
    fig5 = px.bar(top_rt, x='RT', y='Pieza', orientation='h', title="top rt por pieza")
    fig5.update_layout(margin=dict(l=5, r=5, t=30, b=5), height=300)

    pdf.image(save_chart(fig4, 400, 300), 145, 48, 140)
    pdf.image(save_chart(fig5, 400, 300), 145, 110, 140)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. INTERFAZ Y BOTONES
# ==========================================
st.title("📄 Sistema de Reportes Fumiscor")
st.divider()

# Selectores de fecha (simplificados)
c1, c2, c3 = st.columns(3)
with c1: tipo = st.selectbox("Periodo", ["Mensual", "Diario"])
with c2: mes = st.selectbox("Mes", range(1, 13), index=pd.Timestamp.now().month-1)
with c3: anio = st.selectbox("Año", [2024, 2025])

# Carga de datos
df_metrics, df_raw, df_trend, df_piezas = fetch_data_from_db(pd.Timestamp(anio, mes, 1), pd.Timestamp(anio, mes, 28), tipo, mes, anio)

st.write("### Generar Informes")
b1, b2 = st.columns(2)

with b1:
    if st.button("Descargar Informe Productivo ESTAMPADO"):
        pdf = crear_pdf_informe_productivo("Estampado", f"{mes}-{anio}", df_trend, df_piezas, df_metrics, mes, anio)
        st.download_button("Click para descargar", pdf, "Informe_Productivo_Estampado.pdf")

with b2:
    if st.button("Descargar Informe Productivo SOLDADURA"):
        pdf = crear_pdf_informe_productivo("Soldadura", f"{mes}-{anio}", df_trend, df_piezas, df_metrics, mes, anio)
        st.download_button("Click para descargar", pdf, "Informe_Productivo_Soldadura.pdf")
