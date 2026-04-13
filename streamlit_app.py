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
    "P-023": "PRENSAS PROGRESIVAS", "P-024": "PRENSAS PROGRESIVAS", "P-025": "PRENSAS PROGRESIVAS", "P-026": "PRENSAS PROGRESIVAS",
    "P-027": "PRENSAS PROGRESIVAS GRANDES", "BAL-002": "BALANCIN", "BAL-003": "BALANCIN", "BAL-005": "BALANCIN", "BAL-006": "BALANCIN",
    "BAL-007": "BALANCIN", "BAL-008": "BALANCIN", "BAL-009": "BALANCIN", "BAL-010": "BALANCIN", "P-011": "HIDRAULICAS", 
    "P-016": "HIDRAULICAS", "P-017": "HIDRAULICAS", "P-018": "HIDRAULICAS", "P-015": "MECANICAS", "P-019": "MECANICAS", 
    "P-020": "MECANICAS", "P-021": "MECANICAS", "P-022": "MECANICAS", "GOF01": "Gofradora",
    "SOP-003": "PRP", "SOP-005": "PRP", "SOP-008": "PRP", "SOP-009": "PRP", "SOP-010": "PRP",
    "SOP-017": "PRP", "SOP-018": "PRP", "SOP-019": "PRP", "SOP-020": "PRP", "SOP-022": "PRP",
    "SOP-023": "PRP", "SOP-024": "PRP", "SOP-025": "PRP", "DOB-001": "DOBLADORA", "DOB-002": "DOBLADORA", 
    "DOB-003": "DOBLADORA", "DOB-004": "DOBLADORA", "DOB-005": "DOBLADORA", "DOB-006": "DOBLADORA",
    "Celda 01 Fumis": "CELDA SOLDADURA RENAULT", "Celda 02 Fumis": "CELDA SOLDADURA RENAULT",
    "Celda 03 Fumis": "CELDA SOLDADURA RENAULT", "Celda 04 Fumis": "CELDA SOLDADURA RENAULT",
    "Celda 05 Fumis": "CELDA SOLDADURA RENAULT", "Celda 06 Fumis": "CELDA SOLDADURA RENAULT"
}

GRUPOS_ESTAMPADO = ['PRENSAS PROGRESIVAS', 'PRENSAS PROGRESIVAS GRANDES', 'BALANCIN', 'HIDRAULICAS', 'MECANICAS', 'Gofradora']
GRUPOS_SOLDADURA = ['PRP', 'DOBLADORA', 'CELDA SOLDADURA', 'CELDA SOLDADURA RENAULT']

# ==========================================
# 1. FUNCIONES AUXILIARES Y PDF
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

def save_chart(fig, w=600, h=300):
    fig.update_layout(width=w, height=h)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig.write_image(tmp.name, engine="kaleido", scale=2.0)
        return tmp.name

# ==========================================
# 2. CARGA DE DATOS UNIFICADA (SQL)
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d'); fin_str = fecha_fin.strftime('%Y-%m-%d')

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
        
        q_event = f"""
            SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2]
            FROM EVENT_01 e
            LEFT JOIN CELL c ON e.CellId = c.CellId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        
        q_trend = f"""
            SELECT p.Month, c.Name as Máquina,
                   SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                   SUM(p.Good + p.Rework + p.Scrap) as Totales
            FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId
            WHERE p.Year = {anio} AND p.Month <= {mes}
            GROUP BY p.Month, c.Name
        """
        
        q_piezas = f"""
            SELECT c.Name as Máquina, pr.Code as Pieza,
                   SUM(p.Scrap) as Scrap, SUM(p.Rework) as RT
            FROM PROD_D_01 p 
            JOIN CELL c ON p.CellId = c.CellId
            JOIN PRODUCT pr ON p.ProductId = pr.ProductId
            WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}'
            GROUP BY c.Name, pr.Code
        """

        df_metrics = conn.query(q_metrics)
        df_raw = conn.query(q_event)
        df_trend = conn.query(q_trend)
        df_piezas = conn.query(q_piezas)

        # --- CORRECCIÓN DE DATOS VACÍOS ---
        if df_raw.empty:
            # Si no hay eventos, creamos las columnas de todos modos para evitar el KeyError
            df_raw = pd.DataFrame(columns=[
                'Máquina', 'Tiempo (Min)', 'Nivel Evento 1', 'Nivel Evento 2', 
                'Estado_Global', 'Categoria_Macro', 'Detalle_Final'
            ])
        else:
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)
            
            def categorizar_estado(row):
                texto = f"{row.get('Nivel Evento 1','')} {row.get('Nivel Evento 2','')} ".upper()
                if 'PRODUCCION' in texto or 'PRODUCCIÓN' in texto: return 'Producción'
                if 'PARADA PROGRAMADA' in texto: return 'Parada Programada'
                return 'Falla/Gestión'
            
            def clasificar_macro(row):
                n1 = str(row.get('Nivel Evento 1', '')).strip().upper()
                if 'GESTION' in n1 or 'GESTIÓN' in n1: return 'Gestión'
                if 'FALLA' in n1: return str(row.get('Nivel Evento 2', '')).title()
                return n1.title()

            df_raw['Estado_Global'] = df_raw.apply(categorizar_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(clasificar_macro, axis=1)
            df_raw['Detalle_Final'] = df_raw['Categoria_Macro']

        return df_metrics, df_raw, df_trend, df_piezas

    except Exception as e:
        st.error(f"Error de conexión SQL: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (DISPONIBILIDAD)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    theme_hex = '#%02x%02x%02x' % theme_color
    grupos_area = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa_limpio = {str(k).strip().upper(): v for k, v in MAQUINAS_MAP.items()}

    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    
    df_m = df_metrics_pdf.copy()
    if not df_m.empty and df_m['OEE'].max() > 1.5:
        df_m[['OEE', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD']] /= 100.0

    df_m['Grupo'] = df_m['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('Otro')
    df_m = df_m[df_m['Grupo'].isin(grupos_area)]
    
    df_r = df_pdf_raw.copy()
    if not df_r.empty:
        df_r['Grupo_Máquina'] = df_r['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('Otro')
        df_r = df_r[df_r['Grupo_Máquina'].isin(grupos_area)]

    paginas = ['GENERAL'] + [g for g in grupos_area if g in df_m['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L')
        df_m_target = df_m if target == 'GENERAL' else df_m[df_m['Grupo'] == target]
        
        if df_r.empty:
            df_r_target = df_r
        else:
            df_r_target = df_r if target == 'GENERAL' else df_r[df_r['Grupo_Máquina'] == target]
            
        x_barras = 'Grupo' if target == 'GENERAL' else 'Máquina'
        pie_col = 'Grupo_Máquina' if target == 'GENERAL' else 'Categoria_Macro'
        
        # --- ENCABEZADO FORMATO EXCEL ---
        pdf.set_y(10)
        pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True)
        pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True)
        pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C')
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C')
        pdf.set_font("Arial", '', 10)
        pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C')

        # --- KPIs ---
        t_plan = df_m_target['T_Operativo'].sum() + df_m_target['T_Parada'].sum() if not df_m_target.empty else 0
        t_op = df_m_target['T_Operativo'].sum() if not df_m_target.empty else 0
        
        kpis = {
            "OEE": (df_m_target['OEE'] * (df_m_target['T_Operativo'] + df_m_target['T_Parada'])).sum() / t_plan if t_plan > 0 else 0,
            "PERFORMANCE": (df_m_target['PERFORMANCE'] * df_m_target['T_Operativo']).sum() / t_op if t_op > 0 else 0,
            "DISPONIBILIDAD": (df_m_target['DISPONIBILIDAD'] * (df_m_target['T_Operativo'] + df_m_target['T_Parada'])).sum() / t_plan if t_plan > 0 else 0,
            "CALIDAD": df_m_target['CALIDAD'].mean() if not df_m_target.empty else 0
        }

        y_kpi = 25
        for i, (lbl, val) in enumerate(kpis.items()):
            x = 10 + (i * 68.5)
            pdf.set_fill_color(*theme_color)
            pdf.rect(x, y_kpi, 65, 20, 'F')
            pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(255); pdf.cell(65, 6, lbl, 0, 1, 'L')
            pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{val*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        # --- Gráficos Barras OEE ---
        def add_bar(df_in, col, title, x_pos, y_pos):
            if df_in.empty: return
            df_g = df_in.groupby(x_barras)[col].mean().reset_index()
            fig = px.bar(df_g, x=x_barras, y=col, title=title, text_auto='.1%', color_discrete_sequence=[theme_hex])
            fig.update_layout(
                margin=dict(t=35, b=25, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', 
                yaxis=dict(range=[0, 1.1], visible=False), xaxis_title=""
            )
            fig.update_traces(textposition="outside", cliponaxis=False, textfont_size=12)
            img = save_chart(fig, w=600, h=220); pdf.image(img, x=x_pos, y=y_pos, w=135); os.remove(img)

        add_bar(df_m_target, 'OEE', 'OEE (%)', 10, 48)
        add_bar(df_m_target, 'PERFORMANCE', 'PERFORMANCE (%)', 150, 48)
        add_bar(df_m_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%)', 10, 102)
        add_bar(df_m_target, 'CALIDAD', 'CALIDAD (%)', 150, 102)

        # --- Top Fallos y Torta ---
        pdf.set_xy(10, 156); pdf.set_font("Arial", 'B', 10)
        pdf.cell(136, 6, "TOP 5 FALLOS", border='B', ln=True, align='C')
        
        df_f = df_r_target[df_r_target['Estado_Global'] == 'Falla/Gestión'] if not df_r_target.empty else pd.DataFrame()
        if not df_f.empty and df_f['Tiempo (Min)'].sum() > 0:
            t_total = df_f['Tiempo (Min)'].sum()
            top5 = df_f.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
            
            pdf.set_xy(10, 162); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
            pdf.cell(76, 5, "FALLO", border=1, fill=True); pdf.cell(30, 5, "MINUTOS", border=1, align='C', fill=True); pdf.cell(30, 5, "% TOTAL", border=1, align='C', ln=True, fill=True)
            pdf.set_font("Arial", '', 8); pdf.set_text_color(0)
            
            for _, r in top5.iterrows():
                pdf.set_x(10); pdf.cell(76, 6, clean_text(str(r['Detalle_Final']))[:45], border=1)
                pdf.cell(30, 6, f"{r['Tiempo (Min)']:.0f}", border=1, align='C')
                pdf.cell(30, 6, f"{(r['Tiempo (Min)']/t_total)*100:.1f}%", border=1, align='C', ln=True)

            df_pie = df_f.groupby(pie_col)['Tiempo (Min)'].sum().reset_index()
            fig_pie = px.pie(df_pie, values='Tiempo (Min)', names=pie_col, hole=0.3, title="PROPORCIÓN DE PÉRDIDAS", color_discrete_sequence=px.colors.sequential.Blues_r if area.upper()=="ESTAMPADO" else px.colors.sequential.Oranges_r)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', showlegend=False)
            fig_pie.update_layout(margin=dict(t=35, b=10, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)')
            img_pie = save_chart(fig_pie, w=500, h=200); pdf.image(img_pie, 155, 155, 125); os.remove(img_pie)
        else:
            pdf.set_xy(10, 165); pdf.set_font("Arial", 'I', 10); pdf.set_text_color(100)
            pdf.cell(136, 10, "No hay registros de fallas en este período.", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')


# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO (CALIDAD)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, mes_sel, anio_sel):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    theme_hex = '#%02x%02x%02x' % theme_color
    grupos = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa = {k.upper(): v for k, v in MAQUINAS_MAP.items()}
    
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    pdf.add_page(orientation='L')
    
    df_t = df_trend.copy()
    if not df_t.empty:
        df_t['G'] = df_t['Máquina'].str.upper().map(mapa)
        df_t = df_t[df_t['G'].isin(grupos)]
    
    df_p = df_piezas.copy()
    if not df_p.empty:
        df_p['G'] = df_p['Máquina'].str.upper().map(mapa)
        df_p = df_p[df_p['G'].isin(grupos)]

    # --- ENCABEZADO FORMATO EXCEL ---
    pdf.set_y(10)
    pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
    pdf.cell(20, 6, "MES", 1, 0, 'C', fill=True)
    pdf.cell(20, 6, "AÑO", 1, 0, 'C', fill=True)
    pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C', fill=True)
    pdf.cell(40, 6, "AREA", 1, 1, 'C', fill=True)
    
    pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
    pdf.cell(20, 6, str(mes_sel), 1, 0, 'C')
    pdf.cell(20, 6, str(anio_sel), 1, 0, 'C')
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(197, 6, f"PLANTA: {area.upper()}", 1, 0, 'C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C')

    # --- MARCADORES SUPERIORES (Estético) ---
    pdf.set_fill_color(*theme_color)
    for i, txt in enumerate(["PRODUCCIÓN", "CALIDAD", "TENDENCIA SCRAP", "TENDENCIA RT"]):
        pdf.rect(10 + (i*68.5), 25, 65, 12, 'F')
        pdf.set_xy(10 + (i*68.5), 28); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(255); pdf.cell(65, 5, txt, 0, 1, 'C')
    pdf.set_text_color(0)

    if df_t.empty:
        pdf.set_xy(10, 60); pdf.set_font("Arial", 'I', 12); pdf.set_text_color(100)
        pdf.cell(277, 10, "No hay datos de producción registrados para el período seleccionado.", 0, 1, 'C')
        return pdf.output(dest='S').encode('latin-1')

    # --- Izquierda: Evolución Mensual ---
    df_ev = df_t.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
    df_ev['Totales_Div'] = df_ev['Totales'].apply(lambda x: x if x > 0 else 1)
    df_ev['% Scrap'] = (df_ev['Observadas'] / df_ev['Totales_Div']) * 100
    df_ev['% RT'] = (df_ev['Retrabajo'] / df_ev['Totales_Div']) * 100

    meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
    df_ev['Mes_Str'] = df_ev['Month'].map(meses_map)

    f1 = px.bar(df_ev, x='Mes_Str', y='Totales', title="PIEZAS PRODUCIDAS MES A MES", text_auto='.3s', color_discrete_sequence=[theme_hex])
    f2 = px.bar(df_ev, x='Mes_Str', y='% Scrap', title="% DE SCRAP MES A MES", text_auto='.2f', color_discrete_sequence=[theme_hex])
    f3 = px.bar(df_ev, x='Mes_Str', y='% RT', title="% DE RT MES A MES", text_auto='.2f', color_discrete_sequence=[theme_hex])
    
    for f in [f1, f2, f3]: 
        f.update_layout(margin=dict(l=10, r=10, t=35, b=20), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis=dict(visible=False))
        f.update_traces(textposition="outside", cliponaxis=False, textfont_size=11)

    i1 = save_chart(f1, w=500, h=220); pdf.image(i1, 10, 42, 135); os.remove(i1)
    i2 = save_chart(f2, w=500, h=220); pdf.image(i2, 10, 95, 135); os.remove(i2)
    i3 = save_chart(f3, w=500, h=220); pdf.image(i3, 10, 148, 135); os.remove(i3)

    # --- Derecha: Pareto Top Piezas ---
    if not df_p.empty:
        t_s = df_p.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index().sort_values('Scrap', ascending=True)
        t_rt = df_p.groupby('Pieza')['RT'].sum().nlargest(5).reset_index().sort_values('RT', ascending=True)
        
        f4 = px.bar(t_s, x='Scrap', y='Pieza', orientation='h', title="TOP 5 SCRAP POR PIEZA", text_auto=True, color_discrete_sequence=[theme_hex])
        f5 = px.bar(t_rt, x='RT', y='Pieza', orientation='h', title="TOP 5 RT POR PIEZA", text_auto=True, color_discrete_sequence=[theme_hex])
        
        for f in [f4, f5]: 
            f.update_layout(margin=dict(l=10, r=30, t=40, b=20), plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False), yaxis_title="")
            f.update_traces(textposition="outside", cliponaxis=False, textfont_size=11)

        i4 = save_chart(f4, w=500, h=300); pdf.image(i4, 150, 42, 135); os.remove(i4)
        i5 = save_chart(f5, w=500, h=300); pdf.image(i5, 150, 120, 135); os.remove(i5)

    return pdf.output(dest='S').encode('latin-1')


# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Sistema de Reportes Fumiscor")
st.write("Generador de tableros de Gestión a la Vista e Informes Productivos.")
st.divider()

st.write("### 1. Seleccione el Período")
col_tipo, col_fecha = st.columns(2)

with col_tipo:
    tipo_periodo = st.radio("Tipo de Filtro:", ["Diario", "Semanal", "Mensual"], horizontal=True)

with col_fecha:
    today = pd.to_datetime("today").date()
    if tipo_periodo == "Diario":
        d = st.date_input("Día seleccionado:", today)
        ini = fin = pd.to_datetime(d)
        mes_sel = d.month; anio_sel = d.year
        label_periodo = f"{d.strftime('%d/%m/%Y')}"
        
    elif tipo_periodo == "Semanal":
        d = st.date_input("Seleccione un día de la semana:", today)
        dt = pd.to_datetime(d)
        ini = dt - timedelta(days=dt.weekday())
        fin = ini + timedelta(days=6)
        mes_sel = ini.month; anio_sel = ini.year
        label_periodo = f"Semana {ini.isocalendar()[1]} ({ini.strftime('%d/%m')} - {fin.strftime('%d/%m')})"
        
    elif tipo_periodo == "Mensual":
        c1, c2 = st.columns(2)
        with c1: mes_sel = st.selectbox("Mes", range(1, 13), index=today.month-1)
        with c2: anio_sel = st.selectbox("Año", [2024, 2025, 2026], index=2)
        ini = pd.to_datetime(f"{anio_sel}-{mes_sel}-01")
        fin = pd.to_datetime(f"{anio_sel}-{mes_sel}-{calendar.monthrange(anio_sel, mes_sel)[1]}")
        label_periodo = f"{mes_sel}/{anio_sel}"

# Cargar Datos
df_metrics, df_raw, df_trend, df_piezas = fetch_data_from_db(ini, fin, mes_sel, anio_sel)

st.divider()
st.write("### 2. Descargar Reportes")

col_disp, col_prod = st.columns(2)

with col_disp:
    st.markdown("#### ⚙️ Informe de Disponibilidad (OEE)")
    st.write("Métricas OEE, rendimiento, fallos y gráficos de torta.")
    if st.button("Descargar Disponibilidad ESTAMPADO", use_container_width=True):
        with st.spinner("Generando..."):
            pdf = crear_pdf_gestion_a_la_vista("Estampado", label_periodo, df_metrics, df_raw)
            st.download_button("📥 Bajar PDF Estampado", pdf, "Disponibilidad_Estampado.pdf", use_container_width=True)
            
    if st.button("Descargar Disponibilidad SOLDADURA", use_container_width=True):
        with st.spinner("Generando..."):
            pdf = crear_pdf_gestion_a_la_vista("Soldadura", label_periodo, df_metrics, df_raw)
            st.download_button("📥 Bajar PDF Soldadura", pdf, "Disponibilidad_Soldadura.pdf", use_container_width=True)

with col_prod:
    st.markdown("#### 🏭 Informe Productivo (Calidad)")
    st.write("Tendencia mensual, % de Scrap, RT y Pareto de piezas.")
    if st.button("Descargar Productivo ESTAMPADO", use_container_width=True):
        with st.spinner("Generando..."):
            pdf = crear_pdf_informe_productivo("Estampado", label_periodo, df_trend, df_piezas, mes_sel, anio_sel)
            st.download_button("📥 Bajar PDF Estampado", pdf, "Productivo_Estampado.pdf", use_container_width=True)

    if st.button("Descargar Productivo SOLDADURA", use_container_width=True):
        with st.spinner("Generando..."):
            pdf = crear_pdf_informe_productivo("Soldadura", label_periodo, df_trend, df_piezas, mes_sel, anio_sel)
            st.download_button("📥 Bajar PDF Soldadura", pdf, "Productivo_Soldadura.pdf", use_container_width=True)
