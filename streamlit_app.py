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
st.set_page_config(page_title="Gestión a la Vista - Fumiscor", layout="wide", page_icon="📊")

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
    "Cel1 - Rob13 - RUEDA AUX.": "CELDA SOLDADURA", "Cel2 - Rob1 - ALMOHADON": "CELDA SOLDADURA",
    "Cel3 - Rob14 - HANGERS": "CELDA SOLDADURA", "Cel4 - Rob6 - DOB TORCHA": "CELDA SOLDADURA",
    "Cel5 - Rob4 - Respaldo 60/40": "CELDA SOLDADURA", "HANGERS NISSAN": "CELDA SOLDADURA",
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

# ==========================================
# 2. CONEXIÓN A BASE DE DATOS
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, tipo_periodo, mes=None, anio=None):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d')
        fin_str = fecha_fin.strftime('%Y-%m-%d')

        if tipo_periodo == "Mensual":
            q_metrics = f"""
                SELECT c.Name as Máquina, 
                       SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                       SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada,
                       (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE,
                       (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD,
                       (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD,
                       (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE
                FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId
                WHERE p.Month = {mes} AND p.Year = {anio}
                GROUP BY c.Name
            """
        else:
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

        df_metrics = conn.query(q_metrics)

        q_event = f"""
            SELECT e.Id as Evento_Id, c.Name as Máquina, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], 
                   t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4]
            FROM EVENT_01 e
            LEFT JOIN CELL c ON e.CellId = c.CellId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        df_raw = conn.query(q_event)

        if not df_raw.empty:
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)

            def categorizar_estado(row):
                texto_completo = f"{row.get('Nivel Evento 1','')} {row.get('Nivel Evento 2','')} {row.get('Nivel Evento 3','')} {row.get('Nivel Evento 4','')} ".upper()
                if 'PRODUCCION' in texto_completo or 'PRODUCCIÓN' in texto_completo: return 'Producción'
                if 'PROYECTO' in texto_completo: return 'Proyecto'
                if 'BAÑO' in texto_completo or 'BANO' in texto_completo or 'REFRIGERIO' in texto_completo: return 'Descanso'
                if 'PARADA PROGRAMADA' in texto_completo: return 'Parada Programada'
                return 'Falla/Gestión'

            def clasificar_macro(row):
                n1 = str(row.get('Nivel Evento 1', '')).strip().upper()
                n2 = str(row.get('Nivel Evento 2', '')).strip().upper()
                if 'GESTION' in n1 or 'GESTIÓN' in n1: return 'Gestión'
                if 'FALLA' in n1: return n2.title() if n2 not in ['NAN', 'NONE', ''] else 'Falla (Sin área)'
                return n1.title() if n1 not in ['NAN', 'NONE', ''] else 'Sin Clasificar'

            df_raw['Estado_Global'] = df_raw.apply(categorizar_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(clasificar_macro, axis=1)

            def obtener_ultimo_nivel(row):
                niveles = [str(row.get(col, '')).strip() for col in ['Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4']]
                validos = [n for n in niveles if n.lower() not in ['none', 'nan', '', 'null']]
                if not validos: return "Sin detalle en sistema"
                ultimo = validos[-1]; macro = row['Categoria_Macro']
                if row['Estado_Global'] == 'Falla/Gestión' and macro.upper() not in ultimo.upper(): 
                    return f"[{macro}] {ultimo}"
                return ultimo

            df_raw['Detalle_Final'] = df_raw.apply(obtener_ultimo_nivel, axis=1)

        return df_metrics, df_raw

    except Exception as e:
        st.error(f"Error ejecutando consulta a base de datos wii_bi: {e}")
        return pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR DEL DASHBOARD VISUAL (PDF)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw, mes_str, anio_str):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    grupos_area = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa_limpio = {str(k).strip().upper(): v for k, v in MAQUINAS_MAP.items()}

    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    
    # Preparar datos base para el área
    df_met_area = df_metrics_pdf.copy()
    if not df_met_area.empty and df_met_area['OEE'].max() > 1.5:
        df_met_area[['OEE', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD']] /= 100.0

    df_met_area['Grupo'] = df_met_area['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('Otro')
    df_met_area = df_met_area[df_met_area['Grupo'].isin(grupos_area)]
    
    df_raw_area = df_pdf_raw.copy()
    df_raw_area['Grupo_Máquina'] = df_raw_area['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('Otro')
    df_raw_area = df_raw_area[df_raw_area['Grupo_Máquina'].isin(grupos_area)]

    # Determinar qué páginas generar
    paginas_a_generar = ['GENERAL'] + [g for g in grupos_area if g in df_met_area['Grupo'].unique()]

    for target in paginas_a_generar:
        pdf.add_page(orientation='L') # Formato Apaisado (Landscape)
        
        if target == 'GENERAL':
            df_m = df_met_area.copy()
            df_r = df_raw_area.copy()
            x_barras = 'Grupo'
            pie_col = 'Grupo_Máquina'
            titulo_planta = f"PLANTA {area.upper()} - CONSOLIDADO GENERAL"
        else:
            df_m = df_met_area[df_met_area['Grupo'] == target].copy()
            df_r = df_raw_area[df_raw_area['Grupo_Máquina'] == target].copy()
            x_barras = 'Máquina'
            pie_col = 'Categoria_Macro'
            titulo_planta = f"PLANTA {area.upper()} - ÁREA: {target}"

        # --- ENCABEZADO TIPO EXCEL ---
        pdf.set_y(10)
        pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0)
        
        pdf.cell(20, 6, "MES", 1, 0, 'C'); pdf.cell(20, 6, "AÑO", 1, 0, 'C')
        pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C')
        pdf.cell(40, 6, "AREA", 1, 1, 'C')
        
        pdf.set_font("Arial", '', 10)
        pdf.cell(20, 6, str(mes_str) if mes_str else "-", 1, 0, 'C')
        pdf.cell(20, 6, str(anio_str) if anio_str else "-", 1, 0, 'C')
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(197, 6, clean_text(titulo_planta), 1, 0, 'C')
        pdf.set_font("Arial", '', 10)
        pdf.cell(40, 6, str(target), 1, 1, 'C')

        # --- KPI BOXES SUPERIORES ---
        t_plan = df_m['T_Operativo'].sum() + df_m['T_Parada'].sum()
        t_op = df_m['T_Operativo'].sum()
        p_totales = df_m['Buenas'].sum() + df_m['Retrabajo'].sum() + df_m['Observadas'].sum() if not df_m.empty else 0

        kpis = {
            "OEE": (df_m['OEE'] * (df_m['T_Operativo'] + df_m['T_Parada'])).sum() / t_plan if t_plan > 0 else 0,
            "PERFORMANCE": (df_m['PERFORMANCE'] * df_m['T_Operativo']).sum() / t_op if t_op > 0 else 0,
            "DISPONIBILIDAD": (df_m['DISPONIBILIDAD'] * (df_m['T_Operativo'] + df_m['T_Parada'])).sum() / t_plan if t_plan > 0 else 0,
            "CALIDAD": df_m['Buenas'].sum() / p_totales if p_totales > 0 else 0
        }

        y_kpi = 25; w_box = 65; spacing = 3.5; x_start = 10
        for i, (label, val) in enumerate(kpis.items()):
            x = x_start + (i * (w_box + spacing))
            pdf.set_fill_color(*theme_color)
            pdf.rect(x, y_kpi, w_box, 20, 'F')
            pdf.set_xy(x, y_kpi + 2)
            pdf.set_font("Arial", 'B', 10); pdf.set_text_color(255)
            pdf.cell(w_box, 6, label, 0, 1, 'L')
            pdf.set_xy(x, y_kpi + 8)
            pdf.set_font("Arial", 'B', 20)
            pdf.cell(w_box, 10, f"{val*100:.1f}%", 0, 0, 'C')

        pdf.set_text_color(0)

        # --- GENERADOR DE GRÁFICOS DE BARRAS ---
        def add_bar_chart(df, y_col, title, x_pos, y_pos):
            if df.empty: return
            if y_col in ['OEE', 'DISPONIBILIDAD']: df['Peso'] = df['T_Operativo'] + df['T_Parada']
            elif y_col == 'PERFORMANCE': df['Peso'] = df['T_Operativo']
            else: df['Peso'] = df['Buenas'] + df['Retrabajo'] + df['Observadas']
                
            df['Num'] = df[y_col] * df['Peso']
            df_g = df.groupby(x_barras)[['Num', 'Peso']].sum().reset_index()
            df_g['Val'] = (df_g['Num'] / df_g['Peso']) * 100
            
            fig = px.bar(df_g, x=x_barras, y='Val', title=title, text_auto='.1f', color_discrete_sequence=[px.colors.qualitative.Prism[0]])
            fig.update_layout(height=220, width=600, margin=dict(t=35, b=20, l=20, r=20), plot_bgcolor='rgba(0,0,0,0)', 
                              xaxis_title="", yaxis_title="%", yaxis=dict(range=[0, 110]), title_font_size=14)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                fig.write_image(tmp.name, engine="kaleido")
                pdf.image(tmp.name, x=x_pos, y=y_pos, w=135)
                os.remove(tmp.name)

        # ROW 1 (OEE / PERFORMANCE)
        pdf.set_draw_color(0); pdf.rect(10, 48, 136, 52); pdf.rect(149, 48, 138, 52)
        add_bar_chart(df_m, 'OEE', 'GRÁFICO DE BARRAS OEE', 10, 49)
        add_bar_chart(df_m, 'PERFORMANCE', 'GRÁFICO DE BARRAS PERFO', 150, 49)

        # ROW 2 (DISPONIBILIDAD / CALIDAD)
        pdf.rect(10, 102, 136, 52); pdf.rect(149, 102, 138, 52)
        add_bar_chart(df_m, 'DISPONIBILIDAD', 'GRÁFICO DE BARRAS DISPO', 10, 103)
        add_bar_chart(df_m, 'CALIDAD', 'GRÁFICO DE BARRAS CALIDAD', 150, 103)

        # --- ROW 3: TABLA TOP 5 Y GRÁFICO DE TORTA ---
        y_bottom = 156
        pdf.rect(10, y_bottom, 136, 45); pdf.rect(149, y_bottom, 138, 45)
        
        # Tabla Top 5
        pdf.set_xy(10, y_bottom)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(136, 6, "TOP 5 FALLOS", border='B', ln=True, align='C')
        
        df_f = df_r[df_r['Estado_Global'] == 'Falla/Gestión']
        t_falla_total = df_f['Tiempo (Min)'].sum() if not df_f.empty else 1
        
        pdf.set_xy(10, y_bottom + 6)
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(56, 5, "FALLO", border=1, align='C')
        pdf.cell(40, 5, "CATEGORIA", border=1, align='C')
        pdf.cell(40, 5, "% DEL TOTAL DE FALLOS", border=1, align='C', ln=True)
        
        pdf.set_font("Arial", '', 8)
        if not df_f.empty:
            top5 = df_f.groupby(['Detalle_Final', 'Categoria_Macro'])['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(5)
            for _, r in top5.iterrows():
                pdf.set_x(10)
                pdf.cell(56, 6, clean_text(str(r['Detalle_Final']))[:35], border=1)
                pdf.cell(40, 6, clean_text(str(r['Categoria_Macro']))[:22], border=1)
                pdf.cell(40, 6, f"{(r['Tiempo (Min)'] / t_falla_total)*100:.1f}%", border=1, align='C', ln=True)

        # Gráfico Torta
        if not df_f.empty:
            df_pie = df_f.groupby(pie_col)['Tiempo (Min)'].sum().reset_index()
            fig_pie = px.pie(df_pie, values='Tiempo (Min)', names=pie_col, hole=0.3, title="PROPORCIÓN DE FALLO POR ÁREA")
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', showlegend=False)
            fig_pie.update_layout(height=200, width=500, margin=dict(t=35, b=10, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', title_font_size=14, title_x=0.5)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                fig_pie.write_image(tmp.name, engine="kaleido")
                pdf.image(tmp.name, x=155, y=y_bottom + 2, w=125)
                os.remove(tmp.name)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. INTERFAZ DE USUARIO STREAMLIT
# ==========================================
st.markdown('<div class="header-style">📊 Tableros de Gestión a la Vista</div>', unsafe_allow_html=True)
st.write("Generador de tableros de planta listos para imprimir en formato A4 Apaisado.")
st.divider()

col_p1, col_p2, col_p3 = st.columns([1, 1.2, 1.5])

with col_p1:
    st.write("**1. Tipo de Reporte:**")
    pdf_tipo = st.radio("Período:", ["Diario", "Semanal", "Mensual"], horizontal=True, label_visibility="collapsed")

with col_p2:
    st.write("**2. Seleccione el Período:**")
    today = pd.to_datetime("today").date()
    pdf_ini, pdf_fin, pdf_mes, pdf_anio = None, None, None, None
    pdf_label, file_label = "", ""

    if pdf_tipo == "Diario":
        pdf_fecha = st.date_input("Día para Tablero:", value=today)
        pdf_ini = pdf_fin = pd.to_datetime(pdf_fecha)
        pdf_label = f"Dia {pdf_fecha.strftime('%d-%m-%Y')}"
        file_label = pdf_label
        
    elif pdf_tipo == "Semanal":
        fecha_ref = st.date_input("Día de la semana deseada:", value=today)
        dt_ref = pd.to_datetime(fecha_ref)
        pdf_ini = dt_ref - timedelta(days=dt_ref.weekday()); pdf_fin = pdf_ini + timedelta(days=6) 
        semana_num = pdf_ini.isocalendar().week
        pdf_label = f"Semana {semana_num} ({pdf_ini.strftime('%d/%m')} al {pdf_fin.strftime('%d/%m')})"
        file_label = f"Semana_{semana_num}"
        
    elif pdf_tipo == "Mensual":
        c_m, c_y = st.columns(2)
        mes_list = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        with c_m: mes_sel = st.selectbox("Mes", mes_list, index=today.month-1)
        with c_y: anio_sel = st.selectbox("Año", range(2023, today.year + 2), index=today.year-2023)
        pdf_mes = mes_list.index(mes_sel) + 1; pdf_anio = anio_sel
        pdf_ini = pd.to_datetime(f"{pdf_anio}-{pdf_mes}-01")
        last_day = calendar.monthrange(pdf_anio, pdf_mes)[1]
        pdf_fin = pd.to_datetime(f"{pdf_anio}-{pdf_mes}-{last_day}")
        pdf_label = f"{mes_sel} {pdf_anio}"; file_label = f"{mes_sel}_{pdf_anio}"

# Cargar Datos
df_metrics, df_raw = fetch_data_from_db(pdf_ini, pdf_fin, pdf_tipo, mes=pdf_mes, anio=pdf_anio)

with col_p3:
    st.write("**3. Descargar Tableros PDF:**")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("Tablero ESTAMPADO", use_container_width=True):
            with st.spinner("Generando Tableros Estampado..."):
                try:
                    mes_desc = pdf_mes if pdf_tipo == "Mensual" else pdf_ini.month
                    anio_desc = pdf_anio if pdf_tipo == "Mensual" else pdf_ini.year
                    pdf_visual = crear_pdf_gestion_a_la_vista("Estampado", pdf_label, df_metrics, df_raw, mes_desc, anio_desc)
                    st.download_button("Descargar Estampado", data=pdf_visual, file_name=f"Tableros_Estampado_{file_label}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando Visual: {e}")

    with col_btn2:
        if st.button("Tablero SOLDADURA", use_container_width=True):
            with st.spinner("Generando Tableros Soldadura..."):
                try:
                    mes_desc = pdf_mes if pdf_tipo == "Mensual" else pdf_ini.month
                    anio_desc = pdf_anio if pdf_tipo == "Mensual" else pdf_ini.year
                    pdf_visual = crear_pdf_gestion_a_la_vista("Soldadura", pdf_label, df_metrics, df_raw, mes_desc, anio_desc)
                    st.download_button("Descargar Soldadura", data=pdf_visual, file_name=f"Tableros_Soldadura_{file_label}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando Visual: {e}")
