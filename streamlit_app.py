import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

    def add_gradient_background(self):
        r1, g1, b1 = 240, 242, 246
        r2, g2, b2 = 215, 220, 225
        h = self.h; w = self.w
        for i in range(int(h * 2)):
            ratio = i / (h * 2)
            r = int(r1 + (r2 - r1) * ratio); g = int(g1 + (g2 - g1) * ratio); b = int(b1 + (b2 - b1) * ratio)
            self.set_fill_color(r, g, b)
            self.rect(0, i / 2, w, 0.5, 'F')

    def rounded_rect(self, x, y, w, h, r, style=''):
        k = self.k; hp = self.h
        op = 'f' if style == 'F' else 'B' if style in ['FD', 'DF'] else 'S'
        MyArc = 4/3 * ((2 ** 0.5) - 1)
        self._out(f'{(x + r) * k:.2f} {(hp - y) * k:.2f} m')
        xc = x + w - r; yc = y + r
        self._out(f'{xc * k:.2f} {(hp - y) * k:.2f} l')
        self._out(f'{(xc + r * MyArc) * k:.2f} {(hp - y) * k:.2f} {(x + w) * k:.2f} {(hp - yc + r * MyArc) * k:.2f} {(x + w) * k:.2f} {(hp - yc) * k:.2f} c')
        yc = y + h - r
        self._out(f'{(x + w) * k:.2f} {(hp - yc) * k:.2f} l')
        self._out(f'{(x + w) * k:.2f} {(hp - yc - r * MyArc) * k:.2f} {(xc + r * MyArc) * k:.2f} {(hp - y - h) * k:.2f} {xc * k:.2f} {(hp - y - h) * k:.2f} c')
        xc = x + r
        self._out(f'{xc * k:.2f} {(hp - y - h) * k:.2f} l')
        self._out(f'{(xc - r * MyArc) * k:.2f} {(hp - y - h) * k:.2f} {x * k:.2f} {(hp - yc - r * MyArc) * k:.2f} {x * k:.2f} {(hp - yc) * k:.2f} c')
        yc = y + r
        self._out(f'{x * k:.2f} {(hp - yc) * k:.2f} l')
        self._out(f'{x * k:.2f} {(hp - yc + r * MyArc) * k:.2f} {(xc - r * MyArc) * k:.2f} {(hp - y) * k:.2f} {xc * k:.2f} {(hp - y) * k:.2f} c')
        self._out(op)

    def draw_panel(self, x, y, w, h, r=3, bg_color=(255,255,255)):
        self.set_fill_color(210, 210, 210); self.rounded_rect(x + 1.5, y + 1.5, w, h, r, style='F')
        self.set_fill_color(*bg_color); self.set_draw_color(180, 180, 180)
        self.rounded_rect(x, y, w, h, r, style='DF')

    def draw_kpi_panel(self, x, y, w, h, r=3):
        self.set_fill_color(200, 200, 200); self.rounded_rect(x + 1.5, y + 1.5, w, h, r, style='F')
        self.set_fill_color(*self.theme_color); self.rounded_rect(x, y, w, h, r, style='F')

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

def save_chart(fig, w=600, h=300):
    fig.update_layout(width=w, height=h, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig.write_image(tmp.name, engine="kaleido", scale=2.5); return tmp.name

# ==========================================
# 2. CARGA DE DATOS (CORRECCIÓN WIIDEM)
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d'); fin_str = fecha_fin.strftime('%Y-%m-%d')
        
        q_metrics = f"SELECT c.Name as Máquina, SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas, SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada, (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE, (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD, (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD, (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE FROM PROD_D_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}' GROUP BY c.Name"
        q_event = f"SELECT c.Name as Máquina, e.Interval as [Tiempo (Min)], t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4] FROM EVENT_01 e LEFT JOIN CELL c ON e.CellId = c.CellId LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'"
        q_trend_oee = f"SELECT p.Month, c.Name as Máquina, SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada, SUM(p.ProductiveTime + p.DownTime) as T_Planificado, SUM(p.Performance * p.ProductiveTime) as Perf_Num, SUM(p.Availability * (p.ProductiveTime + p.DownTime)) as Disp_Num, SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) as Cal_Num, SUM(p.Oee * (p.ProductiveTime + p.DownTime)) as OEE_Num FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY p.Month, c.Name"
        # OJO: Se usa PROD_M_01 para alinear las cantidades de producción con los datos reales de Wiidem (Evita duplicidad)
        q_trend_piezas = f"SELECT p.Month, c.Name as Máquina, SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas, SUM(p.Good + p.Rework + p.Scrap) as Totales FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId WHERE p.Year = {anio} AND p.Month <= {mes} GROUP BY p.Month, c.Name"
        q_piezas = f"SELECT c.Name as Máquina, pr.Code as Pieza, SUM(p.Scrap) as Scrap, SUM(p.Rework) as RT FROM PROD_D_01 p JOIN CELL c ON p.CellId = c.CellId JOIN PRODUCT pr ON p.ProductId = pr.ProductId WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}' GROUP BY c.Name, pr.Code"

        df_metrics = conn.query(q_metrics)
        df_raw = conn.query(q_event)
        df_trend_oee = conn.query(q_trend_oee)
        df_trend_piezas = conn.query(q_trend_piezas)
        df_piezas = conn.query(q_piezas)

        if not df_trend_oee.empty and not df_trend_piezas.empty:
            df_trend = pd.merge(df_trend_piezas, df_trend_oee, on=['Month', 'Máquina'], how='outer').fillna(0)
        else:
            df_trend = df_trend_piezas if not df_trend_piezas.empty else df_trend_oee

        if df_raw.empty: 
            df_raw = pd.DataFrame(columns=['Máquina', 'Tiempo (Min)', 'Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4', 'Estado_Global', 'Categoria_Macro', 'Detalle_Final'])
        else:
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)
            
            # PURGA ABSOLUTA DE PROYECTO (Elimina fila completa si 'PROYECTO' aparece en algún nivel)
            for col in ['Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4']:
                df_raw[col] = df_raw[col].fillna('')
                df_raw = df_raw[~df_raw[col].str.upper().str.contains('PROYECTO')]

            # ESTADO GLOBAL: Se toma exclusivamente del Nivel Evento 1
            def cat_estado(row):
                t1 = str(row.get('Nivel Evento 1','')).strip().upper()
                if 'PRODUCCION' in t1 or 'PRODUCCIÓN' in t1: return 'Producción'
                if 'PARADA PROGRAMADA' in t1: return 'Parada Programada'
                return 'Falla/Gestión'
            
            # CATEGORÍA MACRO (ÁREAS): Se toma exclusivamente del Nivel Evento 2 (Mantenimiento, Calidad, Gestion, etc.)
            def cat_macro(row):
                n2 = str(row.get('Nivel Evento 2', '')).strip().title()
                return n2 if n2 and n2.lower() not in ['nan', 'none', 'null', ''] else "Sin Área"
            
            # DETALLE FINAL (FALLA PUNTUAL): Se toma exclusivamente del Nivel Evento 3 (Baño, Refrigerio, Fallas de Robot)
            def get_det(row):
                n3 = str(row.get('Nivel Evento 3', '')).strip()
                if n3 and n3.lower() not in ['nan', 'none', 'null', '']:
                    return n3
                # Si el operador omitió el Nivel 3, usamos el Nivel 2 como fallback para dejar la mala carga en evidencia
                n2 = str(row.get('Nivel Evento 2', '')).strip()
                if n2 and n2.lower() not in ['nan', 'none', 'null', '']:
                    return n2
                return "Sin detalle"
                
            df_raw['Estado_Global'] = df_raw.apply(cat_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(cat_macro, axis=1)
            df_raw['Detalle_Final'] = df_raw.apply(get_det, axis=1)

        return df_metrics, df_raw, df_trend, df_piezas
    except Exception as e: 
        st.error(f"Error SQL: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (DISPONIBILIDAD)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw, df_trend):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0); theme_hex = '#%02x%02x%02x' % theme_color
    grupos_area = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    mapa_limpio = {str(k).strip().upper(): v for k, v in MAQUINAS_MAP.items()}
    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    
    df_m = df_metrics_pdf.copy(); df_t = df_trend.copy(); df_r = df_pdf_raw.copy()
    if not df_m.empty and df_m['OEE'].max() > 1.5: df_m[['OEE', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD']] /= 100.0
    for d in [df_m, df_t, df_r]: 
        if not d.empty: d['Grupo'] = d['Máquina'].str.strip().str.upper().map(mapa_limpio).fillna('Otro')
    paginas = ['GENERAL'] + [g for g in grupos_area if not df_m.empty and g in df_m['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        df_m_target = df_m if target == 'GENERAL' else df_m[df_m['Grupo'] == target]
        df_t_target = df_t if target == 'GENERAL' else df_t[df_t['Grupo'] == target]
        df_r_target = df_r if target == 'GENERAL' else df_r[df_r['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C', fill=True)

        t_plan = df_m_target['T_Operativo'].sum() + df_m_target['T_Parada'].sum() if not df_m_target.empty else 0
        t_op = df_m_target['T_Operativo'].sum() if not df_m_target.empty else 0
        kpis = {"OEE": (df_m_target['OEE'] * (df_m_target['T_Operativo'] + df_m_target['T_Parada'])).sum() / t_plan if t_plan > 0 else 0, "PERFORMANCE": (df_m_target['PERFORMANCE'] * df_m_target['T_Operativo']).sum() / t_op if t_op > 0 else 0, "DISPONIBILIDAD": (df_m_target['DISPONIBILIDAD'] * (df_m_target['T_Operativo'] + df_m_target['T_Parada'])).sum() / t_plan if t_plan > 0 else 0, "CALIDAD": df_m_target['CALIDAD'].mean() if not df_m_target.empty else 0}
        
        for i, (lbl, val) in enumerate(kpis.items()):
            x = 10 + (i * 68.5); pdf.draw_kpi_panel(x, y_kpi:=25, 65, 20)
            pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(255); pdf.cell(65, 6, lbl, 0, 1, 'L')
            pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{val*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        def add_trend_bar(df_in, col, title, x_pos, y_pos):
            if df_in.empty: return
            if col == 'OEE': df_g = df_in.groupby('Month')[['OEE_Num', 'T_Planificado']].sum().reset_index(); df_g['Val'] = df_g['OEE_Num'] / df_g['T_Planificado'].replace(0, 1)
            elif col == 'PERFORMANCE': df_g = df_in.groupby('Month')[['Perf_Num', 'T_Operativo']].sum().reset_index(); df_g['Val'] = df_g['Perf_Num'] / df_g['T_Operativo'].replace(0, 1)
            elif col == 'DISPONIBILIDAD': df_g = df_in.groupby('Month')[['Disp_Num', 'T_Planificado']].sum().reset_index(); df_g['Val'] = df_g['Disp_Num'] / df_g['T_Planificado'].replace(0, 1)
            elif col == 'CALIDAD': df_g = df_in.groupby('Month')[['Cal_Num', 'Totales']].sum().reset_index(); df_g['Val'] = df_g['Cal_Num'] / df_g['Totales'].replace(0, 1)
            if df_g['Val'].max() > 1.5: df_g['Val'] /= 100.0
            
            df_g['Mes_Str'] = df_g['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
            def get_c(v): return '#E74C3C' if v < 0.75 else ('#F1C40F' if v <= 0.85 else '#2ECC71')
            df_g['Color'] = df_g['Val'].apply(get_c)
            fig = go.Figure(data=[go.Bar(x=df_g['Mes_Str'], y=df_g['Val'], marker=dict(color=df_g['Color'], line=dict(color='rgba(0,0,0,0.8)', width=2)), text=df_g['Val'], texttemplate='<b>%{text:.1%}</b>', textposition='outside', opacity=0.85)])
            fig.add_hline(y=0.75, line_dash="dash", line_color="#E74C3C", annotation_text="<b>75%</b>", annotation_font_color='black')
            fig.add_hline(y=0.85, line_dash="dash", line_color="#2ECC71", annotation_text="<b>85%</b>", annotation_font_color='black')
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(t=30, b=20, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, max(1.1, df_g['Val'].max()*1.3)], visible=False), xaxis_title="")
            fig.update_traces(textfont=dict(color='black', size=11, family="Arial"), cliponaxis=False)
            img = save_chart(fig, 600, 220); pdf.image(img, x=x_pos+2, y=y_pos+2, w=132); os.remove(img)

        pdf.draw_panel(10, 48, 136, 52); pdf.draw_panel(149, 48, 138, 52)
        add_trend_bar(df_t_target, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48)
        add_trend_bar(df_t_target, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48)
        pdf.draw_panel(10, 102, 136, 52); pdf.draw_panel(149, 102, 138, 52)
        add_trend_bar(df_t_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 102)
        add_trend_bar(df_t_target, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 102)
        
        pdf.draw_panel(10, 156, 136, 45); pdf.draw_panel(149, 156, 138, 45)
        pdf.set_xy(10, 156); pdf.set_font("Times", 'B', 11); pdf.set_text_color(0); pdf.cell(136, 6, "TOP 5 FALLOS", border=0, ln=True, align='C')
        
        df_f = df_r_target[df_r_target['Estado_Global'] == 'Falla/Gestión'] if not df_r_target.empty else pd.DataFrame()
        
        if not df_f.empty and df_f['Tiempo (Min)'].sum() > 0:
            # FILTRO TOP 5: Excluir Baño, Refrigerio y Descanso (Para que solo se vean fallas reales de proceso en la tabla)
            df_f_puras = df_f[~df_f['Detalle_Final'].str.upper().str.contains('BAÑO|BANO|REFRIGERIO|DESCANSO', na=False)]
            top5 = df_f_puras.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
            
            pdf.set_xy(10, 162); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
            pdf.cell(76, 5, "FALLO", border=1, fill=True); pdf.cell(30, 5, "MINUTOS", border=1, align='C', fill=True); pdf.cell(30, 5, "% TOTAL", border=1, align='C', ln=True, fill=True)
            pdf.set_font("Arial", '', 8); pdf.set_text_color(0); pdf.set_fill_color(255, 255, 255)
            
            t_total = df_f['Tiempo (Min)'].sum()
            for _, r in top5.iterrows():
                pdf.set_x(10); pdf.cell(76, 6, clean_text(str(r['Detalle_Final']))[:45], border=1, fill=True)
                pdf.cell(30, 6, f"{r['Tiempo (Min)']:.0f}", border=1, align='C', fill=True)
                pdf.cell(30, 6, f"{(r['Tiempo (Min)']/t_total)*100:.1f}%", border=1, align='C', ln=True, fill=True)
            
            # Gráfico de Barras Apiladas (Macro) - ESTE GRÁFICO TOMA LA COLUMNA 'Categoria_Macro' Y MUESTRA TODAS LAS ÁREAS (Incluido Gestión y Matricería)
            df_macro = df_f.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
            df_macro['%'] = df_macro['Tiempo (Min)'] / t_total
            df_macro['Y'] = "Pérdidas"
            
            fig_stack = px.bar(df_macro, x='%', y='Y', color='Categoria_Macro', orientation='h', color_discrete_sequence=px.colors.qualitative.Safe)
            fig_stack.update_traces(texttemplate='<b>%{x:.1%}</b>', textposition='inside', marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.9, textfont=dict(color='black', size=11))
            fig_stack.update_layout(barmode='stack', title=dict(text="<b>PROPORCIÓN DE PÉRDIDAS ÁREAS MACRO (100%)</b>", font=dict(family="Times", size=13, color="black")), xaxis=dict(visible=False, range=[0, 1]), yaxis=dict(visible=False), margin=dict(t=30, b=5, l=10, r=10), legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5, title="", font=dict(size=10)))
            img_stack = save_chart(fig_stack, 600, 180); pdf.image(img_stack, 151, 158, 134); os.remove(img_stack)
            
    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO (CALIDAD)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, mes_sel, anio_sel, hs_rt):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0); theme_hex = '#%02x%02x%02x' % theme_color
    scrap_c = '#002147' if area.upper() == "ESTAMPADO" else '#722F37' # Azul Navy / Bordeaux
    rt_c = theme_hex
    grupos = GRUPOS_ESTAMPADO if area.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    
    df_t = df_trend.copy(); df_p = df_piezas.copy(); mapa = {k.upper(): v for k, v in MAQUINAS_MAP.items()}
    for d in [df_t, df_p]: 
        if not d.empty: d['Grupo'] = d['Máquina'].str.strip().str.upper().map(mapa).fillna('Otro')
        
    paginas = ['GENERAL'] + [g for g in grupos if not df_t.empty and g in df_t['Grupo'].unique()]

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        df_t_target = df_t if target == 'GENERAL' else df_t[df_t['Grupo'] == target]
        df_p_target = df_p if target == 'GENERAL' else df_p[df_p['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(20, 6, "MES", 1, 0, 'C', fill=True); pdf.cell(20, 6, "AÑO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "AREA", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(20, 6, str(mes_sel), 1, 0, 'C', fill=True); pdf.cell(20, 6, str(anio_sel), 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C', fill=True)

        if df_t_target.empty: continue
        
        df_ev = df_t_target.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
        df_ev['Totales_Div'] = df_ev['Totales'].apply(lambda x: x if x > 0 else 1)
        df_ev['% Scrap'] = ((df_ev['Observadas'] / df_ev['Totales_Div']) * 100).round(2)
        df_ev['% RT'] = ((df_ev['Retrabajo'] / df_ev['Totales_Div']) * 100).round(2)
        meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
        df_ev['Mes_Str'] = df_ev['Month'].map(meses_map)

        f1 = px.bar(df_ev, x='Mes_Str', y='Totales', color_discrete_sequence=[theme_hex]); f1.update_traces(texttemplate='<b>%{y:.3s}</b>')
        f2 = px.bar(df_ev, x='Mes_Str', y='% Scrap', color_discrete_sequence=[theme_hex]); f2.update_traces(texttemplate='<b>%{y:.2f}%</b>')
        f3 = px.bar(df_ev, x='Mes_Str', y='% RT', color_discrete_sequence=[theme_hex]); f3.update_traces(texttemplate='<b>%{y:.2f}%</b>')
        
        titles = ["PIEZAS PRODUCIDAS MES A MES", "% DE SCRAP MES A MES", "% DE RT MES A MES"]
        for i, f in enumerate([f1, f2, f3]): 
            max_y = df_ev['Totales'].max() if i==0 else (df_ev['% Scrap'].max() if i==1 else df_ev['% RT'].max())
            if i == 0: upper_limit = max_y * 1.3 if max_y > 0 else 1
            else: upper_limit = 0.2 
            f.update_yaxes(range=[0, upper_limit])
            f.update_layout(title=dict(text=f"<b>{titles[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=10, t=30, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis=dict(visible=False))
            f.update_traces(textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

        h_box = 60; pdf.draw_panel(10, 22, 135, h_box); pdf.draw_panel(10, 85, 135, h_box); pdf.draw_panel(10, 148, 135, h_box)
        for i, f in enumerate([f1, f2, f3]): pdf.image(save_chart(f, 550, 260), 11, 23+(i*63), 133, h_box-2)

        h_br = 83.5; pdf.draw_panel(150, 22, 135, h_br); pdf.draw_panel(150, 108.5, 135, h_br)
        if not df_p_target.empty:
            t_s = df_p_target.groupby('Pieza')['Scrap'].sum().nlargest(5).reset_index().sort_values('Scrap', ascending=True)
            t_rt = df_p_target.groupby('Pieza')['RT'].sum().nlargest(5).reset_index().sort_values('RT', ascending=True)
            
            f4 = px.bar(t_s, x='Scrap', y='Pieza', orientation='h', color_discrete_sequence=[scrap_c])
            f5 = px.bar(t_rt, x='RT', y='Pieza', orientation='h', color_discrete_sequence=[rt_c])
            
            titles_right = ["TOP 5 SCRAP POR PIEZA", "TOP 5 RT POR PIEZA"]
            for i, f in enumerate([f4, f5]):
                max_x = t_s['Scrap'].max() if i==0 else t_rt['RT'].max()
                upper_limit = max_x * 1.3 if max_x > 0 else 1
                f.update_xaxes(range=[0, upper_limit])
                f.update_layout(title=dict(text=f"<b>{titles_right[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=30, t=35, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False), yaxis=dict(title="", automargin=True, tickfont=dict(color='black', size=10)))
                f.update_traces(texttemplate='<b>%{x}</b>', textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

            i4 = save_chart(f4, w=550, h=330); pdf.image(i4, 151, 23, w=133, h_br-2); os.remove(i4)
            i5 = save_chart(f5, w=550, h=330); pdf.image(i5, 151, 109.5, w=133, h_br-2); os.remove(i5)
            
        pdf.draw_panel(150, 196, 135, 12, 2, (240,240,240)); pdf.set_xy(150, 196); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0); pdf.cell(67.5, 12, "HS DE RT", 0, 0, 'C')
        pdf.draw_panel(217.5, 196, 67.5, 12, 2, (255,255,255)); pdf.set_xy(217.5, 196); pdf.cell(67.5, 12, f"{hs_rt:.1f}", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Reportes Fumiscor")
st.divider()
st.write("### 1. Seleccione el Período")

col_tipo, col_fecha = st.columns(2)
with col_tipo: 
    tipo_periodo = st.radio("Filtro:", ["Diario", "Semanal", "Mensual"], horizontal=True)

with col_fecha:
    today = pd.to_datetime("today").date()
    if tipo_periodo == "Diario": 
        d = st.date_input("Día:", today)
        ini = fin = pd.to_datetime(d)
        m_sel = d.month
        a_sel = d.year
        lab = d.strftime('%d/%m/%Y')
        
    elif tipo_periodo == "Semanal": 
        d = st.date_input("Día de la semana:", today)
        dt = pd.to_datetime(d)
        ini = dt - timedelta(days=dt.weekday())
        fin = ini + timedelta(days=6)
        m_sel = ini.month
        a_sel = ini.year
        lab = f"Sem {ini.isocalendar()[1]} ({ini.strftime('%d/%m')}-{fin.strftime('%d/%m')})"
        
    elif tipo_periodo == "Mensual": 
        c1, c2 = st.columns(2)
        with c1: 
            m_sel = st.selectbox("Mes", range(1, 13), index=today.month-1)
        with c2: 
            a_sel = st.selectbox("Año", [2024, 2025, 2026], index=2)
        
        ini = pd.to_datetime(f"{a_sel}-{m_sel}-01")
        fin = pd.to_datetime(f"{a_sel}-{m_sel}-{calendar.monthrange(a_sel, m_sel)[1]}")
        lab = f"{m_sel}/{a_sel}"

df_m, df_r, df_t, df_p = fetch_data_from_db(ini, fin, m_sel, a_sel)
st.write("### 2. Datos Manuales (Informe Productivo)")
hs_rt = st.number_input("Horas de RT:", 0.0, 1000.0, 0.0, 1.0)

st.divider()
st.write("### 3. Descargar Reportes")
c_d, c_p = st.columns(2)

with c_d:
    st.markdown("#### ⚙️ Informe de Disponibilidad (OEE)")
    if st.button("Disponibilidad ESTAMPADO", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Estampado", crear_pdf_gestion_a_la_vista("Estampado", lab, df_m, df_r, df_t), "Disp_Estampado.pdf")
    if st.button("Disponibilidad SOLDADURA", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Soldadura", crear_pdf_gestion_a_la_vista("Soldadura", lab, df_m, df_r, df_t), "Disp_Soldadura.pdf")

with c_p:
    st.markdown("#### 🏭 Informe Productivo (Calidad)")
    if st.button("Productivo ESTAMPADO", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Estampado", crear_pdf_informe_productivo("Estampado", lab, df_t, df_p, m_sel, a_sel, hs_rt), "Prod_Estampado.pdf")
    if st.button("Productivo SOLDADURA", use_container_width=True): 
        with st.spinner("Generando..."):
            st.download_button("📥 Bajar PDF Soldadura", crear_pdf_informe_productivo("Soldadura", lab, df_t, df_p, m_sel, a_sel, hs_rt), "Prod_Soldadura.pdf")
