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
# 0. CONFIGURACIÓN
# ==========================================
st.set_page_config(page_title="Reportes Fumiscor", layout="wide", page_icon="📊")

# ==========================================
# 1. FUNCIONES AUXILIARES Y PDF
# ==========================================
class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area; self.fecha_str = fecha_str; self.theme_color = theme_color

    def add_gradient_background(self):
        r1, g1, b1 = 240, 242, 246
        r2, g2, b2 = 215, 220, 225
        h = self.h; w = self.w
        for i in range(int(h * 2)):
            ratio = i / (h * 2)
            r = int(r1 + (r2 - r1) * ratio); g = int(g1 + (g2 - g1) * ratio); b = int(b1 + (b2 - b1) * ratio)
            self.set_fill_color(r, g, b); self.rect(0, i / 2, w, 0.5, 'F')

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
        self.set_fill_color(*bg_color); self.set_draw_color(180, 180, 180); self.rounded_rect(x, y, w, h, r, style='DF')

    def draw_kpi_panel(self, x, y, w, h, r=3, bg_color=None):
        bg = bg_color if bg_color else self.theme_color
        self.set_fill_color(200, 200, 200); self.rounded_rect(x + 1.5, y + 1.5, w, h, r, style='F')
        self.set_fill_color(*bg); self.rounded_rect(x, y, w, h, r, style='F')

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

def save_chart(fig, w=600, h=300):
    fig.update_layout(width=w, height=h, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig.write_image(tmp.name, engine="kaleido", scale=2.5); return tmp.name

# ==========================================
# 2. CARGA DE DATOS (NATIVA DESDE SQL)
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, mes, anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        
        ini_str = fecha_ini.strftime('%Y-%m-%d 00:00:00')
        fin_str = fecha_fin.strftime('%Y-%m-%d 23:59:59')
        
        q_metrics = f"""
            SELECT UPPER(f.Name) as Area, UPPER(l.Name) as Grupo, c.Name as Máquina, 
                   SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, SUM(COALESCE(p.Scrap, 0)) as Observadas, 
                   SUM(COALESCE(p.ProductiveTime, 0)) as T_Operativo, SUM(COALESCE(p.DownTime, 0)) as T_Parada, 
                   SUM(COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado, 
                   SUM(COALESCE(p.Performance, 0) * COALESCE(p.ProductiveTime, 0)) as Perf_Num, 
                   SUM(COALESCE(p.Availability, 0) * (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0))) as Disp_Num, 
                   SUM(COALESCE(p.Quality, 0) * (COALESCE(p.Good, 0) + COALESCE(p.Rework, 0) + COALESCE(p.Scrap, 0))) as Cal_Num, 
                   SUM(COALESCE(p.Oee, 0) * (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0))) as OEE_Num 
            FROM PROD_M_03 p 
            JOIN CELL c ON p.CellId = c.CellId 
            LEFT JOIN LINE l ON c.LineId = l.LineId
            LEFT JOIN FACTORY f ON l.FactoryId = f.FactoryId
            WHERE p.Year = {anio} AND p.Month = {mes} 
            GROUP BY f.Name, l.Name, c.Name
        """
        
        q_event = f"""
            SELECT UPPER(f.Name) as Area, UPPER(l.Name) as Grupo, c.Name as Máquina, e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4] 
            FROM EVENT_01 e 
            LEFT JOIN CELL c ON e.CellId = c.CellId 
            LEFT JOIN LINE l ON c.LineId = l.LineId
            LEFT JOIN FACTORY f ON l.FactoryId = f.FactoryId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId 
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId 
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId 
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId 
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        
        q_piezas = f"""
            SELECT UPPER(f.Name) as Area, UPPER(l.Name) as Grupo, c.Name as Máquina, COALESCE(pr.Code, 'S/C') as Pieza, 
                   SUM(COALESCE(p.Scrap, 0)) as Scrap, SUM(COALESCE(p.Rework, 0)) as RT 
            FROM PROD_M_01 p 
            JOIN CELL c ON p.CellId = c.CellId 
            LEFT JOIN LINE l ON c.LineId = l.LineId
            LEFT JOIN FACTORY f ON l.FactoryId = f.FactoryId
            LEFT JOIN PRODUCT pr ON p.ProductId = pr.ProductId 
            WHERE p.Year = {anio} AND p.Month = {mes} 
            GROUP BY f.Name, l.Name, c.Name, pr.Code
        """

        q_trend_oee_monthly = f"""
            SELECT p.Month, UPPER(f.Name) as Area, UPPER(l.Name) as Grupo, c.Name as Máquina, 
                   SUM(COALESCE(p.ProductiveTime, 0)) as T_Operativo, SUM(COALESCE(p.DownTime, 0)) as T_Parada, 
                   SUM(COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0)) as T_Planificado, 
                   SUM(COALESCE(p.Performance, 0) * COALESCE(p.ProductiveTime, 0)) as Perf_Num, 
                   SUM(COALESCE(p.Availability, 0) * (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0))) as Disp_Num, 
                   SUM(COALESCE(p.Quality, 0) * (COALESCE(p.Good, 0) + COALESCE(p.Rework, 0) + COALESCE(p.Scrap, 0))) as Cal_Num, 
                   SUM(COALESCE(p.Oee, 0) * (COALESCE(p.ProductiveTime, 0) + COALESCE(p.DownTime, 0))) as OEE_Num 
            FROM PROD_M_03 p 
            JOIN CELL c ON p.CellId = c.CellId 
            LEFT JOIN LINE l ON c.LineId = l.LineId
            LEFT JOIN FACTORY f ON l.FactoryId = f.FactoryId
            WHERE p.Year = {anio} AND p.Month <= {mes} 
            GROUP BY p.Month, f.Name, l.Name, c.Name
        """
        
        q_trend_piezas_monthly = f"""
            SELECT p.Month, UPPER(f.Name) as Area, UPPER(l.Name) as Grupo, c.Name as Máquina, 
                   SUM(COALESCE(p.Good, 0)) as Buenas, SUM(COALESCE(p.Rework, 0)) as Retrabajo, 
                   SUM(COALESCE(p.Scrap, 0)) as Observadas, 
                   SUM(COALESCE(p.Good, 0) + COALESCE(p.Rework, 0) + COALESCE(p.Scrap, 0)) as Totales 
            FROM PROD_M_01 p 
            JOIN CELL c ON p.CellId = c.CellId 
            LEFT JOIN LINE l ON c.LineId = l.LineId
            LEFT JOIN FACTORY f ON l.FactoryId = f.FactoryId
            WHERE p.Year = {anio} AND p.Month <= {mes} 
            GROUP BY p.Month, f.Name, l.Name, c.Name
        """

        q_m06 = f"SELECT 'GLOBAL' as Nivel, 'GLOBAL' as Grupo, Performance, Availability as Disp, Quality as Cal, Oee FROM PROD_M_06 WHERE Year = {anio} AND Month = {mes}"
        q_m05 = f"SELECT 'FABRICA' as Nivel, UPPER(f.Name) as Grupo, p.Performance, p.Availability as Disp, p.Quality as Cal, p.Oee FROM PROD_M_05 p JOIN FACTORY f ON p.FactoryId = f.FactoryId WHERE p.Year = {anio} AND p.Month = {mes}"
        q_m04 = f"SELECT 'LINEA' as Nivel, UPPER(l.Name) as Grupo, p.Performance, p.Availability as Disp, p.Quality as Cal, p.Oee FROM PROD_M_04 p JOIN LINE l ON p.LineId = l.LineId WHERE p.Year = {anio} AND p.Month = {mes}"

        df_metrics = conn.query(q_metrics).fillna(0)
        df_raw = conn.query(q_event)
        df_piezas = conn.query(q_piezas).fillna(0)
        df_trend_oee = conn.query(q_trend_oee_monthly).fillna(0)
        df_trend_piezas = conn.query(q_trend_piezas_monthly).fillna(0)
        df_oficial = pd.concat([conn.query(q_m06).fillna(0), conn.query(q_m05).fillna(0), conn.query(q_m04).fillna(0)], ignore_index=True)

        # Limpieza estándar para evitar nulos en Area y Grupo
        for df in [df_metrics, df_raw, df_piezas, df_trend_oee, df_trend_piezas]:
            if not df.empty and 'Area' in df.columns:
                df['Area'] = df['Area'].fillna('SIN AREA').astype(str).str.strip().str.upper()
                df['Grupo'] = df['Grupo'].fillna('SIN GRUPO').astype(str).str.strip().str.upper()

        cols_metrics = ['Buenas', 'Retrabajo', 'Observadas', 'T_Operativo', 'T_Parada', 'T_Planificado', 'Perf_Num', 'Disp_Num', 'Cal_Num', 'OEE_Num']
        for c in cols_metrics:
            if c in df_metrics.columns: df_metrics[c] = pd.to_numeric(df_metrics[c], errors='coerce').fillna(0)

        for col in ['Month', 'T_Operativo', 'T_Parada', 'T_Planificado', 'Perf_Num', 'Disp_Num', 'Cal_Num', 'OEE_Num']:
            if col in df_trend_oee.columns: df_trend_oee[col] = pd.to_numeric(df_trend_oee[col], errors='coerce').fillna(0)

        for col in ['Month', 'Buenas', 'Retrabajo', 'Observadas', 'Totales']:
            if col in df_trend_piezas.columns: df_trend_piezas[col] = pd.to_numeric(df_trend_piezas[col], errors='coerce').fillna(0)

        if not df_trend_oee.empty and not df_trend_piezas.empty:
            df_trend = pd.merge(df_trend_piezas, df_trend_oee, on=['Month', 'Area', 'Grupo', 'Máquina'], how='outer').fillna(0)
        else:
            df_trend = df_trend_piezas if not df_trend_piezas.empty else df_trend_oee

        if df_raw.empty: 
            df_raw = pd.DataFrame(columns=['Area', 'Grupo', 'Máquina', 'Tiempo (Min)', 'Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4', 'Estado_Global', 'Categoria_Macro', 'Detalle_Final'])
        else:
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)
            for col in ['Nivel Evento 1', 'Nivel Evento 2', 'Nivel Evento 3', 'Nivel Evento 4']:
                if col in df_raw.columns: df_raw[col] = df_raw[col].fillna('').astype(str)
                else: df_raw[col] = ''
                    
            mask_proyecto = (df_raw['Nivel Evento 1'].str.upper().str.contains('PROYECTO') | df_raw['Nivel Evento 2'].str.upper().str.contains('PROYECTO') | df_raw['Nivel Evento 3'].str.upper().str.contains('PROYECTO') | df_raw['Nivel Evento 4'].str.upper().str.contains('PROYECTO'))
            df_raw = df_raw[~mask_proyecto].copy()

            def cat_estado(row):
                t1 = row['Nivel Evento 1'].strip().upper()
                t2 = row['Nivel Evento 2'].strip().upper()
                if 'PRODUC' in t1 or 'PRODUC' in t2: return 'Producción'
                if 'PARADA' in t1 or 'PARADA' in t2: return 'Parada Programada'
                return 'Falla/Gestión'
            
            def cat_macro(row):
                n1 = row['Nivel Evento 1'].strip().upper()
                n2 = row['Nivel Evento 2'].strip().title()
                if 'GESTION' in n1 or 'GESTIÓN' in n1 or 'GESTION' in n2.upper() or 'GESTIÓN' in n2.upper(): return 'Gestión'
                elif 'FALLA' in n1: return n2 if n2 else 'Fallas Generales'
                return n1.title() if n1 else 'Sin Área'
            
            def get_det(row):
                n1 = row['Nivel Evento 1'].strip().upper()
                n2 = row['Nivel Evento 2'].strip()
                n3 = row['Nivel Evento 3'].strip()
                n4 = row['Nivel Evento 4'].strip()
                
                if n4 and n4.lower() not in ['nan', 'none', 'null', '']: return n4
                if n3 and n3.lower() not in ['nan', 'none', 'null', '']: return n3
                if n2 and n2.lower() not in ['nan', 'none', 'null', '']: return n2
                if n1 and n1.lower() not in ['nan', 'none', 'null', '']: return n1
                return "Sin detalle"
                
            df_raw['Estado_Global'] = df_raw.apply(cat_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(cat_macro, axis=1)
            df_raw['Detalle_Final'] = df_raw.apply(get_det, axis=1)

        return df_metrics, df_raw, df_trend, df_piezas, df_oficial
    except Exception as e: 
        st.error(f"Error SQL: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. MOTOR: GESTIÓN A LA VISTA (DISPONIBILIDAD)
# ==========================================
def crear_pdf_gestion_a_la_vista(area, label_reporte, df_metrics_pdf, df_pdf_raw, df_trend, df_oficial, mes_seleccionado):
    if area.upper() == "ESTAMPADO": theme_color = (15, 76, 129)
    elif area.upper() == "SOLDADURA": theme_color = (211, 84, 0)
    else: theme_color = (40, 40, 40)

    pdf = ReportePDF(f"GESTIÓN A LA VISTA - {area}", label_reporte, theme_color)
    
    df_m = df_metrics_pdf.copy(); df_t = df_trend.copy(); df_r = df_pdf_raw.copy()
            
    df_m_all = df_m.copy(); df_t_all = df_t.copy(); df_r_all = df_r.copy()

    # Filtrar por Área (Fábrica) a menos que sea GLOBAL
    if area.upper() != "GLOBAL":
        df_m = df_m[df_m['Area'] == area.upper()]
        df_t = df_t[df_t['Area'] == area.upper()]
        df_r = df_r[df_r['Area'] == area.upper()]
    
    # Obtener grupos disponibles dinámicamente según la data de la DB
    grupos_area = sorted([g for g in df_m['Grupo'].unique() if g and g != 'SIN GRUPO'])
    paginas = ['GENERAL'] if area.upper() == "GLOBAL" else ['GENERAL'] + grupos_area

    TARGETS = {"OEE": 0.75, "PERFORMANCE": 0.90, "DISPONIBILIDAD": 0.88, "CALIDAD": 0.95}

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        
        if target == 'GENERAL':
            if area.upper() == 'SOLDADURA':
                # Excluir celdas nuevas del target general si Fumiscor lo sigue requiriendo
                df_m_target = df_m[~df_m['Grupo'].str.contains('NUEVA', na=False)]
                df_t_target = df_t[~df_t['Grupo'].str.contains('NUEVA', na=False)]
                df_r_target = df_r[~df_r['Grupo'].str.contains('NUEVA', na=False)]
            elif area.upper() == 'GLOBAL':
                df_m_target = df_m_all; df_t_target = df_t_all; df_r_target = df_r_all
            else:
                df_m_target = df_m; df_t_target = df_t; df_r_target = df_r
        else:
            df_m_target = df_m[df_m['Grupo'] == target]; df_t_target = df_t[df_t['Grupo'] == target]; df_r_target = df_r[df_r['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(40, 6, "PERIODO", 1, 0, 'C', fill=True)
        pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}" if area.upper() != "GLOBAL" else "PLANTA GLOBAL FUMISCOR - RESUMEN GENERAL", 1, 0, 'C', fill=True)
        pdf.cell(40, 6, "INFORME", 1, 1, 'C', fill=True)
        
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(40, 6, label_reporte, 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "DISPONIBILIDAD", 1, 1, 'C', fill=True)

        if not df_m_target.empty:
            if 'Totales' not in df_m_target.columns:
                df_m_target['Totales'] = df_m_target['Buenas'] + df_m_target['Retrabajo'] + df_m_target['Observadas']
            valid_m = df_m_target[(df_m_target['T_Planificado'] > 0) & (df_m_target['T_Operativo'] > 0) & (df_m_target['Totales'] > 0)]
        else:
            valid_m = pd.DataFrame()

        v_oee, v_perf, v_disp, v_cal = 0, 0, 0, 0
        encontrado_oficial = False
        
        if not df_oficial.empty:
            if target == 'GENERAL':
                if area.upper() == 'GLOBAL': row = df_oficial[df_oficial['Nivel'] == 'GLOBAL']
                else: row = df_oficial[(df_oficial['Nivel'] == 'FABRICA') & (df_oficial['Grupo'] == area.upper())]
            else:
                row = df_oficial[(df_oficial['Nivel'] == 'LINEA') & (df_oficial['Grupo'] == target)]
                
            if not row.empty:
                v_oee = row['Oee'].values[0]
                v_perf = row['Performance'].values[0]
                v_disp = row['Disp'].values[0]
                v_cal = row['Cal'].values[0]
                encontrado_oficial = True

        if not encontrado_oficial or (v_oee == 0 and v_perf == 0 and v_disp == 0): 
            t_plan = valid_m['T_Planificado'].sum() if not valid_m.empty else 0
            t_op = valid_m['T_Operativo'].sum() if not valid_m.empty else 0
            t_piezas = valid_m['Totales'].sum() if not valid_m.empty else 0
            
            v_oee = (valid_m['OEE_Num'].sum() / t_plan) if t_plan > 0 else 0
            v_perf = (valid_m['Perf_Num'].sum() / t_op) if t_op > 0 else 0
            v_disp = (valid_m['Disp_Num'].sum() / t_plan) if t_plan > 0 else 0
            v_cal = (valid_m['Cal_Num'].sum() / t_piezas) if t_piezas > 0 else 0
        
        if v_oee > 1.5 or v_perf > 1.5 or v_disp > 1.5:
            v_oee /= 100.0; v_perf /= 100.0; v_disp /= 100.0; v_cal /= 100.0
            
        kpis = {
            "OEE": {"val": v_oee, "obj": TARGETS["OEE"]},
            "PERFORMANCE": {"val": v_perf, "obj": TARGETS["PERFORMANCE"]},
            "DISPONIBILIDAD": {"val": v_disp, "obj": TARGETS["DISPONIBILIDAD"]},
            "CALIDAD": {"val": v_cal, "obj": TARGETS["CALIDAD"]}
        }
        
        for i, (lbl, data) in enumerate(kpis.items()):
            v = data["val"]
            obj = data["obj"]
            
            if v < obj: bg_col, txt_col = (231, 76, 60), 255
            else: bg_col, txt_col = (46, 204, 113), 255

            x = 10 + (i * 68.5)
            pdf.draw_kpi_panel(x, y_kpi:=25, 65, 20, bg_color=bg_col)
            pdf.set_xy(x, y_kpi + 2); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(txt_col); pdf.cell(65, 6, lbl, 0, 1, 'L')
            pdf.set_xy(x, y_kpi + 8); pdf.set_font("Arial", 'B', 20); pdf.cell(65, 10, f"{v*100:.1f}%", 0, 0, 'C')
        pdf.set_text_color(0)

        def add_trend_bar(df_in, col, title, x_pos, y_pos, target_val, off_val=None, draw_large=False):
            if df_in.empty: return
            
            cols_req = ['OEE_Num', 'T_Planificado', 'Perf_Num', 'T_Operativo', 'Disp_Num', 'Cal_Num', 'Totales']
            for c in cols_req:
                if c in df_in.columns: df_in[c] = pd.to_numeric(df_in[c], errors='coerce').fillna(0)
            
            if 'Totales' in df_in.columns: df_valid = df_in[(df_in['T_Planificado'] > 0) & (df_in['T_Operativo'] > 0) & (df_in['Totales'] > 0)]
            else: df_valid = df_in[(df_in['T_Planificado'] > 0) & (df_in['T_Operativo'] > 0)]
                
            if df_valid.empty: return
            
            df_g = df_valid.groupby('Month')[cols_req].sum().reset_index()
            if 'Month' in df_g.columns: df_g['Month'] = df_g['Month'].astype(int)
            
            if col == 'OEE': df_g['Val'] = df_g.apply(lambda r: r['OEE_Num'] / r['T_Planificado'] if r.get('T_Planificado', 0) > 0 else 0, axis=1)
            elif col == 'PERFORMANCE': df_g['Val'] = df_g.apply(lambda r: r['Perf_Num'] / r['T_Operativo'] if r.get('T_Operativo', 0) > 0 else 0, axis=1)
            elif col == 'DISPONIBILIDAD': df_g['Val'] = df_g.apply(lambda r: r['Disp_Num'] / r['T_Planificado'] if r.get('T_Planificado', 0) > 0 else 0, axis=1)
            elif col == 'CALIDAD': df_g['Val'] = df_g.apply(lambda r: r['Cal_Num'] / r['Totales'] if r.get('Totales', 0) > 0 else 0, axis=1)
            else: return
            
            if df_g['Val'].max() > 1.5: df_g['Val'] /= 100.0

            if off_val is not None:
                df_g.loc[df_g['Month'] == mes_seleccionado, 'Val'] = off_val

            ytd_v = 0
            if col == 'OEE': ytd_v = df_valid['OEE_Num'].sum() / df_valid['T_Planificado'].sum() if df_valid['T_Planificado'].sum() > 0 else 0
            elif col == 'PERFORMANCE': ytd_v = df_valid['Perf_Num'].sum() / df_valid['T_Operativo'].sum() if df_valid['T_Operativo'].sum() > 0 else 0
            elif col == 'DISPONIBILIDAD': ytd_v = df_valid['Disp_Num'].sum() / df_valid['T_Planificado'].sum() if df_valid['T_Planificado'].sum() > 0 else 0
            elif col == 'CALIDAD': ytd_v = df_valid['Cal_Num'].sum() / df_valid['Totales'].sum() if df_valid['Totales'].sum() > 0 else 0
            if ytd_v > 1.5: ytd_v /= 100.0

            def get_c(v): return '#2ECC71' if v >= target_val else '#E74C3C'
            
            df_g['Mes_Str'] = df_g['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})
            df_g['Color'] = df_g['Val'].apply(get_c)
            
            ytd_row = pd.DataFrame([{'Month': 99, 'Mes_Str': 'Acum.', 'Val': ytd_v, 'Color': get_c(ytd_v)}])
            df_g = pd.concat([df_g, ytd_row], ignore_index=True)

            max_y = df_g['Val'].max() if not df_g.empty else 1
            upper_limit = max(1.1, max_y * 1.3, target_val * 1.2)

            fig = go.Figure(data=[go.Bar(x=df_g['Mes_Str'], y=df_g['Val'], marker=dict(color=df_g['Color'], line=dict(color='rgba(0,0,0,0.8)', width=2)), text=df_g['Val'], texttemplate='<b>%{text:.1%}</b>', textposition='outside', opacity=0.85)])
            
            fig.add_hline(y=target_val, line_dash="dash", line_color="#2ECC71", line_width=2, annotation_text=f"<b>Obj: {target_val*100:.0f}%</b>", annotation_font_color='black', annotation_position="top left")
            
            if len(df_g) > 1:
                fig.add_vline(x=len(df_g) - 1.5, line_width=2, line_dash="dot", line_color="rgba(0,0,0,0.6)")
            
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(t=35, b=20, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, upper_limit], visible=False), xaxis_title="")
            fig.update_traces(textfont=dict(color='black', size=11, family="Arial"), cliponaxis=False)
            
            w_img, h_img = (600, 300) if draw_large else (600, 220)
            w_pdf = 132 if not draw_large else 134
            img = save_chart(fig, w_img, h_img); pdf.image(img, x=x_pos+2, y=y_pos+2, w=w_pdf); os.remove(img)

        if area.upper() == "GLOBAL":
            pdf.draw_panel(10, 48, 136, 75); pdf.draw_panel(149, 48, 138, 75)
            add_trend_bar(df_t_target, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48, TARGETS["OEE"], v_oee, draw_large=True)
            add_trend_bar(df_t_target, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48, TARGETS["PERFORMANCE"], v_perf, draw_large=True) 
            
            pdf.draw_panel(10, 126, 136, 75); pdf.draw_panel(149, 126, 138, 75)
            add_trend_bar(df_t_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 126, TARGETS["DISPONIBILIDAD"], v_disp, draw_large=True)
            add_trend_bar(df_t_target, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 126, TARGETS["CALIDAD"], v_cal, draw_large=True)
        else:
            pdf.draw_panel(10, 48, 136, 52); pdf.draw_panel(149, 48, 138, 52)
            add_trend_bar(df_t_target, 'OEE', 'OEE (%) - EVOLUCIÓN MENSUAL', 10, 48, TARGETS["OEE"], v_oee)
            add_trend_bar(df_t_target, 'PERFORMANCE', 'PERFORMANCE (%) - EVOLUCIÓN MENSUAL', 150, 48, TARGETS["PERFORMANCE"], v_perf) 
            
            pdf.draw_panel(10, 102, 136, 52); pdf.draw_panel(149, 102, 138, 52)
            add_trend_bar(df_t_target, 'DISPONIBILIDAD', 'DISPONIBILIDAD (%) - EVOLUCIÓN MENSUAL', 10, 102, TARGETS["DISPONIBILIDAD"], v_disp)
            add_trend_bar(df_t_target, 'CALIDAD', 'CALIDAD (%) - EVOLUCIÓN MENSUAL', 150, 102, TARGETS["CALIDAD"], v_cal)
            
            pdf.draw_panel(10, 156, 136, 45); pdf.draw_panel(149, 156, 138, 45)
            pdf.set_xy(10, 156); pdf.set_font("Times", 'B', 11); pdf.set_text_color(0); pdf.cell(136, 6, "TOP 5 FALLOS", border=0, ln=True, align='C')
            
            df_f = df_r_target[df_r_target['Estado_Global'] == 'Falla/Gestión'] if not df_r_target.empty else pd.DataFrame()
            
            if not df_f.empty and df_f['Tiempo (Min)'].sum() > 0:
                excluir = ['BAÑO', 'BANO', 'REFRIGERIO', 'DESCANSO']
                mask_puras = ~df_f['Detalle_Final'].str.upper().apply(lambda x: any(excl in x for excl in excluir))
                df_f_puras = df_f[mask_puras]
                
                top5 = df_f_puras.groupby('Detalle_Final')['Tiempo (Min)'].sum().nlargest(5).reset_index()
                
                pdf.set_xy(10, 162); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(*theme_color); pdf.set_text_color(255)
                pdf.cell(76, 5, "FALLO", border=1, fill=True); pdf.cell(30, 5, "MINUTOS", border=1, align='C', fill=True); pdf.cell(30, 5, "% TOTAL", border=1, align='C', ln=True, fill=True)
                pdf.set_font("Arial", '', 8); pdf.set_text_color(0); pdf.set_fill_color(255, 255, 255)
                
                t_total = df_f['Tiempo (Min)'].sum()
                for _, r in top5.iterrows():
                    pdf.set_x(10); pdf.cell(76, 6, clean_text(str(r['Detalle_Final']))[:45], border=1, fill=True)
                    pdf.cell(30, 6, f"{r['Tiempo (Min)']:.0f}", border=1, align='C', fill=True)
                    pdf.cell(30, 6, f"{(r['Tiempo (Min)']/t_total)*100:.1f}%", border=1, align='C', ln=True, fill=True)
                
                df_macro = df_f.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
                df_macro['%'] = df_macro['Tiempo (Min)'] / t_total
                df_macro['Y'] = "Pérdidas"
                df_macro['Leyenda'] = df_macro.apply(lambda r: f"{r['Categoria_Macro']} ({r['Tiempo (Min)']/60:.1f} hs | {r['%']:.1%})", axis=1)
                
                fig_stack = px.bar(df_macro, x='%', y='Y', color='Leyenda', orientation='h', color_discrete_sequence=px.colors.qualitative.Safe)
                fig_stack.update_traces(texttemplate='<b>%{x:.1%}</b>', textposition='inside', marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.9, textfont=dict(color='black', size=11))
                fig_stack.update_layout(barmode='stack', title=dict(text="<b>PROPORCIÓN DE PÉRDIDAS ÁREAS MACRO (100%)</b>", font=dict(family="Times", size=13, color="black")), xaxis=dict(visible=False, range=[0, 1]), yaxis=dict(visible=False), margin=dict(t=30, b=5, l=10, r=10), legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5, title="", font=dict(size=10)))
                img_stack = save_chart(fig_stack, 600, 180); pdf.image(img_stack, 151, 158, 134); os.remove(img_stack)
            
    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 4. MOTOR: INFORME PRODUCTIVO (CALIDAD)
# ==========================================
def crear_pdf_informe_productivo(area, label_reporte, df_trend, df_piezas, mes_sel, anio_sel, hs_rt):
    theme_color = (15, 76, 129) if area.upper() == "ESTAMPADO" else (211, 84, 0)
    theme_hex = '#%02x%02x%02x' % theme_color
    scrap_c = '#002147' if area.upper() == "ESTAMPADO" else '#722F37' 
    rt_c = theme_hex
    
    pdf = ReportePDF(f"INFORME PRODUCTIVO - {area}", label_reporte, theme_color)
    
    df_t = df_trend.copy(); df_p = df_piezas.copy()
    
    df_t = df_t[df_t['Area'] == area.upper()]
    df_p = df_p[df_p['Area'] == area.upper()]

    grupos_area = sorted([g for g in df_t['Grupo'].unique() if g and g != 'SIN GRUPO'])
    paginas = ['GENERAL'] + grupos_area

    if area.upper() == "ESTAMPADO":
        target_scrap = 0.50; target_rt = 2.00
    else:
        target_scrap = 0.30; target_rt = 2.00

    for target in paginas:
        pdf.add_page(orientation='L'); pdf.set_auto_page_break(False); pdf.add_gradient_background()
        
        if target == 'GENERAL': df_t_target = df_t; df_p_target = df_p
        else: df_t_target = df_t[df_t['Grupo'] == target]; df_p_target = df_p[df_p['Grupo'] == target]
        
        pdf.set_y(10); pdf.set_fill_color(*theme_color); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 10)
        pdf.cell(20, 6, "MES", 1, 0, 'C', fill=True); pdf.cell(20, 6, "AÑO", 1, 0, 'C', fill=True); pdf.cell(197, 6, f"PLANTA {area.upper()} - {target}", 1, 0, 'C', fill=True); pdf.cell(40, 6, "AREA", 1, 1, 'C', fill=True)
        pdf.set_fill_color(255, 255, 255); pdf.set_font("Arial", '', 10); pdf.set_text_color(0)
        pdf.cell(20, 6, str(mes_sel), 1, 0, 'C', fill=True); pdf.cell(20, 6, str(anio_sel), 1, 0, 'C', fill=True); pdf.set_font("Arial", 'B', 10); pdf.cell(197, 6, "EMPRESA: FUMISCOR", 1, 0, 'C', fill=True); pdf.set_font("Arial", '', 10); pdf.cell(40, 6, "PRODUCTIVO", 1, 1, 'C', fill=True)

        if df_t_target.empty: continue
        
        for col in ['Buenas', 'Observadas', 'Retrabajo', 'Totales']:
            if col in df_t_target.columns: df_t_target[col] = pd.to_numeric(df_t_target[col], errors='coerce').fillna(0)

        df_ev = df_t_target.groupby('Month')[['Buenas', 'Observadas', 'Retrabajo', 'Totales']].sum().reset_index()
        if 'Month' in df_ev.columns: df_ev['Month'] = df_ev['Month'].astype(int)
        
        df_ev['Totales_Div'] = df_ev['Totales'].apply(lambda x: x if x > 0 else 1)
        df_ev['% Scrap'] = ((df_ev['Observadas'] / df_ev['Totales_Div']) * 100).round(2)
        df_ev['% RT'] = ((df_ev['Retrabajo'] / df_ev['Totales_Div']) * 100).round(2)
        df_ev['Mes_Str'] = df_ev['Month'].map({1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'})

        df_ev['Color_Scrap'] = df_ev['% Scrap'].apply(lambda x: '#E74C3C' if x > target_scrap else '#2ECC71')
        df_ev['Color_RT'] = df_ev['% RT'].apply(lambda x: '#E74C3C' if x > target_rt else '#2ECC71')

        f1 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['Totales'], marker_color=theme_hex, text=df_ev['Totales'], texttemplate='<b>%{text:.3s}</b>')])
        f2 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['% Scrap'], marker_color=df_ev['Color_Scrap'], text=df_ev['% Scrap'], texttemplate='<b>%{text:.2f}%</b>')])
        f3 = go.Figure(data=[go.Bar(x=df_ev['Mes_Str'], y=df_ev['% RT'], marker_color=df_ev['Color_RT'], text=df_ev['% RT'], texttemplate='<b>%{text:.2f}%</b>')])
        
        titles = ["PIEZAS PRODUCIDAS MES A MES", "% DE SCRAP MES A MES", "% DE RT MES A MES"]
        for i, f in enumerate([f1, f2, f3]): 
            max_y = df_ev['Totales'].max() if i==0 else (df_ev['% Scrap'].max() if i==1 else df_ev['% RT'].max())
            if i == 0: upper_limit = max_y * 1.3 if max_y > 0 else 1
            else: 
                current_target = target_scrap if i == 1 else target_rt
                upper_limit = max(0.2, max_y * 1.3, current_target * 1.5)
                f.add_hline(y=current_target, line_dash="dash", line_width=2, line_color="#E74C3C", annotation_text=f"<b>Obj: {current_target}%</b>", annotation_font_color='black')
                
            f.update_yaxes(range=[0, upper_limit])
            f.update_layout(title=dict(text=f"<b>{titles[i]}</b>", font=dict(family="Times", size=13, color="black")), margin=dict(l=10, r=10, t=30, b=20), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis=dict(visible=False))
            f.update_traces(textposition="outside", cliponaxis=False, textfont=dict(color='black', size=11, family="Arial"), marker_line_color='rgba(0,0,0,0.8)', marker_line_width=2, opacity=0.85)

        h_box = 60; pdf.draw_panel(10, 22, 135, h_box); pdf.draw_panel(10, 85, 135, h_box); pdf.draw_panel(10, 148, 135, h_box)
        i1 = save_chart(f1, w=550, h=260); pdf.image(i1, x=11, y=23, w=133, h=h_box-2); os.remove(i1)
        i2 = save_chart(f2, w=550, h=260); pdf.image(i2, x=11, y=86, w=133, h=h_box-2); os.remove(i2)
        i3 = save_chart(f3, w=550, h=260); pdf.image(i3, x=11, y=149, w=133, h=h_box-2); os.remove(i3)

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

            i4 = save_chart(f4, w=550, h=330); pdf.image(i4, x=151, y=23, w=133, h=h_br-2); os.remove(i4)
            i5 = save_chart(f5, w=550, h=330); pdf.image(i5, x=151, y=109.5, w=133, h=h_br-2); os.remove(i5)
            
        if target == 'GENERAL' and area.upper() == 'ESTAMPADO':
            pdf.draw_panel(150, 196, 135, 12, 2, (240,240,240)); pdf.set_xy(150, 196); pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0); pdf.cell(67.5, 12, "HS DE RT", 0, 0, 'C')
            pdf.draw_panel(217.5, 196, 67.5, 12, 2, (255,255,255)); pdf.set_xy(217.5, 196); pdf.cell(67.5, 12, f"{hs_rt:.1f}", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5. INTERFAZ STREAMLIT
# ==========================================
st.title("📄 Reportes Fumiscor")
st.divider()

st.write("### 1. Seleccione el Período (Mensual)")
col1, col2 = st.columns(2)
today = pd.to_datetime("today").date()
with col1: 
    m_sel = st.selectbox("Mes", range(1, 13), index=today.month-1)
with col2: 
    a_sel = st.selectbox("Año", [2024, 2025, 2026], index=2)

ini = pd.to_datetime(f"{a_sel}-{m_sel}-01")
fin = pd.to_datetime(f"{a_sel}-{m_sel}-{calendar.monthrange(a_sel, m_sel)[1]}")
lab = f"{m_sel}/{a_sel}"

with st.spinner("Conectando con la base de datos de Fumiscor..."):
    df_m, df_r, df_t, df_p, df_oficial = fetch_data_from_db(ini, fin, m_sel, a_sel)

st.write("### 2. Datos Manuales (Informe Productivo)")
hs_rt = st.number_input("Horas de RT (Solo válido para Estampado General):", min_value=0.0, max_value=1000.0, value=0.0, step=1.0)

st.divider()

st.write("### 2.5. Corrección de Indicadores Oficiales (Wiidem)")
st.info("Estos son los valores que figurarán en el PDF. Si Wiidem los tiene calculados, aparecen aquí. Si no, **el sistema los pre-calculó automáticamente para evitar que queden en 0**. Puede editarlos libremente si es necesario.")

def calcular_kpis_base(df_m_raw):
    if df_m_raw.empty: return pd.DataFrame()
    df = df_m_raw.copy()
    df['Totales'] = df['Buenas'] + df['Retrabajo'] + df['Observadas']
    
    resultados = []
    def calc_r(name, nivel, data):
        if data.empty: return {'Nivel': nivel, 'Grupo': name, 'Performance': 0.0, 'Disp': 0.0, 'Cal': 0.0, 'Oee': 0.0}
        t_plan = data['T_Planificado'].sum()
        t_op = data['T_Operativo'].sum()
        t_pz = data['Totales'].sum()
        
        return {
            'Nivel': nivel, 'Grupo': name,
            'Performance': (data['Perf_Num'].sum() / t_op * 100) if t_op > 0 else 0,
            'Disp': (data['Disp_Num'].sum() / t_plan * 100) if t_plan > 0 else 0,
            'Cal': (data['Cal_Num'].sum() / t_pz * 100) if t_pz > 0 else 0,
            'Oee': (data['OEE_Num'].sum() / t_plan * 100) if t_plan > 0 else 0
        }

    # Nivel Global
    resultados.append(calc_r('GLOBAL', 'GLOBAL', df))
    
    # Nivel Fábrica (ej. ESTAMPADO, SOLDADURA)
    areas_unicas = [a for a in df['Area'].unique() if pd.notna(a) and a != 'SIN AREA']
    for a in areas_unicas:
        resultados.append(calc_r(a, 'FABRICA', df[df['Area'] == a]))
        
    # Nivel Línea (ej. BALANCINES, PRENSAS)
    grupos_unicos = [g for g in df['Grupo'].unique() if pd.notna(g) and g != 'SIN GRUPO']
    for g in grupos_unicos:
        resultados.append(calc_r(g, 'LINEA', df[df['Grupo'] == g]))
        
    return pd.DataFrame(resultados)

df_base_editor = calcular_kpis_base(df_m)

if not df_base_editor.empty and not df_oficial.empty:
    df_base_editor.set_index(['Nivel', 'Grupo'], inplace=True)
    df_of_idx = df_oficial.set_index(['Nivel', 'Grupo'])
    for col in ['Performance', 'Disp', 'Cal', 'Oee']:
        if col in df_of_idx.columns:
            valid_vals = df_of_idx[df_of_idx[col] > 0][col]
            df_base_editor.update(valid_vals)
    df_base_editor.reset_index(inplace=True)
elif df_base_editor.empty:
    df_base_editor = pd.DataFrame(columns=['Nivel', 'Grupo', 'Performance', 'Disp', 'Cal', 'Oee'])

df_oficial_editado = st.data_editor(
    df_base_editor,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Nivel": st.column_config.TextColumn("Nivel", disabled=True),
        "Grupo": st.column_config.TextColumn("Grupo", disabled=True),
        "Performance": st.column_config.NumberColumn("Performance", format="%.4f", step=0.01),
        "Disp": st.column_config.NumberColumn("Disponibilidad", format="%.4f", step=0.01),
        "Cal": st.column_config.NumberColumn("Calidad", format="%.4f", step=0.01),
        "Oee": st.column_config.NumberColumn("OEE", format="%.4f", step=0.01),
    }
)

st.divider()
st.write("### 3. Preparar y Descargar Reportes")
c_d, c_p, c_g = st.columns(3)

with c_d:
    st.markdown("#### ⚙️ Disponibilidad (OEE)")
    if not df_m.empty:
        if st.button("⚙️ Preparar PDF Estampado", use_container_width=True):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_oee_est_fumis'] = crear_pdf_gestion_a_la_vista("Estampado", lab, df_m, df_r, df_t, df_oficial_editado, m_sel)
        if 'pdf_oee_est_fumis' in st.session_state:
            st.download_button("📥 Bajar PDF Estampado", data=st.session_state['pdf_oee_est_fumis'], file_name="FUMISCOR_Gestion_Vista_ESTAMPADO.pdf", mime="application/pdf", use_container_width=True)
            
        st.write("---")
        
        if st.button("⚙️ Preparar PDF Soldadura", use_container_width=True):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_oee_sol_fumis'] = crear_pdf_gestion_a_la_vista("Soldadura", lab, df_m, df_r, df_t, df_oficial_editado, m_sel)
        if 'pdf_oee_sol_fumis' in st.session_state:
            st.download_button("📥 Bajar PDF Soldadura", data=st.session_state['pdf_oee_sol_fumis'], file_name="FUMISCOR_Gestion_Vista_SOLDADURA.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.error("No hay datos.")

with c_p:
    st.markdown("#### 🏭 Informe Productivo (Calidad)")
    if not df_t.empty:
        if st.button("🏭 Preparar Prod. Estampado", use_container_width=True):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_prod_est_fumis'] = crear_pdf_informe_productivo("Estampado", lab, df_t, df_p, m_sel, a_sel, hs_rt)
        if 'pdf_prod_est_fumis' in st.session_state:
            st.download_button("📥 Bajar Prod. Estampado", data=st.session_state['pdf_prod_est_fumis'], file_name="FUMISCOR_Productivo_Vista_ESTAMPADO.pdf", mime="application/pdf", use_container_width=True)
        
        st.write("---")
        
        if st.button("🏭 Preparar Prod. Soldadura", use_container_width=True):
            with st.spinner("Generando documento..."):
                st.session_state['pdf_prod_sol_fumis'] = crear_pdf_informe_productivo("Soldadura", lab, df_t, df_p, m_sel, a_sel, hs_rt)
        if 'pdf_prod_sol_fumis' in st.session_state:
            st.download_button("📥 Bajar Prod. Soldadura", data=st.session_state['pdf_prod_sol_fumis'], file_name="FUMISCOR_Productivo_Vista_SOLDADURA.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.error("No hay datos.")

with c_g:
    st.markdown("#### 🌎 Reporte Maestro")
    if not df_m.empty:
        if st.button("🌎 Preparar PDF Global", use_container_width=True):
            with st.spinner("Generando documento maestro..."):
                st.session_state['pdf_oee_glob_fumis'] = crear_pdf_gestion_a_la_vista("GLOBAL", lab, df_m, df_r, df_t, df_oficial_editado, m_sel)
        if 'pdf_oee_glob_fumis' in st.session_state:
            st.download_button("📥 Bajar PDF Global", data=st.session_state['pdf_oee_glob_fumis'], file_name="FUMISCOR_Vista_GENERAL.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.error("No hay datos.")
