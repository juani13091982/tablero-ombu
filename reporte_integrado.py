import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import matplotlib.image as mpimg
import textwrap
import re

# =========================================================================
# 1. MOTOR DE TEMATIZACIÓN DINÁMICA (COLORES SINCRONIZADOS)
# =========================================================================
st.set_page_config(page_title="Tablero CGP Pro - Ombú S.A.", layout="wide", initial_sidebar_state="collapsed")

def inject_theme(color_status):
    st.markdown(f"""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"], footer {{display: none !important;}}
        .main {{ background-color: #0f172a; }}
        .block-container {{padding-top: 1rem !important; background-color: #0f172a;}} 
        
        /* Sticky Header */
        div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {{
            position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
            background-color: rgba(15, 23, 42, 0.98) !important; z-index: 99999 !important;
            padding: 5px 10px 15px 10px !important; border-bottom: 3px solid {color_status} !important;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
        }}

        /* Tarjeta OEE Master */
        .card-dark {{
            background-color: #1e293b; border-radius: 15px; padding: 20px; 
            margin-bottom: 15px; border: 1px solid #334155;
            box-shadow: 0px 10px 20px rgba(0,0,0,0.2);
        }}

        /* Pilares OEE */
        .pillar-box {{ 
            background: #0f172a; border-radius: 12px; padding: 15px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 5px solid {color_status} !important;
        }}
        .p-label {{ font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }}
        .p-val {{ font-size: 30px; font-weight: 900; color: white; margin: 5px 0; }}
        .p-formula {{ font-size: 10px; color: {color_status}; margin-top: 8px; border-top: 1px solid #334155; padding-top: 5px; font-style: italic; }}

        /* KPIs de Soporte (Tu Main) */
        .kpi-footer {{
            background: #1e293b; padding: 18px; border-radius: 12px;
            text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-top: 4px solid {color_status};
        }}
        .kpi-footer h2 {{ margin: 5px 0; font-size: 32px; color: white; font-weight: 900; }}
        .kpi-footer p {{ margin: 0; font-size: 11px; color: #94a3b8; font-weight: bold; text-transform: uppercase; }}

        /* Tablas */
        .stDataFrame {{ background: #1e293b; border-radius: 10px; border: 1px solid #334155; }}
    </style>
    """, unsafe_allow_html=True)

plt.rcParams.update({'font.size': 12, 'font.weight': 'bold'})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<br><br><div style='text-align:center;'><h1 style='color:white;'>🏭 SISTEMA OMBÚ</h1><p style='color:#94a3b8;'>Acceso Control de Gestión v4.0</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u, p = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR AL TABLERO", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. CARGA DE DATOS Y LÓGICA DE NEGOCIO (THE TITAN)
# =========================================================================
def safe_match(s_list, val):
    if pd.isna(val): return False
    v_norm = re.sub(r'[^A-Z0-9]', '', str(val).upper())
    for s in s_list:
        s_norm = re.sub(r'[^A-Z0-9]', '', str(s).upper())
        if s_norm == v_norm and s_norm != "": return True
    return False

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO', 'MECANICO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'LOGISTICA']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'ERROR']): return "🔎 Ajustar Calidad"
    return "⚡ Investigar Causa"

@st.cache_data(ttl=300)
def load_data_integrada():
    u_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    u_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    d_ef = pd.read_excel(u_ef).loc[:, ~pd.read_excel(u_ef).columns.duplicated()]
    d_im = pd.read_excel(u_im).loc[:, ~pd.read_excel(u_im).columns.duplicated()]
    
    # --- NORMALIZACIÓN ---
    d_ef.columns = [str(c).strip() for c in d_ef.columns]
    mapping = {
        'Fecha': 'Fecha', 'Planta': 'Planta', 'Línea': 'Linea', 'Puesto_Trabajo': 'Puesto',
        'HH_STD_TOTAL': 'HH_STD', 'HH_Disponibles': 'HH_Disp', 'HH_Productivas_C/GAP': 'HH_Prod',
        'Costo_Improd._$': 'Costo', 'Es_Ultimo_Puesto': 'Es_Ultimo'
    }
    d_ef = d_ef.rename(columns=mapping)
    
    # ORDEN CRONOLÓGICO REAL
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce', dayfirst=True)
    d_ef = d_ef.sort_values('Fecha')
    d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')

    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)

    # Improductivas
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    c_hh_im = next((c for c in d_im.columns if 'HH' in c and 'IMP' in c), None)
    c_mot_im = next((c for c in d_im.columns if 'TIPO' in c or 'MOTIVO' in c), None)
    if c_hh_im: d_im.rename(columns={c_hh_im: 'HH_Imp'}, inplace=True)
    if c_mot_im: d_im.rename(columns={c_mot_im: 'Motivo'}, inplace=True)
    d_im['HH_Imp'] = pd.to_numeric(d_im.get('HH_Imp', 0), errors='coerce').fillna(0).abs()
    
    # Operarios
    c_nom = next((c for c in d_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in d_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: d_im['OPERARIO'] = d_im[c_nom].astype(str) + ' ' + d_im[c_ape].astype(str)
    else: d_im['OPERARIO'] = "S/D"

    return d_ef, d_im

try:
    df_ef, df_im = load_data_integrada()
except Exception as e:
    st.error(f"Error cargando Drive: {e}"); st.stop()

# =========================================================================
# 4. HEADER Y FILTROS MAESTROS
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h_col1, h_col2 = st.columns([4, 1])
    h_col1.markdown("<h2 style='color:white; margin:0;'>REPORTING INDUSTRIAL CGP - INTEGRACIÓN TOTAL</h2>", unsafe_allow_html=True)
    if h_col2.button("🚪 Salir", use_container_width=True): 
        st.session_state['auth'] = False; st.rerun()

with st.sidebar:
    st.header("⚙️ FILTROS MAESTROS")
    lista_meses = df_ef['Mes_Str'].unique().tolist()
    sel_mes = st.multiselect("📅 Mes (Cronológico)", lista_meses)
    
    sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))
    sel_linea = st.multiselect("⚙️ Línea", sorted(df_ef['Linea'].dropna().unique()))

# Aplicación cruzada
df_f = df_ef.copy()
df_im_f = df_im.copy()

if sel_mes: df_f = df_f[df_f['Mes_Str'].isin(sel_mes)]
if sel_planta: 
    df_f = df_f[df_f['Planta'].isin(sel_planta)]
    c_pl = next((c for c in df_im_f.columns if 'PLANTA' in c), None)
    if c_pl: df_im_f = df_im_f[df_im_f[c_pl].astype(str).isin(sel_planta)]
if sel_linea: df_f = df_f[df_f['Linea'].isin(sel_linea)]

# =========================================================================
# 5. CÁLCULOS TÉCNICOS OEE (MAQUETA PARA MAÑANA)
# =========================================================================
def get_sum(series): return float(series.sum())

t_std = get_sum(df_f['HH_STD'])
t_disp = get_sum(df_f['HH_Disp'])
t_prod = get_sum(df_f['HH_Prod'])
t_costo = get_sum(df_f['Costo'])
t_imp = get_sum(df_im_f['HH_Imp'])

# Fórmulas Ombú
disponibilidad = ((t_disp - t_imp) / t_disp * 100) if t_disp > 0 else 0.0
rendimiento = (t_std / t_prod * 100) if t_prod > 0 else 0.0
calidad_sim = 98.0 # Simulado hasta recibir columnas de Calidad

oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Sincronizado
color_status = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 65 else "#ef4444"
inject_theme(color_status)

# =========================================================================
# 6. VELOCÍMETRO OEE PRO
# =========================================================================
st.markdown("<div class='card-dark'>", unsafe_allow_html=True)
fig_oee = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "INDICADOR OEE GLOBAL", 'font': {'color': 'white', 'size': 16}},
    number = {'suffix': "%", 'font': {'color': color_status, 'size': 80}},
    gauge = {
        'axis': {'range': [None, 100], 'tickcolor': "white"},
        'bar': {'color': "white"}, 'bgcolor': "rgba(0,0,0,0)",
        'steps': [{'range': [0, 65], 'color': '#ef4444'},{'range': [65, 85], 'color': '#eab308'},{'range': [85, 100], 'color': '#22c55e'}],
        'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': 85}
    }
))
fig_oee.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350, margin=dict(l=30, r=30, t=50, b=0))
st.plotly_chart(fig_oee, use_container_width=True)

# Pilares desglosados con fórmulas escritas
st.markdown(f"""
    <div style='display: flex; gap: 15px;'>
        <div class='pillar-box'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <p class='p-val'>{min(disponibilidad, 100.0):.1f}%</p>
            <p class='p-formula'>(HH Disp - HH Imp) / HH Disp</p>
        </div>
        <div class='pillar-box'>
            <p class='p-label'>🚀 Rendimiento</p>
            <p class='p-val'>{min(rendimiento, 100.0):.1f}%</p>
            <p class='p-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-box'>
            <p class='p-label'>💎 Calidad</p>
            <p class='p-val'>{calidad_sim:.1f}%</p>
            <p class='p-formula'>Piezas OK / Total Producido <br><span style='color:#ef4444'>(SIMULADO)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. INTEGRACIÓN MÉTRICAS MAIN (SOPORTE)
# =========================================================================
st.markdown("<h4 style='color:white; margin-bottom:15px; font-size:18px;'>📊 GESTIÓN DE PLANTA Y COSTOS</h4>", unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)

with c1:
    ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0.0
    st.markdown(f"<div class='kpi-footer'><p>Eficiencia Real</p><h2 style='color:{color_status};'>{ef_real:.1f}%</h2></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='kpi-footer'><p>HH Standard Totales</p><h2>{t_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#ef4444;'><p>Costo Ineficiencia</p><h2 style='color:#ef4444;'>${t_costo:,.0f}</h2></div>", unsafe_allow_html=True)
with c4:
    gap_h = t_disp - t_prod
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#3b82f6;'><p>Gap HH (Sin Cargar)</p><h2 style='color:#3b82f6;'>{int(gap_h)}h</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS Y PARETO
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns([2, 1])

with g_col1:
    st.markdown(f"<h4 style='color:white; border-left: 4px solid {color_status}; padding-left:10px;'>📈 Evolución de Eficiencia</h4>", unsafe_allow_html=True)
    if not df_f.empty:
        fig_trend, ax_trend = plt.subplots(figsize=(10, 4.5), facecolor='#0f172a')
        ag_trend = df_f.groupby('Mes_Str', sort=False).agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag_trend['Ef'] = (ag_trend['HH_STD'] / ag_trend['HH_Disp'] * 100)
        ax_trend.plot(ag_trend.index, ag_trend['Ef'], marker='o', color=color_status, linewidth=4, markersize=10, label='% Efic. Real')
        ax_trend.fill_between(ag_trend.index, ag_trend['Ef'], color=color_status, alpha=0.1)
        ax_trend.axhline(85, color='#22c55e', linestyle='--', alpha=0.6, label="Meta (85%)")
        ax_trend.set_ylim(0, 110); ax_trend.tick_params(colors='white'); ax_trend.grid(axis='y', alpha=0.1)
        ax_trend.legend(facecolor='#1e293b', labelcolor='white')
        st.pyplot(fig_trend)

with g_col2:
    st.markdown("<h4 style='color:white; border-left: 4px solid #ef4444; padding-left:10px;'>⚠️ Top 5 Motivos de Parada</h4>", unsafe_allow_html=True)
    if not df_im_f.empty:
        res_pareto = df_im_f.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig_p, ax_p = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_pareto.plot(kind='barh', color='#ef4444', ax=ax_p)
        ax_p.invert_yaxis(); ax_p.tick_params(colors='white')
        st.pyplot(fig_p)

# =========================================================================
# 9. TABLA DE DETALLES Y MOTOR DE ACCIONES
# =========================================================================
st.markdown("---")
st.header("📋 ANÁLISIS DETALLADO Y ACCIONES")

if not df_im_f.empty:
    df_im_f['Acción Sugerida'] = df_im_f['Detalle'].apply(generar_accion_sugerida)
    cols_display = ['OPERARIO', 'Motivo', 'Detalle', 'HH_Imp', 'Acción Sugerida']
    df_ver = df_im_f[cols_display].sort_values(by='HH_Imp', ascending=False)
    st.dataframe(df_ver, use_container_width=True, hide_index=True)
else:
    st.info("Sin registros de improductividad para los filtros seleccionados.")

st.caption("Ombu Industrial Intelligence © 2026 | Desarrollado para Control de Gestión")
