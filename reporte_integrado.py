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
# 1. CONFIGURACIÓN Y ESTILO DINÁMICO (INTEGRACIÓN VISUAL)
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")

def inyectar_estilo_sincronizado(color_kpi):
    st.markdown(f"""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="manage-app-button"], 
        #MainMenu, footer, .stAppDeployButton, .viewerBadge_container {{display: none !important; visibility: hidden !important;}}
        .block-container {{padding-top: 1rem !important; padding-bottom: 1.5rem !important; background-color: #0f172a;}}
        
        div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {{
            position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
            background-color: rgba(14, 17, 23, 0.98) !important; z-index: 99999 !important;
            padding: 5px 10px 15px 10px !important; border-bottom: 2px solid {color_kpi} !important;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
        }}
        
        /* Tarjeta OEE Master */
        .card-oee {{
            background-color: #1e293b; border-radius: 20px; padding: 25px; 
            margin-bottom: 20px; border: 1px solid #334155;
            box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
        }}

        /* Pilares Sincronizados */
        .pillar-item {{ 
            background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 5px solid {color_kpi} !important;
        }}
        .p-label {{ font-size: 12px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }}
        .p-val {{ font-size: 32px; font-weight: 900; color: white; margin: 0; }}
        .p-formula {{ font-size: 10px; color: {color_kpi}; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; font-style: italic; }}

        /* KPIs Originales del Main */
        .kpi-integrated {{
            background: #1e293b; border-radius: 15px; padding: 20px; text-align: center;
            border-top: 4px solid {color_kpi}; box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        }}
        .kpi-integrated h2 {{ color: white !important; font-size: 38px !important; margin: 5px 0; }}
        .kpi-integrated h4 {{ color: #94a3b8 !important; font-size: 14px !important; text-transform: uppercase; }}
    </style>
    """, unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True); c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='text-align:center;'><h2 style='color:white;'>GESTIÓN INDUSTRIAL OMBÚ</h2></div>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in, p_in = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE Y CARGA (LÓGICA DEL MAIN)
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
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'LOGISTICA']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'ERROR']): return "🔎 Ajustar Calidad"
    return "⚡ Investigar Causa"

@st.cache_data(ttl=300)
def cargar_datos_completos():
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    d_ef = pd.read_excel(url_ef)
    d_im = pd.read_excel(url_im)
    
    d_ef.columns = d_ef.columns.str.strip()
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    
    # Forzado Numérico
    for col in ['HH_STD_TOTAL', 'HH_Disponibles', 'HH_Productivas_C/GAP', 'Costo_Improd._$']:
        if col in d_ef.columns: d_ef[col] = pd.to_numeric(d_ef[col], errors='coerce').fillna(0)
    
    # Orden Cronológico
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce', dayfirst=True)
    d_ef = d_ef.sort_values('Fecha')
    d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
    # Improductivas
    if 'HH_IMPRODUCTIVAS' not in d_im.columns:
        c_imp = next((c for c in d_im.columns if 'HH' in c and 'IMP' in c), d_im.columns[0])
        d_im.rename(columns={c_imp: 'HH_IMPRODUCTIVAS'}, inplace=True)
    d_im['HH_IMPRODUCTIVAS'] = pd.to_numeric(d_im['HH_IMPRODUCTIVAS'], errors='coerce').fillna(0).abs()
    
    # Operario (Nombre + Apellido)
    c_nom = next((c for c in d_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in d_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: d_im['OPERARIO'] = d_im[c_nom].astype(str) + ' ' + d_im[c_ape].astype(str)
    else: d_im['OPERARIO'] = "S/D"
    
    return d_ef, d_im

try:
    df_ef, df_im = cargar_datos_completos()
except Exception as e:
    st.error(f"Error cargando Drive: {e}"); st.stop()

# =========================================================================
# 4. FILTROS MAESTROS (SIDEBAR CRONOLÓGICO)
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h1, h2 = st.columns([4, 1])
    h1.markdown("<h3 style='margin:0; color:white;'>REPORTING INDUSTRIAL CGP - INTEGRACIÓN TOTAL</h3>", unsafe_allow_html=True)
    if h2.button("🚪 Salir", use_container_width=True): 
        st.session_state['autenticado'] = False; st.rerun()

with st.sidebar:
    st.header("🎛️ FILTROS MAESTROS")
    lista_meses = df_ef['Mes_Str'].unique().tolist()
    sel_mes = st.multiselect("📅 Mes", lista_meses, placeholder="Seleccionar Mes")
    sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))
    sel_linea = st.multiselect("⚙️ Línea", sorted(df_ef['Línea'].dropna().unique()))

# Aplicar filtros
df_ef_f = df_ef.copy()
df_im_f = df_im.copy()

if sel_mes: df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(sel_mes)]
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
    c_pl = next((c for c in df_im_f.columns if 'PLANTA' in c), None)
    if c_pl: df_im_f = df_im_f[df_im_f[c_pl].astype(str).isin(sel_planta)]
if sel_linea: df_ef_f = df_ef_f[df_ef_f['Línea'].isin(sel_linea)]

# =========================================================================
# 5. CÁLCULOS OEE (FORMULACIÓN OMBÚ SIMULADA)
# =========================================================================
def get_sum(series): return float(series.sum())

t_std = get_sum(df_ef_f['HH_STD_TOTAL'])
t_disp = get_sum(df_ef_f['HH_Disponibles'])
t_prod = get_sum(df_ef_f['HH_Productivas_C/GAP'])
t_costo = get_sum(df_ef_f['Costo_Improd._$'])
t_imp = get_sum(df_im_f['HH_IMPRODUCTIVAS'])

# 1. DISPONIBILIDAD = (HH Disponibles - HH Improductivas) / HH Disponibles
disponibilidad = ((t_disp - t_imp) / t_disp * 100) if t_disp > 0 else 0.0

# 2. RENDIMIENTO = Eficiencia Productiva (HH Standard / HH Productivas)
rendimiento = (t_std / t_prod * 100) if t_prod > 0 else 0.0

# 3. CALIDAD = 98.0% (Simulado)
calidad_sim = 98.0

oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Sincronizado
color_main = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 65 else "#ef4444"
inyectar_estilo_sincronizado(color_main)

# =========================================================================
# 6. MÓDULO OEE PRO: VELOCÍMETRO + PILARES
# =========================================================================
st.markdown("<div class='card-oee'>", unsafe_allow_html=True)
fig_oee = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "INDICADOR OEE GLOBAL", 'font': {'color': 'white', 'size': 16}},
    number = {'suffix': "%", 'font': {'color': color_main, 'size': 75}},
    gauge = {
        'axis': {'range': [None, 100], 'tickcolor': "white"},
        'bar': {'color': "white"}, 'bgcolor': "rgba(0,0,0,0)",
        'steps': [{'range': [0, 65], 'color': '#ef4444'},{'range': [65, 85], 'color': '#eab308'},{'range': [85, 100], 'color': '#22c55e'}],
        'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': 85}
    }
))
fig_oee.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350, margin=dict(l=30, r=30, t=50, b=0))
st.plotly_chart(fig_oee, use_container_width=True)

st.markdown(f"""
    <div style='display: flex; gap: 15px;'>
        <div class='pillar-item'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <h2 style='color: white; margin:0;'>{min(disponibilidad, 100.0):.1f}%</h2>
            <p class='p-formula'>(HH Disp - HH Imp) / HH Disp</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>🚀 Rendimiento</p>
            <h2 style='color: white; margin:0;'>{min(rendimiento, 100.0):.1f}%</h2>
            <p class='p-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>💎 Calidad</p>
            <h2 style='color: white; margin:0;'>{calidad_sim:.1f}%</h2>
            <p class='p-formula'>Piezas OK / Total <span style='color:#ef4444'>(SIM)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. KPIs DE SOPORTE (INTEGRACIÓN DEL MAIN)
# =========================================================================
st.markdown("<h4 style='color:white; margin-bottom:15px;'>📊 MÉTRICAS DE GESTIÓN Y COSTOS</h4>", unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
with c1: 
    ef_r = (t_std / t_disp * 100) if t_disp > 0 else 0
    st.markdown(f"<div class='kpi-integrated'><h4>Eficiencia Real</h4><h2 style='color:{color_main};'>{ef_r:.1f}%</h2></div>", unsafe_allow_html=True)
with c2: 
    st.markdown(f"<div class='kpi-integrated'><h4>HH Standard</h4><h2>{t_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3: 
    st.markdown(f"<div class='kpi-integrated' style='border-top-color:#ef4444;'><h4>Costo Ineficiencia</h4><h2 style='color:#ef4444;'>${t_costo:,.0f}</h2></div>", unsafe_allow_html=True)
with c4: 
    st.markdown(f"<div class='kpi-integrated' style='border-top-color:#3b82f6;'><h4>Gap de Horas</h4><h2 style='color:#3b82f6;'>{int(t_disp - t_prod)}h</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS DE LAS MÉTRICAS (M1 A M6)
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns(2)

with g_col1:
    st.markdown(f"<h4 style='color:white; border-left: 5px solid {color_main}; padding-left:10px;'>1. TENDENCIA EFICIENCIA REAL</h4>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 6), facecolor='#0f172a')
        ag = df_ef_f.groupby('Mes_Str', sort=False).agg({'HH_STD_TOTAL':'sum', 'HH_Disponibles':'sum'})
        ag['Ef'] = (ag['HH_STD_TOTAL'] / ag['HH_Disponibles'] * 100)
        ax1.plot(ag.index, ag['Ef'], marker='o', color=color_main, linewidth=4, markersize=10)
        ax1.fill_between(ag.index, ag['Ef'], color=color_main, alpha=0.1)
        ax1.set_ylim(0, 110); ax1.tick_params(colors='white'); ax1.grid(axis='y', alpha=0.1)
        st.pyplot(fig1)

with g_col2:
    st.markdown("<h4 style='color:white; border-left: 5px solid #ef4444; padding-left:10px;'>2. PARETO CAUSAS IMPRODUCTIVAS</h4>", unsafe_allow_html=True)
    if not df_im_f.empty:
        res = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(10, 6), facecolor='#0f172a')
        res.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis(); ax2.tick_params(colors='white')
        st.pyplot(fig2)

# =========================================================================
# 9. TABLA DE DETALLES (MESA DE TRABAJO)
# =========================================================================
st.markdown("---")
st.header("📋 DETALLE OPERATIVO Y ACCIONES")
if not df_im_f.empty:
    df_im_f['Acción Sugerida'] = df_im_f['DETALLE'].apply(generar_accion_sugerida)
    # Mostramos lo más importante del Main
    cols = ['OPERARIO', 'TIPO_PARADA', 'DETALLE', 'HH_IMPRODUCTIVAS', 'Acción Sugerida']
    df_mesa = df_im_f[cols].sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
    st.dataframe(df_mesa, use_container_width=True, hide_index=True)
else:
    st.info("No hay improductividad registrada para este periodo.")

st.caption("Ombu Industrial Intelligence © 2026 | Datos en Tiempo Real")
