import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =========================================================================
# 1. DISEÑO DE INTERFAZ "ELITE" (CSS DINÁMICO)
# =========================================================================
st.set_page_config(page_title="Tablero Industrial Ombú", layout="wide", initial_sidebar_state="collapsed")

def inject_custom_css(main_color):
    st.markdown(f"""
    <style>
        header, [data-testid="stHeader"], footer {{display: none !important;}}
        .main {{ background-color: #0f172a; }}
        .block-container {{padding-top: 1rem !important; background-color: #0f172a;}} 

        /* Header Ombú */
        .header-ombu {{
            background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
            padding: 20px; border-radius: 15px; color: white;
            text-align: center; margin-bottom: 25px;
            box-shadow: 0px 10px 20px rgba(0,0,0,0.3);
        }}

        /* Tarjeta Maestra OEE */
        .card-oee {{
            background-color: #1e293b; border-radius: 20px; padding: 25px; 
            margin-bottom: 20px; border: 1px solid #334155;
            box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
        }}

        /* Pilares Sincronizados */
        .pillar-item {{ 
            background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 5px solid {main_color} !important;
        }}
        .p-label {{ font-size: 12px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }}
        .p-val {{ font-size: 32px; font-weight: 900; color: white; margin: 0; }}
        .p-formula {{ font-size: 10px; color: {main_color}; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; font-style: italic; }}

        /* KPIs de Soporte */
        .kpi-footer {{
            background: #1e293b; padding: 20px; border-radius: 15px;
            text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-top: 4px solid {main_color};
        }}
        .kpi-footer h2 {{ margin: 5px 0; font-size: 32px; color: white; }}
        .kpi-footer p {{ margin: 0; font-size: 11px; color: #94a3b8; font-weight: bold; text-transform: uppercase; }}
    </style>
    """, unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<div style='text-align:center; margin-top: 100px;'><h1 style='color:white;'>🏭 SISTEMA OMBÚ</h1><p style='color:#94a3b8;'>Integración Total v3.0</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u, p = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR AL SISTEMA", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS (MÁXIMA RESISTENCIA + ORDEN CRONOLÓGICO)
# =========================================================================
def safe_val(series):
    try:
        res = series.sum()
        return float(res.iloc[0]) if hasattr(res, 'iloc') else float(res)
    except: return 0.0

@st.cache_data(ttl=300)
def load_all_data():
    u_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    u_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    d_ef, d_im = pd.read_excel(u_ef), pd.read_excel(u_im)
    
    # Normalización Inteligente Eficiencias
    d_ef.columns = [str(c).strip().upper() for c in d_ef.columns]
    mapping = {'FECHA':'Fecha', 'PLANTA':'Planta', 'STD':'HH_STD', 'DISP':'HH_Disp', 'PROD':'HH_Prod', 'COSTO':'Costo'}
    new_cols = {}
    for c in d_ef.columns:
        for k, v in mapping.items():
            if k in c: new_cols[c] = v
    d_ef = d_ef.rename(columns=new_cols).loc[:, ~d_ef.columns.duplicated()]

    # Asegurar que Fecha sea datetime para ordenar bien
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef = d_ef.sort_values('Fecha') # <--- ORDEN CRONOLÓGICO
    d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')

    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)

    # Normalización Improductivas
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    m_im = {}
    for c in d_im.columns:
        if 'HH' in c and 'IMP' in c: m_im[c] = 'HH_Imp'
        elif 'TIPO' in c or 'MOTIVO' in c: m_im[c] = 'Motivo'
    d_im = d_im.rename(columns=m_im)
    if 'HH_Imp' in d_im.columns: d_im['HH_Imp'] = pd.to_numeric(d_im['HH_Imp'], errors='coerce').fillna(0)

    return d_ef, d_im

try:
    df_ef, df_im = load_all_data()
except Exception as e:
    st.error(f"Error de archivos: {e}"); st.stop()

# =========================================================================
# 4. FILTROS MAESTROS (CRONOLÓGICOS)
# =========================================================================
st.markdown("<div class='header-ombu'><h1>REPORTE INTEGRADO: OEE & GESTIÓN DE PLANTA</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("FILTROS MAESTROS")
    # Extraemos meses únicos manteniendo el orden de aparición (que ya es cronológico)
    lista_meses = df_ef['Mes_Str'].unique().tolist()
    sel_mes = st.multiselect("📅 Seleccionar Mes", lista_meses)
    sel_planta = st.multiselect("🏭 Seleccionar Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes_Str'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# Filtro espejo para Improductivas
dff_im = df_im.copy()
# (Aquí iría lógica de cruce por planta si el archivo im lo permitiera)

# =========================================================================
# 5. CÁLCULOS TÉCNICOS (SIMULADOS PARA MAÑANA)
# =========================================================================
h_std = safe_val(dff['HH_STD'])
h_disp = safe_val(dff['HH_Disp'])
h_prod = safe_val(dff['HH_Prod'])
costo_tot = safe_val(dff['Costo'])
h_imp = safe_val(dff_im['HH_Imp'])

# 1. DISPONIBILIDAD = (HH Disponibles - HH Improductivas) / HH Disponibles
# Nota: h_imp viene de dff_im. Se simula para el total seleccionado.
disponibilidad = ((h_disp - h_imp) / h_disp * 100) if h_disp > 0 else 0.0

# 2. RENDIMIENTO = Eficiencia Productiva (HH Standard / HH Productivas)
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0

# 3. CALIDAD = Piezas OK / Total (Simulado)
calidad_sim = 98.0

# OEE = D x R x C
oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Maestra Sincronizada
if oee_final >= 85: color_main = "#22c55e" 
elif oee_final >= 65: color_main = "#eab308" 
else: color_main = "#ef4444" 

# Inyectamos el CSS con el color detectado
inject_custom_css(color_main)

# =========================================================================
# 6. MÓDULO OEE PRO: EL VELOCÍMETRO
# =========================================================================
st.markdown("<div class='card-oee'>", unsafe_allow_html=True)

fig = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "INDICADOR OEE GLOBAL", 'font': {'color': 'white', 'size': 14}},
    number = {'suffix': "%", 'font': {'color': color_main, 'size': 75}},
    gauge = {
        'axis': {'range': [None, 100], 'tickcolor': "white"},
        'bar': {'color': "white"},
        'bgcolor': "rgba(0,0,0,0)",
        'steps': [
            {'range': [0, 65], 'color': '#ef4444'},
            {'range': [65, 85], 'color': '#eab308'},
            {'range': [85, 100], 'color': '#22c55e'}
        ],
        'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': 85}
    }
))
fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350, margin=dict(l=30, r=30, t=50, b=0))
st.plotly_chart(fig, use_container_width=True)

# Pilares con Fórmulas Solicitadas
st.markdown(f"""
    <div class='pillar-grid' style='display: flex; gap:15px;'>
        <div class='pillar-item'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <h2 style='color: white; margin:0;'>{min(disponibilidad, 100.0):.1f}%</h2>
            <p class='p-formula'><b>(HH Disp - HH Imp) / HH Disp</b></p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>🚀 Rendimiento</p>
            <h2 style='color: white; margin:0;'>{min(rendimiento, 100.0):.1f}%</h2>
            <p class='p-formula'><b>HH Standard / HH Productivas</b></p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>💎 Calidad</p>
            <h2 style='color: white; margin:0;'>{calidad_sim:.1f}%</h2>
            <p class='p-formula'><b>Piezas OK / Total Producido</b> <br><span style='color:#ef4444'>(DATO SIMULADO)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. INTEGRACIÓN DE MÉTRICAS DEL MAIN
# =========================================================================
st.markdown("<h4 style='color:white; margin-bottom:15px;'>📊 MÉTRICAS DE GESTIÓN INTEGRAL</h4>", unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
with c1: 
    ef_real = (h_std/h_disp*100 if h_disp>0 else 0)
    st.markdown(f"<div class='kpi-footer'><p>Eficiencia Real</p><h2 style='color:{color_main};'>{ef_real:.1f}%</h2></div>", unsafe_allow_html=True)
with c2: 
    st.markdown(f"<div class='kpi-footer'><p>HH Standard Totales</p><h2>{h_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3: 
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#ef4444;'><p>Costo Ineficiencia</p><h2 style='color:#ef4444;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)
with c4: 
    gap_h = h_disp - h_prod
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#3b82f6;'><p>Gap HH (Sin Cargar)</p><h2 style='color:#3b82f6;'>{int(gap_h)}h</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS Y TENDENCIAS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns([2, 1])

with g_col1:
    st.markdown("<h4 style='color:white;'>📈 Evolución Cronológica de Eficiencia</h4>", unsafe_allow_html=True)
    if not dff.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 4.5), facecolor='#0f172a')
        # Agrupamos respetando el orden cronológico
        ag = dff.groupby('Mes_Str', sort=False).agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax1.plot(ag.index, ag['Ef'], marker='o', color=color_main, linewidth=4, markersize=10, label='% Eficiencia Real')
        ax1.fill_between(ag.index, ag['Ef'], color=color_main, alpha=0.1)
        ax1.axhline(85, color='#22c55e', linestyle='--', alpha=0.6, label="Meta (85%)")
        ax1.set_ylim(0, 110); ax1.tick_params(colors='white')
        ax1.grid(axis='y', linestyle='--', alpha=0.1)
        ax1.legend(facecolor='#1e293b', labelcolor='white')
        st.pyplot(fig1)

with g_col2:
    st.markdown("<h4 style='color:white;'>⚠️ Top Causas Improductivas</h4>", unsafe_allow_html=True)
    if 'HH_Imp' in dff_im.columns:
        res_imp = dff_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis()
        ax2.set_xlabel("Horas Totales", color='#94a3b8')
        ax2.tick_params(colors='white')
        st.pyplot(fig2)

st.markdown("---")
st.caption("Ombu Industrial Intelligence © 2026 | Desarrollado para Control de Gestión de Planta")
