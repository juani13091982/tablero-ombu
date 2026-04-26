import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt

# =========================================================================
# 1. DISEÑO DE INTERFAZ "EXCELENCIA" (CSS CUSTOM DARK MODE)
# =========================================================================
st.set_page_config(page_title="KPI Master - Ombú S.A.", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    header, [data-testid="stHeader"], footer {display: none !important;}
    .main { background-color: #0f172a; }
    .block-container {padding-top: 1rem !important; background-color: #0f172a;} 

    /* Header Ombú Pro */
    .header-ombu {
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 25px; border-radius: 15px; color: white;
        text-align: center; margin-bottom: 25px;
        box-shadow: 0px 10px 20px rgba(0,0,0,0.3);
    }

    /* Contenedor del Velocímetro y Pilares */
    .card-dark {
        background-color: #1e293b; 
        border-radius: 20px; padding: 25px; margin-bottom: 20px;
        border: 1px solid #334155; box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
    }

    /* Pilares OEE */
    .pillar-grid { display: flex; justify-content: space-around; gap: 15px; margin-top: 15px; }
    .pillar-item { 
        background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
        border: 1px solid #334155; text-align: center;
    }
    .p-label { font-size: 13px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }
    .p-formula { font-size: 10px; color: #3b82f6; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; font-style: italic; }

    /* KPIs de Integración (Lo que venía del Main) */
    .kpi-footer {
        background: #1e293b; padding: 20px; border-radius: 15px;
        text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border-top: 4px solid #3b82f6;
    }
    .kpi-footer h2 { margin: 5px 0; font-size: 32px; font-weight: 900; }
    .kpi-footer p { margin: 0; font-size: 12px; color: #94a3b8; font-weight: bold; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<div style='text-align:center; margin-top: 100px;'><h1 style='color:white;'>🏭 SISTEMA OMBÚ</h1><p style='color:#94a3b8;'>Integración de Métricas v2.0</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("ACCEDER AL TABLERO", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS (MÁXIMA RESISTENCIA)
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
    d_ef = pd.read_excel(u_ef); d_im = pd.read_excel(u_im)
    
    # Normalización Inteligente de Columnas (Mapea por fragmentos de texto)
    d_ef.columns = [str(c).strip().upper() for c in d_ef.columns]
    mapping = {
        'FECHA':'Fecha', 'PLANTA':'Planta', 'STD':'HH_STD', 
        'DISP':'HH_Disp', 'PROD':'HH_Prod', 'GAP':'HH_Prod', 'COSTO':'Costo'
    }
    new_cols = {}
    for c in d_ef.columns:
        for k, v in mapping.items():
            if k in c: new_cols[c] = v
    d_ef = d_ef.rename(columns=new_cols)
    d_ef = d_ef.loc[:, ~d_ef.columns.duplicated()]

    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes'] = d_ef['Fecha'].dt.strftime('%b-%Y')

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
    st.error(f"Error cargando archivos de Google Drive: {e}"); st.stop()

# =========================================================================
# 4. HEADER Y FILTROS INTEGRADOS
# =========================================================================
st.markdown("<div class='header-ombu'><h1>INTEGRACIÓN TOTAL: OEE & EFICIENCIA DE PLANTA</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("PANEL DE FILTROS")
    sel_mes = st.multiselect("📅 Mes", sorted(df_ef['Mes'].dropna().unique()))
    sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS TÉCNICOS SINCRONIZADOS
# =========================================================================
h_std = safe_val(dff['HH_STD'])
h_disp = safe_val(dff['HH_Disp'])
h_prod = safe_val(dff['HH_Prod'])
costo_tot = safe_val(dff['Costo'])

# Cálculos de Pilares
disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
calidad_sim = 98.0
oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100

# Lógica de colores unificada para todo el tablero
if oee_final >= 85: color_main = "#22c55e" # Verde
elif oee_final >= 65: color_main = "#eab308" # Amarillo
else: color_main = "#ef4444" # Rojo

# =========================================================================
# 6. MÓDULO OEE PRO: VELOCÍMETRO + FÓRMULAS
# =========================================================================
st.markdown("<div class='card-dark'>", unsafe_allow_html=True)

# Velocímetro con Plotly (Sincronizado en color)
fig = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "INDICADOR OEE GLOBAL", 'font': {'color': 'white', 'size': 14}},
    number = {'suffix': "%", 'font': {'color': color_main, 'size': 75}},
    gauge = {
        'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "white"},
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

# Pilares Desglosados con Fórmulas Claritas
st.markdown(f"""
    <div class='pillar-grid'>
        <div class='pillar-item' style='border-top: 4px solid {color_main};'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <h2 style='color: white; margin:0;'>{min(disponibilidad, 100.0):.1f}%</h2>
            <p class='p-formula'><b>Fórmula:</b> HH Productivas / HH Disponibles</p>
        </div>
        <div class='pillar-item' style='border-top: 4px solid {color_main};'>
            <p class='p-label'>🚀 Rendimiento</p>
            <h2 style='color: white; margin:0;'>{min(rendimiento, 100.0):.1f}%</h2>
            <p class='p-formula'><b>Fórmula:</b> HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-item' style='border-top: 4px solid {color_main};'>
            <p class='p-label'>💎 Calidad</p>
            <h2 style='color: white; margin:0;'>{calidad_sim:.1f}%</h2>
            <p class='p-formula'><b>Fórmula:</b> Piezas OK / Total Producido <br><span style='color:#ef4444'>(DATO SIMULADO)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. INTEGRACIÓN DE KPIs DEL MAIN (SOPORTE ECONÓMICO)
# =========================================================================
st.markdown("<h4 style='color:white; margin-bottom:15px;'>📊 MÉTRICAS DE SOPORTE INTEGRADO</h4>", unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
with c1: 
    st.markdown(f"<div class='kpi-footer'><p>Eficiencia Real</p><h2 style='color:{color_main};'>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h2></div>", unsafe_allow_html=True)
with c2: 
    st.markdown(f"<div class='kpi-footer'><p>HH Standard Producidas</p><h2 style='color:white;'>{h_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3: 
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#ef4444;'><p>Costo Ineficiencia</p><h2 style='color:#ef4444;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)
with c4: 
    gap_h = h_disp - h_prod
    st.markdown(f"<div class='kpi-footer' style='border-top-color:#3b82f6;'><p>Horas No Cargadas (GAP)</p><h2 style='color:#3b82f6;'>{int(gap_h)}h</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS INTEGRADOS CON LEYENDAS Y TÍTULOS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns([2, 1])

with g_col1:
    st.markdown("<h4 style='color:white;'>📈 Tendencia Mensual de Eficiencia Real</h4>", unsafe_allow_html=True)
    if not dff.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 4.5), facecolor='#0f172a')
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax1.plot(ag.index, ag['Ef'], marker='o', color='#3b82f6', linewidth=4, markersize=10, label='% Eficiencia Real')
        ax1.fill_between(ag.index, ag['Ef'], color='#3b82f6', alpha=0.1)
        ax1.axhline(85, color='#22c55e', linestyle='--', alpha=0.6, label="Meta Clase Mundial (85%)")
        ax1.set_xlabel("Mes / Año", color='#94a3b8', fontsize=10)
        ax1.set_ylabel("Porcentaje Eficiencia (%)", color='#94a3b8', fontsize=10)
        ax1.set_ylim(0, 110); ax1.tick_params(colors='white')
        ax1.grid(axis='y', linestyle='--', alpha=0.2)
        ax1.legend(facecolor='#1e293b', labelcolor='white', loc='lower right')
        st.pyplot(fig1)

with g_col2:
    st.markdown("<h4 style='color:white;'>⚠️ Causas de Parada</h4>", unsafe_allow_html=True)
    if 'HH_Imp' in df_im.columns:
        fig2, ax2 = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_imp = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis()
        ax2.set_xlabel("Horas Totales Perdidas", color='#94a3b8', fontsize=10)
        ax2.tick_params(colors='white')
        st.pyplot(fig2)

st.markdown("---")
st.caption("Reporting Ombú Industrial Intelligence © 2026 | Datos en Tiempo Real desde Google Drive")
