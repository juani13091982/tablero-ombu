import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import re

# =========================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y DISEÑO CSS "PREMIUM"
# =========================================================================
st.set_page_config(page_title="Tablero de Control Industrial - Ombú", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ocultar basura de Streamlit */
    header, [data-testid="stHeader"], footer {display: none !important;}
    .block-container {padding-top: 0.5rem !important; background-color: #f4f7f9;}
    
    /* Header Estilo Multinacional */
    .main-header {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 20px;
        border-radius: 0px 0px 20px 20px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0px 4px 15px rgba(0,0,0,0.2);
    }

    /* Cartel de OEE "HERMOSO" */
    .oee-container {
        background: white;
        border-radius: 20px;
        padding: 30px;
        text-align: center;
        border: 1px solid #e0e6ed;
        box-shadow: 0px 10px 25px rgba(0,0,0,0.05);
        margin-bottom: 25px;
    }
    .oee-value {
        font-size: 80px;
        font-weight: 900;
        margin: 0;
        line-height: 1;
    }
    .oee-label {
        font-size: 18px;
        color: #64748b;
        letter-spacing: 2px;
        font-weight: bold;
    }
    
    /* Grid de Pilares OEE */
    .pilar-grid {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 15px;
        margin-top: 20px;
    }
    .pilar-card {
        padding: 15px;
        border-radius: 15px;
        background: #f8fafc;
        border: 1px solid #f1f5f9;
    }
    .pilar-val { font-size: 24px; font-weight: bold; color: #1e293b; }
    .pilar-txt { font-size: 12px; color: #94a3b8; text-transform: uppercase; }

    /* Tarjetas KPI Rápidas */
    .kpi-card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        border-left: 8px solid #1E3A8A;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.03);
    }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<br><br><div style='text-align:center;'><h1>OMBÚ S.A.</h1><p>Control de Gestión Industrial</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("ACCEDER AL TABLERO", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Credenciales incorrectas")
    st.stop()

# =========================================================================
# 3. CARGA DE DATOS Y LIMPIEZA
# =========================================================================
def safe_num(series):
    try: return float(series.sum())
    except: return 0.0

@st.cache_data(ttl=300)
def load_data():
    u_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    u_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    d_ef = pd.read_excel(u_ef)
    d_im = pd.read_excel(u_im)
    
    # Normalizar Eficiencias
    d_ef.columns = [str(c).strip().upper() for c in d_ef.columns]
    m_ef = {}
    for c in d_ef.columns:
        if 'FECHA' in c: m_ef[c] = 'Fecha'
        elif 'PLANTA' in c: m_ef[c] = 'Planta'
        elif 'STD' in c: m_ef[c] = 'HH_STD'
        elif 'DISP' in c: m_ef[c] = 'HH_Disp'
        elif 'GAP' in c or 'PROD' in c: m_ef[c] = 'HH_Prod'
        elif 'COSTO' in c: m_ef[c] = 'Costo'
    d_ef = d_ef.rename(columns=m_ef)
    d_ef = d_ef.loc[:, ~d_ef.columns.duplicated()]
    
    # Normalizar Improductivas
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    m_im = {}
    for c in d_im.columns:
        if 'HH' in c and 'IMP' in c: m_im[c] = 'HH_Imp'
        elif 'TIPO' in c or 'MOTIVO' in c: m_im[c] = 'Motivo'
        elif 'DETALLE' in c: m_im[c] = 'Detalle'
    d_im = d_im.rename(columns=m_im)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)
    if 'HH_Imp' in d_im.columns: d_im['HH_Imp'] = pd.to_numeric(d_im['HH_Imp'], errors='coerce').fillna(0)
    
    return d_ef, d_im

df_ef, df_im = load_data()

# =========================================================================
# 4. FILTROS (STICKY)
# =========================================================================
st.markdown("<div class='main-header'><h1>REPORTING INDUSTRIAL CGP - OMBÚ</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR_pD_pZ48aE7rE-j9J2N7xXgXp6qZp3T6w2A&s", width=150) # Logo simulado
    st.header("CONFIGURACIÓN")
    sel_mes = st.multiselect("📅 Seleccionar Mes", sorted(df_ef['Mes'].dropna().unique()))
    sel_planta = st.multiselect("🏭 Seleccionar Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS OEE DEFINITIVOS
# =========================================================================
h_std = safe_num(dff['HH_STD'])
h_disp = safe_num(dff['HH_Disp'])
h_prod = safe_num(dff['HH_Prod'])
costo_tot = safe_num(dff['Costo'])

disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
calidad_sim = 98.0
oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100

# Limitar para que no explote visualmente
disponibilidad = min(disponibilidad, 100.0)
rendimiento = min(rendimiento, 100.0)

# =========================================================================
# 6. DISEÑO: EL CARTEL DE OEE "HERMOSO"
# =========================================================================
color_oee = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 60 else "#ef4444"

st.markdown(f"""
<div class='oee-container'>
    <p class='oee-label'>OEE: OVERALL EQUIPMENT EFFECTIVENESS</p>
    <h1 class='oee-value' style='color: {color_oee};'>{oee_final:.1f}%</h1>
    <div style='background: #e2e8f0; height: 12px; border-radius: 10px; width: 60%; margin: 20px auto; overflow: hidden;'>
        <div style='background: {color_oee}; width: {min(oee_final, 100)}%; height: 100%;'></div>
    </div>
    <div class='pilar-grid'>
        <div class='pilar-card'>
            <div class='pilar-txt'>⏱️ Disponibilidad</div>
            <div class='pilar-val'>{disponibilidad:.1f}%</div>
        </div>
        <div class='pilar-card'>
            <div class='pilar-txt'>🚀 Rendimiento</div>
            <div class='pilar-val'>{rendimiento:.1f}%</div>
        </div>
        <div class='pilar-card'>
            <div class='pilar-txt'>💎 Calidad</div>
            <div class='pilar-val'>{calidad_sim:.1f}%</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KPIs Secundarios
c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div class='kpi-card'><h6>EFICIENCIA REAL</h6><h2 style='color:#1E3A8A;'>{(h_std/h_disp*100):.1f}%</h2></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-card' style='border-left-color:#22c55e;'><h6>TOTAL PRODUCIDO</h6><h2 style='color:#16a34a;'>{h_std:,.0f} <span style='font-size:15px;'>HH STD</span></h2></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-card' style='border-left-color:#ef4444;'><h6>COSTO INEFICIENCIA</h6><h2 style='color:#dc2626;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS Y DETALLE
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g1, g2 = st.columns([2, 1])

with g1:
    st.subheader("📈 Evolución de Eficiencia Real por Mes")
    if not dff.empty:
        fig, ax = plt.subplots(figsize=(10, 4))
        # Estilo gráfico moderno
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax.plot(ag.index, ag['Ef'], marker='o', color='#1E3A8A', linewidth=4, markersize=10, label='% Eficiencia')
        ax.fill_between(ag.index, ag['Ef'], alpha=0.1, color='#1E3A8A')
        ax.axhline(85, color='green', linestyle='--', alpha=0.5, label='Meta 85%')
        ax.set_ylim(0, 110)
        ax.legend()
        st.pyplot(fig)

with g2:
    st.subheader("⚠️ Top Causas Improductivas")
    if 'HH_Imp' in df_im.columns:
        res_imp = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(6, 7.5))
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis()
        st.pyplot(fig2)

st.markdown("---")
st.subheader("📋 Detalle de Registros Improductivos")
if not df_im.empty:
    st.dataframe(df_im.sort_values('HH_Imp', ascending=False), use_container_width=True, hide_index=True)

st.success("✅ Tablero Premium cargado. ¡OEEEEE!")
