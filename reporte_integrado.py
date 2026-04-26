import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =========================================================================
# 1. DISEÑO DE INTERFAZ "PLATINUM" (CSS CUSTOM)
# =========================================================================
st.set_page_config(page_title="Dashboard OEE - Ombú S.A.", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    header, [data-testid="stHeader"], footer {display: none !important;}
    .block-container {padding-top: 1rem !important; background-color: #f8fafc;}

    /* Encabezado Principal */
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 20px; border-radius: 15px; color: white;
        text-align: center; margin-bottom: 25px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.2);
    }

    /* Contenedor Maestro OEE */
    .oee-hero {
        background: white; border-radius: 20px; padding: 40px;
        text-align: center; border: 1px solid #e2e8f0;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.05); margin-bottom: 30px;
    }
    .oee-main-val { font-size: 90px; font-weight: 900; line-height: 1; margin: 10px 0; }
    .oee-formula-global { font-size: 16px; color: #64748b; font-style: italic; margin-bottom: 20px; }
    
    /* Barra de Progreso */
    .bar-bg { background: #f1f5f9; height: 20px; border-radius: 10px; width: 80%; margin: 0 auto 30px auto; overflow: hidden; border: 1px solid #e2e8f0; }
    .bar-fill { height: 100%; border-radius: 10px; transition: width 1.5s ease-in-out; }

    /* Pilares y Fórmulas */
    .pillar-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; }
    .pillar-card { background: #ffffff; padding: 20px; border-radius: 15px; border: 1px solid #f1f5f9; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); transition: 0.3s; }
    .pillar-card:hover { transform: translateY(-5px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
    .pillar-name { font-size: 14px; color: #1e3a8a; font-weight: 800; text-transform: uppercase; margin-bottom: 5px; }
    .pillar-val { font-size: 32px; font-weight: 800; color: #1e293b; margin-bottom: 10px; }
    .pillar-formula { font-size: 11px; color: #94a3b8; background: #f8fafc; padding: 8px; border-radius: 8px; border: 1px dashed #cbd5e1; }

    /* Estilo de KPIs secundarios */
    .kpi-sub {
        background: white; padding: 20px; border-radius: 12px;
        border-bottom: 4px solid #1e3a8a; text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
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
        st.markdown("<div style='text-align:center; margin-top: 100px;'><h1>🏭 SISTEMA OMBÚ</h1><p>Control de Gestión de Planta</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS RESISTENTE (BLINDADO)
# =========================================================================
def safe_calc(series):
    try:
        res = series.sum()
        return float(res.iloc[0]) if hasattr(res, 'iloc') else float(res)
    except: return 0.0

@st.cache_data(ttl=300)
def load_data():
    u_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    u_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    d_ef = pd.read_excel(u_ef)
    d_im = pd.read_excel(u_im)
    
    # Normalizador Inteligente
    def fix_cols(df):
        df.columns = [str(c).strip().upper() for c in df.columns]
        mapping = {}
        for c in df.columns:
            if 'FECHA' in c: mapping[c] = 'Fecha'
            elif 'PLANTA' in c: mapping[c] = 'Planta'
            elif 'STD' in c: mapping[c] = 'HH_STD'
            elif 'DISP' in c: mapping[c] = 'HH_Disp'
            elif 'PROD' in c or 'GAP' in c: mapping[c] = 'HH_Prod'
            elif 'COSTO' in c: mapping[c] = 'Costo'
        df = df.rename(columns=mapping)
        return df.loc[:, ~df.columns.duplicated()]

    d_ef = fix_cols(d_ef)
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
    # Improductivas
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    m_im = {}
    for c in d_im.columns:
        if 'HH' in c and 'IMP' in c: m_im[c] = 'HH_Imp'
        elif 'TIPO' in c or 'MOTIVO' in c: m_im[c] = 'Motivo'
    d_im = d_im.rename(columns=m_im)
    
    return d_ef, d_im

try:
    df_ef, df_im = load_data()
except Exception as e:
    st.error(f"Error: {e}"); st.stop()

# =========================================================================
# 4. FILTROS
# =========================================================================
st.markdown("<div class='main-header'><h1>DASHBOARD INTEGRADO: OEE & EFICIENCIA</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://www.maquinariapropia.com.ar/uploads/marcas/logo-ombu.png", width=150)
    st.header("FILTROS")
    sel_mes = st.multiselect("📅 Mes", sorted(df_ef['Mes'].dropna().unique()))
    sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS TÉCNICOS (CON FÓRMULAS)
# =========================================================================
h_std = safe_calc(dff['HH_STD'])
h_disp = safe_calc(dff['HH_Disp'])
h_prod = safe_calc(dff['HH_Prod'])
costo_tot = safe_calc(dff['Costo'])

# Disponibilidad = Tiempo Operativo / Tiempo Disponible
disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0

# Rendimiento = Tiempo Standard Producido / Tiempo Operativo Real
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0

# Calidad = (Total - Descarte) / Total -> Simulado 98%
calidad_sim = 98.0

# OEE = Disponibilidad x Rendimiento x Calidad
oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100

# Limitar visualmente
disponibilidad_v = min(disponibilidad, 100.0)
rendimiento_v = min(rendimiento, 100.0)

# Color por Performance
color_oee = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 60 else "#ef4444"

# =========================================================================
# 6. INTEGRACIÓN: EL CARTEL DE OEE PREMIUM CON FÓRMULAS
# =========================================================================
st.markdown(f"""
<div class='oee-hero'>
    <p class='oee-formula-global'>OEE = Disponibilidad &times; Rendimiento &times; Calidad</p>
    <h1 class='oee-main-val' style='color: {color_oee};'>{oee_final:.1f}%</h1>
    <div class='bar-bg'>
        <div class='bar-fill' style='width: {min(max(oee_final,0),100)}%; background: {color_oee};'></div>
    </div>
    
    <div class='pillar-grid'>
        <div class='pillar-card'>
            <p class='pillar-name'>⏱️ Disponibilidad</p>
            <p class='pillar-val'>{disponibilidad_v:.1f}%</p>
            <div class='pillar-formula'>
                <b>Fórmula:</b><br>HH Productivas / HH Disponibles
            </div>
        </div>
        
        <div class='pillar-card'>
            <p class='pillar-name'>🚀 Rendimiento</p>
            <p class='pillar-val'>{rendimiento_v:.1f}%</p>
            <div class='pillar-formula'>
                <b>Fórmula:</b><br>HH Standard / HH Productivas
            </div>
        </div>
        
        <div class='pillar-card'>
            <p class='pillar-name'>💎 Calidad</p>
            <p class='pillar-val'>{calidad_sim:.1f}%</p>
            <div class='pillar-formula'>
                <b>Fórmula:</b><br>(Piezas OK / Total) <span style='color:red;'>(SIM)</span>
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KPIs de Apoyo
c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f"<div class='kpi-sub'><small>EFIC. REAL</small><h3>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h3></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-sub'><small>HH PRODUCIDAS</small><h3>{h_std:,.0f}</h3></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-sub' style='border-bottom-color:red;'><small>COSTO INEF.</small><h3 style='color:red;'>${costo_tot:,.0f}</h3></div>", unsafe_allow_html=True)
with c4: st.markdown(f"<div class='kpi-sub' style='border-bottom-color:orange;'><small>GAP HORAS</small><h3 style='color:orange;'>{int(h_disp - h_prod)}h</h3></div>", unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS INTEGRADOS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
col_g1, col_g2 = st.columns([2, 1])

with col_g1:
    st.subheader("📈 Tendencia Mensual de Eficiencia")
    if not dff.empty:
        fig, ax = plt.subplots(figsize=(10, 4))
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax.plot(ag.index, ag['Ef'], marker='o', color='#1e3a8a', linewidth=4, markersize=10)
        ax.fill_between(ag.index, ag['Ef'], color='#1e3a8a', alpha=0.1)
        ax.axhline(85, color='green', linestyle='--', alpha=0.6, label="Meta Clase Mundial")
        ax.set_ylim(0, 110)
        ax.legend()
        st.pyplot(fig)

with col_g2:
    st.subheader("⚠️ Top Causas de Parada")
    if 'HH_Imp' in df_im.columns:
        res = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(5, 6.5))
        res.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis()
        ax2.set_title("Horas Perdidas")
        st.pyplot(fig2)

st.markdown("---")
st.caption("Sistema de Control Industrial - Ombú S.A. 2026")
