import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =========================================================================
# 1. ESTÉTICA INTEGRADA "MODERNA" (CSS)
# =========================================================================
st.set_page_config(page_title="Tablero Integrado Ombú", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Estética Global */
    header, [data-testid="stHeader"], footer {display: none !important;}
    .main { background-color: #f4f7f9; }
    
    /* Título Superior */
    .header-band {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 1.5rem; border-radius: 0 0 20px 20px;
        color: white; text-align: center; margin-bottom: 2rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* EL CARTEL DE OEE (INTEGRACIÓN AGRADABLE) */
    .oee-container {
        background: white; border-radius: 20px; padding: 2rem;
        border: 1px solid #e2e8f0; box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        margin-bottom: 2rem; text-align: center;
    }
    .oee-title { color: #64748b; font-size: 0.9rem; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; }
    .oee-main-value { font-size: 5rem; font-weight: 900; margin: 0.5rem 0; line-height: 1; }
    .oee-formula-main { font-size: 1rem; color: #94a3b8; font-style: italic; margin-bottom: 1.5rem; }

    /* Los 3 Pilares del OEE */
    .pillar-row { display: flex; justify-content: space-between; gap: 1rem; margin-top: 1rem; }
    .pillar-box { 
        background: #f8fafc; border-radius: 15px; padding: 1.5rem; flex: 1;
        border: 1px solid #f1f5f9; transition: 0.3s;
    }
    .pillar-box:hover { background: #ffffff; transform: translateY(-5px); box-shadow: 0 10px 15px rgba(0,0,0,0.05); }
    .pillar-name { font-size: 0.85rem; font-weight: 700; color: #1e3a8a; text-transform: uppercase; }
    .pillar-val { font-size: 2rem; font-weight: 800; color: #1e293b; margin: 0.5rem 0; }
    .pillar-formula-text { font-size: 0.75rem; color: #94a3b8; line-height: 1.2; border-top: 1px dashed #cbd5e1; padding-top: 0.5rem; }

    /* Tarjetas de Métricas Secundarias */
    .kpi-metric-card {
        background: white; padding: 1.5rem; border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02); border-left: 5px solid #1E3A8A;
        text-align: center;
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
        st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏭 SISTEMA DE GESTIÓN OMBÚ</h1><p>Acceso Protegido</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR"):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS (VERSION DE ALTA RESISTENCIA)
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
    
    # Mapeo Inteligente para no fallar con "Costo", "HH", etc.
    def smart_normalize(df):
        df.columns = [str(c).strip().upper() for c in df.columns]
        map_cols = {}
        for c in df.columns:
            if 'FECHA' in c: map_cols[c] = 'Fecha'
            elif 'PLANTA' in c: map_cols[c] = 'Planta'
            elif 'STD' in c: map_cols[c] = 'HH_STD'
            elif 'DISP' in c: map_cols[c] = 'HH_Disp'
            elif 'PROD' in c or 'GAP' in c: map_cols[c] = 'HH_Prod'
            elif 'COSTO' in c: map_cols[c] = 'Costo'
        df = df.rename(columns=map_cols)
        return df.loc[:, ~df.columns.duplicated()]

    d_ef = smart_normalize(d_ef)
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c not in d_ef.columns: d_ef[c] = 0.0
        d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
    # Improductivas
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    m_im = {}
    for c in d_im.columns:
        if 'HH' in c and 'IMP' in c: m_im[c] = 'HH_Imp'
        elif 'TIPO' in c or 'MOTIVO' in c: m_im[c] = 'Motivo'
    d_im = d_im.rename(columns=m_im)
    if 'HH_Imp' in d_im.columns: d_im['HH_Imp'] = pd.to_numeric(d_im['HH_Imp'], errors='coerce').fillna(0)

    return d_ef, d_im

try:
    df_ef, df_im = load_data()
except Exception as e:
    st.error(f"Error de conexión con Excel: {e}"); st.stop()

# =========================================================================
# 4. FILTROS
# =========================================================================
st.markdown("<div class='header-band'><h1>REPORTE INTEGRADO: EFICIENCIA & OEE</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://www.maquinariapropia.com.ar/uploads/marcas/logo-ombu.png", width=150)
    st.header("CONTROLES")
    sel_mes = st.multiselect("📅 Seleccionar Mes", sorted(df_ef['Mes'].dropna().unique()))
    sel_planta = st.multiselect("🏭 Seleccionar Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS TÉCNICOS (EXPLICATIVOS)
# =========================================================================
h_std = safe_calc(dff['HH_STD'])
h_disp = safe_calc(dff['HH_Disp'])
h_prod = safe_calc(dff['HH_Prod'])
costo_tot = safe_calc(dff['Costo'])

# Disponibilidad: HH Productivas / HH Disponibles
disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
# Rendimiento: HH Standard Producidas / HH Productivas Reales
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
# Calidad: Simulado al 98%
calidad_sim = 98.0

# OEE Final = D x R x C
oee_val = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100
color_oee = "#22c55e" if oee_val >= 85 else "#eab308" if oee_val >= 60 else "#ef4444"

# =========================================================================
# 6. INTEGRACIÓN: EL CARTEL DE OEE PREMIUM CON FÓRMULAS
# =========================================================================
st.markdown(f"""
<div class='oee-container'>
    <p class='oee-title'>OEE: Efectividad General de los Equipos</p>
    <p class='oee-formula-main'>Fórmula: Disponibilidad &times; Rendimiento &times; Calidad</p>
    <h1 class='oee-main-value' style='color: {color_oee};'>{oee_val:.1f}%</h1>
    
    <div style='background: #e2e8f0; height: 12px; border-radius: 10px; width: 60%; margin: 1.5rem auto; overflow: hidden;'>
        <div style='background: {color_oee}; width: {min(max(oee_val,0),100)}%; height: 100%; transition: 1.5s;'></div>
    </div>

    <div class='pillar-row'>
        <div class='pillar-box'>
            <p class='pillar-name'>⏱️ Disponibilidad</p>
            <p class='pillar-val'>{min(disponibilidad, 100.0):.1f}%</p>
            <p class='pillar-formula-text'>HH Productivas / HH Disponibles</p>
        </div>
        <div class='pillar-box'>
            <p class='pillar-name'>🚀 Rendimiento</p>
            <p class='pillar-val'>{min(rendimiento, 100.0):.1f}%</p>
            <p class='pillar-formula-text'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-box'>
            <p class='pillar-name'>💎 Calidad</p>
            <p class='pillar-val'>{calidad_sim:.1f}%</p>
            <p class='pillar-formula-text'>(Piezas OK / Total) <span style='color:red;'>(SIM)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KPIs Secundarios (Integración del Main)
k1, k2, k3 = st.columns(3)
with k1: st.markdown(f"<div class='kpi-metric-card'><h6>EFICIENCIA REAL</h6><h2 style='color:#1E3A8A;'>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h2></div>", unsafe_allow_html=True)
with k2: st.markdown(f"<div class='kpi-metric-card' style='border-left-color:#16a34a;'><h6>HH PRODUCIDAS</h6><h2 style='color:#16a34a;'>{h_std:,.0f}</h2></div>", unsafe_allow_html=True)
with k3: st.markdown(f"<div class='kpi-metric-card' style='border-left-color:#dc2626;'><h6>COSTO INEFICIENCIA</h6><h2 style='color:#dc2626;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS Y DETALLE
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
c_g1, c_g2 = st.columns([2, 1])

with c_g1:
    st.subheader("📈 Tendencia Mensual de Eficiencia Real")
    if not dff.empty:
        fig, ax = plt.subplots(figsize=(10, 4.5), facecolor='#f4f7f9')
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax.plot(ag.index, ag['Ef'], marker='o', color='#1e3a8a', linewidth=4, markersize=10, label='% Eficiencia')
        ax.fill_between(ag.index, ag['Ef'], color='#1e3a8a', alpha=0.1)
        ax.axhline(85, color='green', linestyle='--', alpha=0.5, label='Meta Mundial (85%)')
        ax.set_ylim(0, 110)
        ax.legend()
        st.pyplot(fig)

with c_g2:
    st.subheader("⚠️ Top Causas de Parada")
    if 'HH_Imp' in df_im.columns:
        res_imp = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(6, 8.5), facecolor='#f4f7f9')
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis()
        ax2.set_title("Horas Totales Perdidas")
        st.pyplot(fig2)

st.markdown("---")
st.caption("Reporting Industrial Ombú S.A. - 2026")
