import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt

# =========================================================================
# 1. DISEÑO DE INTERFAZ "PLATINUM" (CSS CUSTOM)
# =========================================================================
st.set_page_config(page_title="Gestión Industrial Ombú", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    header, [data-testid="stHeader"], footer {display: none !important;}
    .block-container {padding-top: 1rem !important; background-color: #0f172a;} 

    .ombu-header {
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 20px; border-radius: 15px; color: white;
        text-align: center; margin-bottom: 25px;
        box-shadow: 0px 10px 20px rgba(0,0,0,0.3);
    }

    .gauge-card {
        background-color: #1e293b; 
        border-radius: 20px; padding: 25px; margin-bottom: 20px;
        border: 1px solid #334155; box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
    }

    .pillar-row { display: flex; justify-content: space-around; gap: 15px; margin-top: 10px; }
    .pillar-box { 
        background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
        border: 1px solid #334155; text-align: center;
    }
    .p-title { font-size: 13px; font-weight: 700; color: #94a3b8; text-transform: uppercase; margin-bottom: 5px; }
    .p-val { font-size: 32px; font-weight: 900; color: white; margin: 0; }
    .p-formula { font-size: 11px; color: #3b82f6; font-style: italic; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; }

    .kpi-main {
        background: #1e293b; padding: 20px; border-radius: 15px;
        border-left: 6px solid #3b82f6; text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px;
    }
    .kpi-main h2 { color: white !important; font-size: 35px; margin: 5px 0; }
    .kpi-main h6 { color: #94a3b8 !important; margin: 0; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<div style='text-align:center; margin-top: 100px;'><h1 style='color:white;'>🏭 SISTEMA OMBÚ</h1><p style='color:#94a3b8;'>Control de Gestión de Planta</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("ACCEDER SISTEMA", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS (ALTA RESISTENCIA)
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
    
    df_ef = pd.read_excel(u_ef)
    df_im = pd.read_excel(u_im)
    
    def normalize_df(df):
        df.columns = [str(c).strip().upper() for c in df.columns]
        m = {}
        for c in df.columns:
            if 'FECHA' in c: m[c] = 'Fecha'
            elif 'PLANTA' in c: m[c] = 'Planta'
            elif 'STD' in c: m[c] = 'HH_STD'
            elif 'DISP' in c: m[c] = 'HH_Disp'
            elif 'PROD' in c or 'GAP' in c: m[c] = 'HH_Prod'
            elif 'COSTO' in c: m[c] = 'Costo'
        df = df.rename(columns=m)
        return df.loc[:, ~df.columns.duplicated()]

    df_ef = normalize_df(df_ef)
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c not in df_ef.columns: df_ef[c] = 0.0
        df_ef[c] = pd.to_numeric(df_ef[c], errors='coerce').fillna(0)
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce')
    df_ef['Mes'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    m_im = {}
    for c in df_im.columns:
        if 'HH' in c and 'IMP' in c: m_im[c] = 'HH_Imp'
        elif 'TIPO' in c or 'MOTIVO' in c: m_im[c] = 'Motivo'
    df_im = df_im.rename(columns=m_im)
    if 'HH_Imp' in d_im.columns if 'd_im' in locals() else 'HH_Imp' in df_im.columns: 
        df_im['HH_Imp'] = pd.to_numeric(df_im['HH_Imp'], errors='coerce').fillna(0)

    return df_ef, df_im

try:
    df_ef, df_im = load_data()
except Exception as e:
    st.error(f"Error cargando Drive: {e}"); st.stop()

# =========================================================================
# 4. HEADER Y FILTROS
# =========================================================================
st.markdown("<div class='ombu-header'><h1>REPORTING INDUSTRIAL CGP v1.0</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR_pD_pZ48aE7rE-j9J2N7xXgXp6qZp3T6w2A&s", width=150)
    st.header("⚙️ CONFIGURACIÓN")
    f_mes = st.multiselect("📅 Seleccionar Mes", sorted(df_ef['Mes'].dropna().unique()))
    f_planta = st.multiselect("🏭 Seleccionar Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if f_mes: dff = dff[dff['Mes'].isin(f_mes)]
if f_planta: dff = dff[dff['Planta'].isin(f_planta)]

# =========================================================================
# 5. CÁLCULOS TÉCNICOS OEE
# =========================================================================
h_std = safe_calc(dff['HH_STD'])
h_disp = safe_calc(dff['HH_Disp'])
h_prod = safe_calc(dff['HH_Prod'])
costo_tot = safe_calc(dff['Costo'])

disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
calidad_sim = 98.0

oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100
color_oee = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 65 else "#ef4444"

# =========================================================================
# 6. EL CARTEL DE OEE (VELOCÍMETRO + FÓRMULAS)
# =========================================================================
st.markdown("<div class='gauge-card'>", unsafe_allow_html=True)

# Velocímetro Plotly
fig = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    number = {'suffix': "%", 'font': {'color': 'white', 'size': 80}},
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
fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350, margin=dict(l=20, r=20, t=50, b=0))
st.plotly_chart(fig, use_container_width=True)

# Pilares con Fórmulas Detalladas
st.markdown(f"""
    <div class='pillar-row'>
        <div class='pillar-box'>
            <p class='p-title'>⏱️ Disponibilidad</p>
            <p class='p-val'>{min(disponibilidad, 100.0):.1f}%</p>
            <p class='p-formula'>HH Productivas / HH Disponibles</p>
        </div>
        <div class='pillar-box'>
            <p class='p-title'>🚀 Rendimiento</p>
            <p class='p-val'>{min(rendimiento, 100.0):.1f}%</p>
            <p class='p-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-box'>
            <p class='p-title'>💎 Calidad</p>
            <p class='p-val'>{calidad_sim:.1f}%</p>
            <p class='p-formula'>Piezas OK / Total <span style='color:#ef4444'>(SIM)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. INTEGRACIÓN CON EL "MAIN" (MÉTRICAS DE SOPORTE)
# =========================================================================
c1, c2, c3 = st.columns(3)
with c1: 
    st.markdown(f"<div class='kpi-main'><h6>EFICIENCIA REAL</h6><h2 style='color:#3b82f6;'>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h2></div>", unsafe_allow_html=True)
with c2: 
    st.markdown(f"<div class='kpi-main' style='border-left-color:#22c55e;'><h6>TOTAL PRODUCIDO</h6><h2>{h_std:,.0f} <small style='font-size:15px; color:#94a3b8;'>HH STD</small></h2></div>", unsafe_allow_html=True)
with c3: 
    st.markdown(f"<div class='kpi-main' style='border-left-color:#ef4444;'><h6>COSTO INEFICIENCIA</h6><h2 style='color:#ef4444;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS Y TENDENCIAS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g1, g2 = st.columns([2, 1])

with g1:
    if not dff.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 4), facecolor='#0f172a')
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax1.plot(ag.index, ag['Ef'], marker='o', color='#3b82f6', linewidth=4, markersize=10)
        ax1.set_title("Evolución Eficiencia Real (%)", color='white', pad=20)
        ax1.set_ylim(0, 110); ax1.tick_params(colors='white')
        ax1.grid(axis='y', linestyle='--', alpha=0.2)
        st.pyplot(fig1)

with g2:
    if 'HH_Imp' in df_im.columns:
        fig2, ax2 = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_imp = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis(); ax2.set_title("Top 5 Pérdidas (HH)", color='white')
        ax2.tick_params(colors='white')
        st.pyplot(fig2)

st.caption("OMBU Industrial Intelligence © 2026")
