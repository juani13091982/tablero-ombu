import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =========================================================================
# 1. DISEÑO DE INTERFAZ "MAESTRO" (CSS CUSTOM + DARK BLUE THEME)
# =========================================================================
st.set_page_config(page_title="Tablero CGP Pro - Ombú S.A.", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ocultar basura de Streamlit */
    header, [data-testid="stHeader"], footer {display: none !important;}
    .block-container {padding-top: 1rem !important; background-color: #0f172a;} /* Fondo Oscuro Profundo */
    
    /* Header multinacional */
    .ombu-header {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 20px; border-radius: 10px; color: white;
        text-align: center; margin-bottom: 25px;
        box-shadow: 0px 10px 20px rgba(0,0,0,0.3);
    }
    .ombu-header h1 { color: white !important; font-size: 30px !important; margin:0;}

    /* Contenedor del Velocímetro y Pilares (DISEÑO INTEGRADO) */
    .gauge-wrapper {
        background-color: #1e293b; /* Fondo tarjeta oscuro */
        border-radius: 20px; padding: 2rem; margin-bottom: 2rem;
        box-shadow: 0px 20px 25px rgba(0,0,0,0.1); border: 1px solid #334155;
    }

    /* Pilares OEE */
    .pillar-row { display: flex; justify-content: space-around; gap: 1rem; margin-top: 1.5rem; }
    .pillar-box { 
        background: #0f172a; border-radius: 15px; padding: 1.2rem; flex: 1;
        border: 1px solid #334155; text-align: center;
    }
    .pillar-name { font-size: 13px; font-weight: 700; color: #a0aec0; text-transform: uppercase; }
    .pillar-val { font-size: 32px; font-weight: 900; color: white; margin: 0.5rem 0; }
    .pillar-formula { font-size: 11px; color: #718096; font-style: italic; border-top: 1px solid #334155; padding-top: 5px; }

    /* Tarjetas de Apoyo */
    .kpi-mini-card {
        background: #1e293b; padding: 20px; border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 5px solid #3B82F6;
        text-align: center; color: white;
    }
    .kpi-mini-card h2 { color: white !important; font-size: 35px; }
    .kpi-mini-card h6 { color: #a0aec0 !important; font-size: 13px; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<br><br><div style='text-align:center;'><h1>🏭 OMBÚ S.A.</h1><p>Sistema de Gestión Pro</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("ACCEDER SISTEMA"):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("Acceso denegado")
    st.stop()

# =========================================================================
# 3. CARGA DE DATOS Y LIMPIEZA (MOTOR BLINDADO v8.0)
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
    d_ef = pd.read_excel(u_ef); d_im = pd.read_excel(u_im)
    
    # Normalizador Inteligente
    def normalizar(df):
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

    d_ef = normalizar(d_ef)
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c not in d_ef.columns: d_ef[c] = 0.0
        d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
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
    st.error(f"Error cargando Drive: {e}"); st.stop()

# =========================================================================
# 4. FILTROS (DISEÑO MÁS AGRADABLE)
# =========================================================================
st.markdown("<div class='ombu-header'><h1>REPORTING INDUSTRIAL CGP v1.0</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/e/e6/Placeholder_OMBU.png", width=150)
    st.header("⚙️ CONFIGURACIÓN")
    f_mes = st.multiselect("📅 Mes", sorted(df_ef['Mes'].dropna().unique()))
    f_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if f_mes: dff = dff[dff['Mes'].isin(f_mes)]
if f_planta: dff = dff[dff['Planta'].isin(f_planta)]

# =========================================================================
# 5. CÁLCULOS OEE Y FÓRMULAS
# =========================================================================
h_std = safe_calc(dff['HH_STD'])
h_disp = safe_calc(dff['HH_Disp'])
h_prod = safe_calc(dff['HH_Prod'])
costo_tot = safe_calc(dff['Costo'])

disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
calidad_sim = 98.0
oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100

# =========================================================================
# 6. INTEGRACIÓN: EL VELOCÍMETRO (GAUGE) Y PILARES
# =========================================================================
st.markdown("<div class='gauge-wrapper'>", unsafe_allow_html=True)

# --- VELOCÍMETRO (PLOTLY) ---
fig = go.Figure(go.Indicator(
    mode = "gauge+number",
    value = oee_final,
    title = {'text': "OEE GENERAL (%)", 'font': {'color': 'white', 'size': 18}},
    number = {'valueformat': ".1f", 'font': {'color': 'white'}},
    gauge = {
        'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "white"},
        'bar': {'color': "#1E3A8A"},
        'bgcolor': "white",
        'borderwidth': 2,
        'bordercolor': "#f4f7f9",
        'steps': [
            {'range': [0, 60], 'color': '#ef4444'}, # Rojo
            {'range': [60, 85], 'color': '#eab308'}, # Amarillo
            {'range': [85, 100], 'color': '#22c55e'} # Verde
        ],
        'threshold': {
            'line': {'color': "white", 'width': 4},
            'thickness': 0.75,
            'value': 85
        }
    }
))
fig.update_layout(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font={'color': "white", 'family': "sans-serif"},
    margin=dict(l=20, r=20, t=20, b=20),
    height=300
)
st.plotly_chart(fig, use_container_width=True)

# --- PILARES Y FÓRMULAS ESCRITAS ---
st.markdown(f"""
    <div class='pillar-row'>
        <div class='pillar-box'>
            <p class='pillar-name'>⏱️ Disponibilidad</p>
            <p class='pillar-val'>{min(disponibilidad, 100.0):.1f}%</p>
            <p class='pillar-formula'>HH Productivas / HH Disponibles</p>
        </div>
        <div class='pillar-box'>
            <p class='pillar-name'>🚀 Rendimiento</p>
            <p class='pillar-val'>{min(rendimiento, 100.0):.1f}%</p>
            <p class='pillar-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-box'>
            <p class='pillar-name'>💎 Calidad</p>
            <p class='pillar-val'>{calidad_sim:.1f}%</p>
            <p class='pillar-formula'>(Piezas OK / Total) <span style='color:red;'>(SIM)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KPIs Secundarios
k1, k2, k3 = st.columns(3)
with k1: st.markdown(f"<div class='kpi-mini-card'><h6>EFICIENCIA REAL</h6><h2 style='color:#3B82F6;'>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h2></div>", unsafe_allow_html=True)
with k2: st.markdown(f"<div class='kpi-mini-card' style='border-left-color:#22c55e;'><h6>HH PRODUCIDAS</h6><h2 style='color:#22c55e;'>{h_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-mini-card' style='border-left-color:#ef4444;'><h6>COSTO INEFICIENCIA</h6><h2 style='color:#ef4444;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
c_gr1, c_gr2 = st.columns([2, 1])

with c_gr1:
    if not dff.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 4.5), facecolor='#0f172a')
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax1.plot(ag.index, ag['Ef'], marker='o', color='#3B82F6', linewidth=4, markersize=10)
        # ax1.set_ylabel("Standard", color='white')
        ax1.set_title("Tendencia Eficiencia Real", color='white')
        ax1.set_ylim(0, 110); ax1.tick_params(colors='white')
        st.pyplot(fig1)

with c_gr2:
    if 'HH_Imp' in df_im.columns:
        fig2, ax2 = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_imp = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        res_imp.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis(); ax2.set_title("Top Pérdidas", color='white')
        ax2.tick_params(colors='white')
        st.pyplot(fig2)

st.caption("OMBU v8.0 PRO Online")
