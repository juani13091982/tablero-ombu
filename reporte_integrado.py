import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =========================================================================
# 1. DISEÑO DE INTERFAZ "PREMIUM" (CSS CUSTOM)
# =========================================================================
st.set_page_config(page_title="Dashboard Industrial Ombú", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ocultar elementos de Streamlit */
    header, [data-testid="stHeader"], footer {display: none !important;}
    .block-container {padding-top: 1rem !important; background-color: #f1f5f9;}

    /* Encabezado */
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%);
        padding: 25px; border-radius: 15px; color: white;
        text-align: center; margin-bottom: 25px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }

    /* Cartel OEE Maestro */
    .oee-box {
        background: white; border-radius: 20px; padding: 40px;
        text-align: center; border: 1px solid #e2e8f0;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1); margin-bottom: 30px;
    }
    .oee-title { color: #64748b; font-size: 14px; font-weight: 800; letter-spacing: 0.1em; text-transform: uppercase; }
    .oee-value { font-size: 100px; font-weight: 900; margin: 10px 0; line-height: 1; }
    
    /* Progress Bar OEE */
    .oee-bar-bg { background: #f1f5f9; height: 18px; border-radius: 10px; width: 70%; margin: 20px auto; overflow: hidden; border: 1px solid #e2e8f0; }
    .oee-bar-fill { height: 100%; border-radius: 10px; transition: width 1s ease-in-out; }

    /* Pilares OEE */
    .pillar-row { display: flex; justify-content: space-around; margin-top: 30px; gap: 20px; }
    .pillar-card { background: #f8fafc; padding: 20px; border-radius: 15px; flex: 1; border: 1px solid #f1f5f9; }
    .pillar-val { font-size: 28px; font-weight: 800; color: #1e293b; }
    .pillar-lbl { font-size: 12px; color: #94a3b8; font-weight: 600; text-transform: uppercase; }

    /* Mini Cards */
    .kpi-mini {
        background: white; padding: 20px; border-radius: 15px;
        border-left: 6px solid #1e3a8a; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
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
        st.markdown("<div style='text-align:center; margin-top: 100px;'><h1>🏭 OMBÚ S.A.</h1><p>Control de Gestión de Planta</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("ACCEDER SISTEMA", use_container_width=True):
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
    
    # Normalizador Inteligente de Columnas
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
    
    # Asegurar columnas aunque no existan
    for c in ['Fecha', 'Planta', 'HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c not in d_ef.columns: d_ef[c] = 0 if c != 'Fecha' and c != 'Planta' else 'S/D'
        if c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
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
    st.error(f"Error en datos: {e}"); st.stop()

# =========================================================================
# 4. FILTROS
# =========================================================================
st.markdown("<div class='main-header'><h1>REPORTING INDUSTRIAL CGP - OMBÚ</h1></div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ CONFIGURACIÓN")
    sel_mes = st.multiselect("📅 Mes", sorted(df_ef['Mes'].dropna().unique()))
    sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))

dff = df_ef.copy()
if sel_mes: dff = dff[dff['Mes'].isin(sel_mes)]
if sel_planta: dff = dff[dff['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS OEE
# =========================================================================
h_std = safe_calc(dff['HH_STD'])
h_disp = safe_calc(dff['HH_Disp'])
h_prod = safe_calc(dff['HH_Prod'])
costo_tot = safe_calc(dff['Costo'])

disponibilidad = (h_prod / h_disp * 100) if h_disp > 0 else 0.0
rendimiento = (h_std / h_prod * 100) if h_prod > 0 else 0.0
calidad_sim = 98.0
oee_final = (disponibilidad/100 * rendimiento/100 * calidad_sim/100) * 100

# Colores dinámicos
color_oee = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 60 else "#ef4444"

# =========================================================================
# 6. EL CARTEL OEE "HERMOSO" (MAQUETADO HTML)
# =========================================================================
st.markdown(f"""
<div class='oee-box'>
    <p class='oee-title'>OEE: Efectividad General de los Equipos</p>
    <h1 class='oee-value' style='color: {color_oee};'>{oee_final:.1f}%</h1>
    <div class='oee-bar-bg'>
        <div class='oee-bar-fill' style='width: {min(max(oee_final,0),100)}%; background: {color_oee};'></div>
    </div>
    <div class='pillar-row'>
        <div class='pillar-card'>
            <p class='pillar-lbl'>⏱️ Disponibilidad</p>
            <p class='pillar-val'>{min(disponibilidad, 100.0):.1f}%</p>
        </div>
        <div class='pillar-card'>
            <p class='pillar-lbl'>🚀 Rendimiento</p>
            <p class='pillar-val'>{min(rendimiento, 100.0):.1f}%</p>
        </div>
        <div class='pillar-card'>
            <p class='pillar-lbl'>💎 Calidad (Sim)</p>
            <p class='pillar-val'>{calidad_sim:.1f}%</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Mini KPIs
c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div class='kpi-mini'><h6>EFICIENCIA REAL</h6><h2 style='color:#1e3a8a;'>{(h_std/h_disp*100 if h_disp>0 else 0):.1f}%</h2></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='kpi-mini' style='border-left-color:#16a34a;'><h6>TOTAL PRODUCIDO</h6><h2 style='color:#16a34a;'>{h_std:,.0f} <span style='font-size:14px;'>HH STD</span></h2></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div class='kpi-mini' style='border-left-color:#dc2626;'><h6>COSTO INEFICIENCIA</h6><h2 style='color:#dc2626;'>${costo_tot:,.0f}</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
col_g1, col_g2 = st.columns([2, 1])

with col_g1:
    st.subheader("📈 Tendencia de Eficiencia Real")
    if not dff.empty:
        fig, ax = plt.subplots(figsize=(10, 3.5))
        ag = dff.groupby('Mes').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag['Ef'] = (ag['HH_STD'] / ag['HH_Disp'] * 100)
        ax.plot(ag.index, ag['Ef'], marker='o', color='#1e3a8a', linewidth=3, markersize=8)
        ax.fill_between(ag.index, ag['Ef'], color='#1e3a8a', alpha=0.1)
        ax.set_ylim(0, 110)
        ax.grid(axis='y', linestyle='--', alpha=0.5)
        st.pyplot(fig)

with col_g2:
    st.subheader("⚠️ Top Pérdidas")
    if 'HH_Imp' in df_im.columns:
        res = df_im.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(5, 6))
        res.plot(kind='barh', color='#dc2626', ax=ax2)
        ax2.invert_yaxis()
        st.pyplot(fig2)

st.success("✅ Tablero Premium Ombú v6.0 Online. ¡Esto sí es nivel mundial!")
