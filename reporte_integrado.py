import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import matplotlib.image as mpimg
import textwrap
import re

# =========================================================================
# 1. CONFIGURACIÓN Y ESCUDO VISUAL (RESPONSIVO)
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
    header, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="manage-app-button"], 
    #MainMenu, footer, .stAppDeployButton, .viewerBadge_container {display: none !important; visibility: hidden !important;}
    .block-container {padding-top: 1rem !important; padding-bottom: 1.5rem !important;}
    div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {
        position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
        background-color: rgba(14, 17, 23, 0.98) !important; z-index: 99999 !important;
        padding: 5px 10px 15px 10px !important; border-bottom: 2px solid #1E3A8A !important;
        box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
    }
    .kpi-grid { display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 12px; }
    .kpi-costo { grid-row: span 2; }
    .mobile-only { display: none !important; }
    @media (max-width: 768px) {
        .mobile-only { display: block !important; }
        [data-testid="stHorizontalBlock"] { flex-direction: column !important; }
        [data-testid="stHorizontalBlock"] > div { width: 100% !important; min-width: 100% !important; }
        .kpi-grid { grid-template-columns: 1fr !important; }
        .kpi-costo { grid-row: span 1 !important; }
    }
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold'})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o, caja_b = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if not st.session_state['autenticado']:
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='text-align:center;'><h2>GESTIÓN INDUSTRIAL OMBÚ</h2></div>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['autenticado'] = True
                    st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE Y CARGA (CON SUPER ESCUDO)
# =========================================================================
def safe_match(s_list, val):
    if pd.isna(val): return False
    v_norm = re.sub(r'[^A-Z0-9]', '', str(val).upper())
    for s in s_list:
        if re.sub(r'[^A-Z0-9]', '', str(s).upper()) == v_norm: return True
    return False

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if '✅' in d: return "🎯 ACCIÓN GLOBAL"
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO', 'MECANICO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'ESPERA', 'LOGISTICA']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'ERROR']): return "🔎 Ajustar Calidad"
    return "⚡ Investigar Causa"

try:
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    df_ef = pd.read_excel(url_ef)
    df_im = pd.read_excel(url_im)

    def normalizar_ef(df):
        df.columns = [str(c).strip().upper() for c in df.columns]
        mapping = {}
        for c in df.columns:
            if 'FECHA' in c: mapping[c] = 'Fecha'
            elif 'PLANTA' in c: mapping[c] = 'Planta'
            elif 'LÍNEA' in c or 'LINEA' in c: mapping[c] = 'Linea'
            elif 'ULTIMO' in c or 'ÚLTIMO' in c: mapping[c] = 'Es_Ultimo_Puesto'
            elif 'PUESTO' in c and 'Puesto_Trabajo' not in mapping.values(): mapping[c] = 'Puesto_Trabajo'
            elif 'CANT' in c and 'PROD' in c: mapping[c] = 'Cant_Prod'
            elif 'STD' in c: mapping[c] = 'HH_STD'
            elif 'DISP' in c: mapping[c] = 'HH_Disp'
            elif 'GAP' in c: mapping[c] = 'HH_Prod_GAP'
            elif 'PROD' in c and 'HH_Productivas' not in mapping.values(): mapping[c] = 'HH_Productivas'
            elif 'COSTO' in c: mapping[c] = 'Costo'
        df = df.rename(columns=mapping)
        return df.loc[:, ~df.columns.duplicated()]

    df_ef = normalizar_ef(df_ef)
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    # Arreglos de seguridad
    if 'HH_IMPRODUCTIVAS' not in df_im.columns:
        col_imp = [c for c in df_im.columns if 'HH' in c and 'IMP' in c]
        df_im['HH_IMPRODUCTIVAS'] = df_im[col_imp[0]] if col_imp else 0
    if 'TIPO_PARADA' not in df_im.columns:
        col_t = [c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c]
        df_im['TIPO_PARADA'] = df_im[col_t[0]] if col_t else "S/D"

    for col in ['HH_STD', 'HH_Disp', 'HH_Prod_GAP', 'HH_Productivas', 'Costo', 'Cant_Prod']:
        if col in df_ef.columns: df_ef[col] = pd.to_numeric(df_ef[col], errors='coerce').fillna(0)
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce')
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    
    # Puesto en improductivas
    c_pu_im = next((c for c in df_im.columns if 'PUESTO' in c), None)

except Exception as e:
    st.error(f"Error crítico: {e}"); st.stop()

# =========================================================================
# 4. FILTROS EN CASCADA
# =========================================================================
st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
with st.container():
    c_kpi, c_fil = st.columns([3.5, 1])
    with c_fil:
        st.markdown("**🎛️ FILTROS**")
        s_mes = st.multiselect("Mes", sorted(df_ef['Mes_Str'].dropna().unique()), placeholder="📅 Mes")
        s_pl = st.multiselect("Planta", sorted(df_ef['Planta'].dropna().unique()), placeholder="🏭 Planta")

df_f = df_ef.copy()
if s_mes: df_f = df_f[df_f['Mes_Str'].isin(s_mes)]
if s_pl: df_f = df_f[df_f['Planta'].isin(s_pl)]

# =========================================================================
# 5. CÁLCULOS Y KPIs (CONVERTIDOS A FLOAT PARA EVITAR ERRORES)
# =========================================================================
t_std = float(df_f['HH_STD'].sum())
t_disp = float(df_f['HH_Disp'].sum())
t_prod_gap = float(df_f['HH_Prod_GAP'].sum())
t_costo = float(df_f['Costo'].sum())

ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0.0
ef_prod = (t_std / t_prod_gap * 100) if t_prod_gap > 0 else 0.0
oee = ef_real * 0.98  # Calidad simulada al 98%

# =========================================================================
# 6. DISEÑO DE INTERFAZ (KPIs + OEE)
# =========================================================================
with c_kpi:
    # --- MÓDULO OEE ---
    color_oee = "#4CAF50" if oee >= 85 else ("#FFEB3B" if oee >= 60 else "#F44336")
    st.markdown(f"""
    <div style="background: white; border: 2px solid #1E3A8A; padding: 15px; border-radius: 10px; text-align: center; margin-bottom: 15px; box-shadow: 2px 2px 10px rgba(0,0,0,0.1);">
        <h3 style="margin: 0; color: #1E3A8A;">🏆 OEE: Eficiencia General <span style="font-size:14px; color:red;">(SIMULADO)</span></h3>
        <p style="margin: 5px; font-size: 14px;">Disponibilidad &times; Rendimiento &times; Calidad (98%)</p>
        <div style="background: #eee; border-radius: 20px; height: 30px; width: 100%; border: 1px solid #ccc;">
            <div style="background: {color_oee}; width: {min(oee, 100):.1f}%; height: 100%; border-radius: 20px; display: flex; align-items: center; justify-content: center; font-weight: bold;">
                {oee:.1f}%
            </div>
        </div>
    </div>
    <div class="kpi-grid">
        <div style="background: #f8f9fa; border-left: 5px solid #1E3A8A; padding: 10px; border-radius: 5px; text-align:center;">
            <p style="margin:0; font-size:12px;">EFICIENCIA REAL</p>
            <h2 style="margin:0;">{ef_real:.1f}%</h2>
        </div>
        <div style="background: #E8F5E9; border-left: 5px solid #2E7D32; padding: 10px; border-radius: 5px; text-align:center;">
            <p style="margin:0; font-size:12px;">EFICIENCIA PROD.</p>
            <h2 style="margin:0; color:#2E7D32;">{ef_prod:.1f}%</h2>
        </div>
        <div class="kpi-costo" style="background: #FFEBEE; border: 1px solid #C62828; padding: 15px; border-radius: 5px; text-align:center;">
            <p style="margin:0; font-size:16px; color:#C62828;">COSTO HH IMP.</p>
            <h2 style="margin:5px 0; color:#C62828; font-size:35px;">${t_costo:,.0f}</h2>
        </div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================================
# 7. GRÁFICOS (METRICAS 1 A 6)
# =========================================================================
st.markdown("---")
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("1. EFICIENCIA REAL POR MES")
    if not df_f.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 5))
        ag1 = df_f.groupby('Mes_Str').agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag1['%'] = (ag1['HH_STD'] / ag1['HH_Disp'] * 100)
        ag1['HH_Disp'].plot(kind='bar', ax=ax1, color='black', alpha=0.3, label='Disp')
        ag1['HH_STD'].plot(kind='bar', ax=ax1, color='#1E3A8A', label='STD')
        ax1.legend(); st.pyplot(fig1)

with col_b:
    st.subheader("2. EFICIENCIA PRODUCTIVA")
    if not df_f.empty:
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        ag2 = df_f.groupby('Mes_Str').agg({'HH_STD':'sum', 'HH_Prod_GAP':'sum'})
        ag2['%'] = (ag2['HH_STD'] / ag2['HH_Prod_GAP'] * 100)
        ax2.plot(ag2.index, ag2['%'], marker='s', color='green', linewidth=3)
        ax2.axhline(100, color='red', linestyle='--')
        st.pyplot(fig2)

# ... (Sección de Pareto y más gráficos) ...
st.markdown("---")
st.subheader("7. DETALLE DE IMPRODUCTIVIDAD Y ACCIONES SUGERIDAS")
if not df_im.empty:
    # Filtro de seguridad para la tabla
    df_im_f = df_im.copy()
    if s_pl: 
        c_pl_im = next((c for c in df_im.columns if 'PLANTA' in c), None)
        if c_pl_im: df_im_f = df_im_f[df_im_f[c_pl_im].isin(s_pl)]
    
    # Creamos la columna de acciones
    df_im_f['ACCIÓN SUGERIDA'] = df_im_f['TIPO_PARADA'].apply(generar_accion_sugerida)
    st.dataframe(df_im_f[['TIPO_PARADA', 'HH_IMPRODUCTIVAS', 'ACCIÓN SUGERIDA']].sort_values('HH_IMPRODUCTIVAS', ascending=False), use_container_width=True)

st.success("✅ Tablero OMBU v3.0 cargado con éxito. ¡OEEEEE!")
