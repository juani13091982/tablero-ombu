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
# 1. CONFIGURACIÓN Y ESCUDO VISUAL
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
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold'})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if not st.session_state['autenticado']:
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<h2 style='text-align:center;'>GESTIÓN OMBÚ</h2>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar"):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['autenticado'] = True
                    st.rerun()
                else: st.error("Incorrecto")
    st.stop()

# =========================================================================
# 3. CARGA DE DATOS (CON SUPER ESCUDO)
# =========================================================================
try:
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    df_ef = pd.read_excel(url_ef)
    df_im = pd.read_excel(url_im)
    
    # NORMALIZADOR DE COLUMNAS (Mata el error de la foto)
    def normalizar(df):
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
            elif 'COSTO' in c: mapping[c] = 'Costo'
        return df.rename(columns=mapping)

    df_ef = normalizar(df_ef)
    df_im = normalizar(df_im)

    # ASEGURAR COLUMNAS CRÍTICAS (Si no existen, las crea vacías)
    for col in ['Fecha', 'Planta', 'Linea', 'Puesto_Trabajo', 'Es_Ultimo_Puesto', 'HH_STD', 'HH_Disp', 'HH_Prod_GAP', 'Cant_Prod', 'Costo']:
        if col not in df_ef.columns: df_ef[col] = 0 if col in ['HH_STD', 'HH_Disp', 'HH_Prod_GAP', 'Cant_Prod', 'Costo'] else "S/D"
    
    if 'HH_IMPRODUCTIVAS' not in df_im.columns:
        # Busca cualquier columna que diga HH e IMP
        col_imp = [c for c in df_im.columns if 'HH' in c and 'IMP' in c]
        if col_imp: df_im.rename(columns={col_imp[0]: 'HH_IMPRODUCTIVAS'}, inplace=True)
        else: df_im['HH_IMPRODUCTIVAS'] = 0

    # Limpieza de fechas
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce')
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    
except Exception as e:
    st.error(f"Error cargando datos: {e}")
    st.stop()

# =========================================================================
# 4. FILTROS Y LÓGICA
# =========================================================================
st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
with st.sidebar:
    st.title("FILTROS")
    sel_mes = st.multiselect("Mes", sorted(df_ef['Mes_Str'].dropna().unique()))
    sel_planta = st.multiselect("Planta", sorted(df_ef['Planta'].dropna().unique()))

df_f = df_ef.copy()
if sel_mes: df_f = df_f[df_f['Mes_Str'].isin(sel_mes)]
if sel_planta: df_f = df_f[df_f['Planta'].isin(sel_planta)]

# =========================================================================
# 5. EL TABLERO (OEE Y KPIs)
# =========================================================================
st.title("🚀 TABLERO INTEGRADO CGP")

# Cálculos
t_std = df_f['HH_STD'].sum()
t_disp = df_f['HH_Disp'].sum()
t_prod = df_f['HH_Prod_GAP'].sum()
t_costo = df_f['Costo'].sum()

ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0
ef_prod = (t_std / t_prod * 100) if t_prod > 0 else 0
calidad_sim = 98.0
oee = ef_real * (calidad_sim / 100)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    st.metric("EFICIENCIA REAL", f"{ef_real:.1f}%")
with c2:
    st.metric("EFICIENCIA PROD.", f"{ef_prod:.1f}%")
with c3:
    st.metric("COSTO TOTAL", f"${t_costo:,.0f}")

# --- MÓDULO OEE ---
st.markdown(f"""
<div style="background: #f0f4f8; border: 2px solid #1E3A8A; padding: 20px; border-radius: 15px; text-align: center;">
    <h2 style="color: #1E3A8A; margin:0;">🏆 OEE SIMULADO: {oee:.1f}%</h2>
    <div style="background: #ddd; border-radius: 20px; height: 30px; margin-top:10px;">
        <div style="background: {'#4CAF50' if oee > 80 else '#FFC107'}; width: {min(oee, 100)}%; height: 100%; border-radius: 20px;"></div>
    </div>
    <p style="margin-top:5px;">Rendimiento Real &times; Calidad Simulada (98%)</p>
</div>
""", unsafe_allow_html=True)

# --- GRÁFICO 1 ---
st.subheader("1. Evolución Eficiencia Real")
if not df_f.empty:
    fig, ax = plt.subplots(figsize=(10, 4))
    ag = df_f.groupby('Mes_Str')['HH_STD'].sum()
    ag.plot(kind='bar', ax=ax, color='#1E3A8A')
    st.pyplot(fig)

st.success("¡Tablero cargado con éxito! Dale OEEEEE!")
