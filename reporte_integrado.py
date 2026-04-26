import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe

# =========================================================================
# 1. CONFIGURACIÓN VISUAL
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide")
st.markdown("""
<style>
    header, [data-testid="stHeader"], [data-testid="stToolbar"], footer {display: none !important;}
    .block-container {padding-top: 1rem !important;}
    div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {
        position: -webkit-sticky; position: sticky; top: 0px;
        background-color: white; z-index: 99;
        padding: 10px; border-bottom: 2px solid #1E3A8A;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if not st.session_state['autenticado']:
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        st.markdown("<h2 style='text-align:center; color:#1E3A8A;'>GESTIÓN OMBÚ S.A.</h2>", unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['autenticado'] = True
                    st.rerun()
                else: st.error("❌ Credenciales incorrectas")
    st.stop()

# =========================================================================
# 3. CARGA DE DATOS (SUPER BLINDADA)
# =========================================================================
@st.cache_data(ttl=600)
def cargar_datos():
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    # Leemos archivos ignorando errores de motor
    d_ef = pd.read_excel(url_ef)
    d_im = pd.read_excel(url_im)
    
    def limpiar_columnas(df):
        # Limpieza extrema de nombres de columnas
        df.columns = [str(c).strip().upper() for c in df.columns]
        m = {}
        for c in df.columns:
            if 'FECHA' in c: m[c] = 'Fecha'
            elif 'PLANTA' in c: m[c] = 'Planta'
            elif 'LÍNEA' in c or 'LINEA' in c: m[c] = 'Linea'
            elif 'PUESTO' in c and 'ULTIMO' not in c and 'ÚLTIMO' not in c: m[c] = 'Puesto'
            elif 'CANT' in c and 'PROD' in c: m[c] = 'Cant_Prod'
            elif 'STD' in c: m[c] = 'HH_STD'
            elif 'DISP' in c: m[c] = 'HH_Disp'
            elif 'GAP' in c: m[c] = 'HH_Prod_GAP'
            elif 'COSTO' in c: m[c] = 'Costo'
        df = df.rename(columns=m)
        # Nos quedamos solo con las columnas que renombramos (evita duplicados ocultos)
        columnas_validas = [v for v in m.values()]
        return df.loc[:, df.columns.isin(columnas_validas)]

    d_ef = limpiar_columnas(d_ef)
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    
    # Aseguramos que las columnas numéricas sean números de verdad
    for col in ['HH_STD', 'HH_Disp', 'HH_Prod_GAP', 'Cant_Prod', 'Costo']:
        if col in d_ef.columns:
            d_ef[col] = pd.to_numeric(d_ef[col], errors='coerce').fillna(0)
    
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    return d_ef, d_im

try:
    df_ef, df_im = cargar_datos()
except Exception as e:
    st.error(f"Error cargando el Excel: {e}")
    st.stop()

# =========================================================================
# 4. FILTROS
# =========================================================================
st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
with st.sidebar:
    st.title("⚙️ PANEL DE CONTROL")
    meses = sorted(df_ef['Mes_Str'].dropna().unique())
    sel_mes = st.multiselect("📅 Seleccionar Mes", meses)
    plantas = sorted(df_ef['Planta'].dropna().unique())
    sel_planta = st.multiselect("🏭 Seleccionar Planta", plantas)

df_f = df_ef.copy()
if sel_mes: df_f = df_f[df_f['Mes_Str'].isin(sel_mes)]
if sel_planta: df_f = df_f[df_f['Planta'].isin(sel_planta)]

# =========================================================================
# 5. CÁLCULOS (ESTA PARTE MATA EL ERROR DE LA FOTO)
# =========================================================================
# Usamos .sum().sum() por si Pandas detectó columnas "mellizas" 
# y sumamos todo para que el resultado sea SIEMPRE un único número.
t_std = float(df_f['HH_STD'].values.sum()) if 'HH_STD' in df_f.columns else 0.0
t_disp = float(df_f['HH_Disp'].values.sum()) if 'HH_Disp' in df_f.columns else 0.0
t_prod = float(df_f['HH_Prod_GAP'].values.sum()) if 'HH_Prod_GAP' in df_f.columns else 0.0
t_costo = float(df_f['Costo'].values.sum()) if 'Costo' in df_f.columns else 0.0

# Eficiencias seguras
ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0.0
ef_prod = (t_std / t_prod * 100) if t_prod > 0 else 0.0
calidad_sim = 98.0
oee = ef_real * (calidad_sim / 100)

# =========================================================================
# 6. DISEÑO DEL TABLERO
# =========================================================================
st.title("🚀 TABLERO INTEGRADO CGU - OMBÚ")

# Carteles Principales
c1, c2, c3 = st.columns(3)
c1.metric("EFICIENCIA REAL", f"{ef_real:.1f}%")
c2.metric("EFICIENCIA PROD.", f"{ef_prod:.1f}%")
c3.metric("COSTO HH IMP.", f"${t_costo:,.0f}")

# Módulo OEE Destacado
color_oee = "#4CAF50" if oee > 80 else "#FFC107" if oee > 60 else "#F44336"
st.markdown(f"""
<div style="background: linear-gradient(135deg, #f0f4f8, #d9e2ec); border: 2px solid #1E3A8A; padding: 25px; border-radius: 15px; text-align: center; margin-top: 20px;">
    <h2 style="color: #1E3A8A; margin:0; font-size: 28px;">🏆 OEE SIMULADO: {oee:.1f}%</h2>
    <div style="background: #ccc; border-radius: 20px; height: 35px; margin-top:15px; border: 1px solid #999;">
        <div style="background: {color_oee}; width: {min(oee, 100):.1f}%; height: 100%; border-radius: 20px; transition: width 1s;"></div>
    </div>
    <p style="margin-top:10px; color: #555; font-weight: bold;">Cálculo: Efic. Real ({ef_real:.1f}%) &times; Calidad Simulada (98%)</p>
</div>
""", unsafe_allow_html=True)

# Gráfico de tendencia simple
st.subheader("📈 Evolución de HH STD por Mes")
if not df_f.empty:
    fig, ax = plt.subplots(figsize=(10, 4))
    df_f.groupby('Mes_Str')['HH_STD'].sum().plot(kind='bar', ax=ax, color='#1E3A8A', edgecolor='black')
    plt.xticks(rotation=0)
    st.pyplot(fig)

st.success("✅ Tablero funcionando. ¡Dale OEEEEE!")
