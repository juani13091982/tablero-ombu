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
    
    /* CAJA PEGAJOSA (MANTIENE HEADER Y FILTROS FIJOS) */
    div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {
        position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
        background-color: rgba(14, 17, 23, 0.98) !important; z-index: 99999 !important;
        padding: 5px 10px 10px 10px !important; border-bottom: 2px solid #1E3A8A !important;
        box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
    }
    
    /* Etiquetas solo en PC */
    div[data-testid="stMultiSelect"] label p { font-size: 15px !important; font-weight: bold !important; color: #90CAF9 !important; margin-bottom: -5px !important; }

    .kpi-grid { display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 8px; }
    .kpi-ef-real { grid-column: 1; grid-row: 1; }
    .kpi-ef-prod { grid-column: 2; grid-row: 1; }
    .kpi-costo { grid-column: 3; grid-row: 1 / span 2; }
    .kpi-top-ef { grid-column: 1; grid-row: 2; }
    .kpi-top-imp { grid-column: 2; grid-row: 2; }

    h3 { white-space: nowrap !important; }
    .sub-title { font-size: 14px; color: #aaa; margin-top: -5px; margin-bottom: 15px; display: block; }

    /* ==================================================================== */
    /* VISTA EXCLUSIVA PARA CELULARES (PERFECCIÓN ABSOLUTA 2x2)             */
    /* ==================================================================== */
    @media (max-width: 1024px) {
        div[data-testid="stHorizontalBlock"]:has(form) { display: flex !important; justify-content: center !important; width: 100% !important; }
        div[data-testid="stHorizontalBlock"]:has(form) > div:not(:has(form)) { display: none !important; }
        div[data-testid="stHorizontalBlock"]:has(form) > div:has(form) { width: 100% !important; max-width: 450px !important; }

        div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {
            position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
            padding: 5px 5px 5px 5px !important;
            border-bottom: 2px solid #1E3A8A !important;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.5) !important;
            z-index: 99999 !important;
        }

        div[data-testid="stHorizontalBlock"]:has(#header-anchor) { display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 5px !important; }
        div[data-testid="stHorizontalBlock"]:has(#header-anchor) > div:nth-child(1) { display: none !important; }
        div[data-testid="stHorizontalBlock"]:has(#header-anchor) > div:nth-child(2) { width: 100% !important; display: flex !important; justify-content: center !important; }
        div[data-testid="stHorizontalBlock"]:has(#header-anchor) > div:nth-child(3) { display: none !important; }
        div[data-testid="stHorizontalBlock"]:has(#header-anchor) h3 { font-size: 18px !important; margin: 0px !important; text-align: center !important; width: 100% !important;}

        html body div[data-testid="stHorizontalBlock"]:has(#filtro-row) { display: flex !important; flex-direction: row !important; flex-wrap: wrap !important; width: 100% !important; gap: 4px 2% !important; margin-top: 0px !important; }
        html body div[data-testid="stHorizontalBlock"]:has(#filtro-row) > div[data-testid="column"] { width: 49% !important; min-width: 49% !important; max-width: 49% !important; flex: 1 1 49% !important; padding: 0px !important; }

        div[data-testid="stMultiSelect"] > label { display: none !important; height: 0px !important; margin: 0px !important; padding: 0px !important; visibility: hidden !important;}
        [data-testid="stMultiSelect"] { margin-top: -12px !important; margin-bottom: 0px !important; }
        
        .stMultiSelect div[data-baseweb="select"] { font-size: 13px !important; padding: 0 !important; min-height: 36px !important; height: 36px !important; border-radius: 6px !important; overflow: hidden !important;}
        .stMultiSelect div[data-baseweb="select"] span { font-size: 12px !important; padding: 0px 2px !important; }

        h2 { font-size: 18px !important; white-space: normal !important; margin-bottom: 5px !important; padding-bottom: 0px !important; line-height: 1.2 !important;}
        .sub-title { margin-top: 5px !important; margin-bottom: 15px !important; font-size: 13px !important; display: block !important; }
        hr { margin: 10px 0px !important; }

        .kpi-grid { display: flex !important; flex-direction: column !important; gap: 6px !important; }
        .kpi-grid h2 { font-size: 40px !important; line-height: 1.0 !important; white-space: normal !important; overflow: visible !important;}
        .kpi-grid h4 { font-size: 18px !important; margin-bottom: 0px !important; }
        .kpi-costo h2 { font-size: 42px !important; }

        .block-container > div > div > div > div[data-testid="stHorizontalBlock"]:nth-of-type(n+3) { display: flex !important; flex-direction: column !important; width: 100% !important; }
        .block-container > div > div > div > div[data-testid="stHorizontalBlock"]:nth-of-type(n+3) > div[data-testid="column"] { width: 100% !important; min-width: 100% !important; margin-bottom: 15px !important; }
    }
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b = [pe.withStroke(linewidth=3, foreground='white')]
efecto_n = [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o, caja_b = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
USUARIOS_PERMITIDOS = {"Métricas_Ombú": "CGP_2026"}
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True); c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        l1, l2, l3 = st.columns([1, 1, 1])
        with l2:
            try: st.image("LOGO OMBÚ.jpg", width=160)
            except: st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        st.markdown("<div style='text-align:center;'><h2 style='color:#1E3A8A;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</h2><p>Acceso Restringido - Control de Gestión</p></div>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in = st.text_input("Usuario Corporativo")
            p_in = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: 
                    st.session_state['autenticado'] = True; st.rerun()
                else: 
                    st.error("❌ Credenciales incorrectas.")

if not st.session_state['autenticado']: mostrar_login(); st.stop()

# =========================================================================
# 3. MOTOR DE DATOS Y FUNCIONES
# =========================================================================
def set_escala_y(ax, vmax, factor=1.6): ax.set_ylim(0, vmax * factor if vmax > 0 else 100)
def normalizar_lista(s_list): return [re.sub(r'[^A-Z0-9]', '', str(s).upper()) for s in s_list]
def dibujar_meses(ax, n_meses):
    for i in range(n_meses): ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def add_tendencia(ax, x, y):
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color='darkgray', linestyle=':', linewidth=4, path_effects=efecto_b, zorder=4, label='Tendencia')

def agregar_sello_agua(fig):
    try:
        img = mpimg.imread("LOGO OMBÚ.jpg"); ax_logo = fig.add_axes([0.88, 0.02, 0.08, 0.08], zorder=1)
        ax_logo.imshow(img, alpha=0.35); ax_logo.axis('off')
    except: pass

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if '✅' in d: return "🎯 ACCIÓN GLOBAL"
    if any(x in d for x in ['ROTURA', 'FALLA', 'CORTE', 'MANTENIMIENTO', 'MECANICO', 'ELECTRICO', 'ROT']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'ESPERA', 'PUENTE', 'GRUA', 'ABASTECIMIENTO', 'INSUMO']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'PLANO', 'ERROR', 'DEFECTO']): return "🔎 Ajustar Calidad"
    if any(x in d for x in ['LIMPIEZA', 'ORDEN', '5S']): return "🧹 Optimizar 5S"
    if any(x in d for x in ['REUNION', 'CHARLA', 'CAPACITACION', 'PERSONAL', 'AUSENCIA']): return "👥 Gestionar RRHH"
    if any(x in d for x in ['PREPARACION', 'SET UP', 'SETUP', 'CAMBIO']): return "⏱️ Reducir Set-Up"
    return "⚡ Investigar Causa"

@st.cache_data(ttl=300)
def cargar_datos():
    url_ef = "https://docs.google.com/spreadsheets/d/1_kO7GtjnlGHYgMnY7pkJGdRuFRAJB8xp8pGj8I0jq5E/export?format=xlsx"
    url_im = "https://docs.google.com/spreadsheets/d/1GwshvXAotIShBPcX69vlE8_juO6w5aAFBCaXEsgSdCY/export?format=xlsx"
    df_ef = pd.read_excel(url_ef)
    df_im = pd.read_excel(url_im)
    df_ef.columns = df_ef.columns.str.strip()
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    cols_numericas = ['HH_STD_TOTAL', 'HH_Disponibles', 'Cant._Prod._A1', 'HH_Productivas_C/GAP', 'Costo_Improd._$']
    c_std_u_potencial = next((c for c in df_ef.columns if 'STD' in str(c).upper() and ('UNID' in str(c).upper() or '/ U' in str(c).upper())), None)
    if c_std_u_potencial and c_std_u_potencial not in cols_numericas: cols_numericas.append(c_std_u_potencial)

    for col in cols_numericas:
        if col in df_ef.columns: df_ef[col] = pd.to_numeric(df_ef[col], errors='coerce').fillna(0)
    
    if 'TIPO_PARADA' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c), df_im.columns[0]): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), df_im.columns[0]): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'DETALLE' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c), df_im.columns[0]): 'DETALLE'}, inplace=True)
        
    df_im['HH_IMPRODUCTIVAS'] = pd.to_numeric(df_im['HH_IMPRODUCTIVAS'], errors='coerce').fillna(0).abs()
    c_pu_det = next((c for c in df_im.columns if 'PUESTO' in c), None)
    if not c_pu_det: c_pu_det = 'PUESTO_X'
    if c_pu_det not in df_im.columns: df_im[c_pu_det] = "S/D"
    
    c_nom = next((c for c in df_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in df_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '') + ' ' + df_im[c_ape].astype(str).replace('nan', '')
    elif c_nom: df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '')
    else: df_im['OPERARIO'] = "S/D"
    df_im['OPERARIO'] = df_im['OPERARIO'].str.strip().replace('', 'S/D')

    c_fec = next((c for c in df_im.columns if 'A3' in str(c).upper() or 'INICIO' in str(c).upper() or 'FECHA' in str(c).upper()), None)
    df_im['FECHA_EXACTA'] = pd.to_datetime(df_im[c_fec], errors='coerce', dayfirst=True) if c_fec else pd.NaT
    if 'FECHA' in df_im.columns: df_im['FECHA'] = pd.to_datetime(df_im['FECHA'], errors='coerce', dayfirst=True).dt.to_period('M').dt.to_timestamp()
    else: df_im['FECHA'] = df_im['FECHA_EXACTA'].dt.to_period('M').dt.to_timestamp()
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce', dayfirst=True).dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_im['MES_STR'] = df_im['FECHA'].dt.strftime('%b-%Y') 

    def norm_s(s): return s.astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    if 'Planta' in df_ef.columns: df_ef['NORM_PLANTA'] = norm_s(df_ef['Planta'])
    if 'Linea' in df_ef.columns: df_ef['NORM_LINEA'] = norm_s(df_ef['Linea'])
    if 'Puesto_Trabajo' in df_ef.columns: df_ef['NORM_PUESTO'] = norm_s(df_ef['Puesto_Trabajo'])

    c_pl_im = next((c for c in df_im.columns if 'PLANTA' in str(c).upper()), None)
    c_li_im = next((c for c in df_im.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), None)
    c_pu_im = next((c for c in df_im.columns if 'PUESTO' in str(c).upper()), None)
    if c_pl_im: df_im['NORM_PLANTA'] = norm_s(df_im[c_pl_im])
    if c_li_im: df_im['NORM_LINEA'] = norm_s(df_im[c_li_im])
    if c_pu_im: df_im['NORM_PUESTO'] = norm_s(df_im[c_pu_im])

    return df_ef, df_im, c_pl_im, c_li_im, c_pu_im

try:
    df_ef, df_im, orig_col_pl, orig_col_li, orig_col_pu = cargar_datos()
except Exception as e: 
    st.error(f"Error crítico cargando datos: {e}"); st.stop()

# =========================================================================
# 4. FILTROS FIJOS (STICKY EN PC Y CELULAR)
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h_l, h_t, h_s = st.columns([0.8, 3.5, 0.7])
    with h_l:
        st.markdown("<span id='header-anchor'></span>", unsafe_allow_html=True)
        try: st.image("LOGO OMBÚ.jpg", width=90)
        except: st.markdown("##### OMBÚ")
    with h_t: st.markdown("<h3 style='text-align:center;'>TABLERO INTEGRADO C.G.P.</h3>", unsafe_allow_html=True)
    with h_s: 
        if st.button("🚪 Salir", use_container_width=True): st.session_state['autenticado'] = False; st.rerun()

    st.markdown("<span id='filtro-row'></span>", unsafe_allow_html=True)
    f_mes, f_pl, f_li, f_pu = st.columns(4)
    
    meses_disp = sorted(list(set(df_ef['Mes_Str'].dropna().unique()) | set(df_im['MES_STR'].dropna().unique())))
    with f_mes: 
        s_mes = st.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + meses_disp, placeholder="📅 Mes...")
        
    df_base_ef, df_base_im = df_ef.copy(), df_im.copy()
    if s_mes and "🎯 Acumulado YTD" not in s_mes:
        df_base_ef = df_base_ef[df_base_ef['Mes_Str'].isin(s_mes)]
        df_base_im = df_base_im[df_base_im['MES_STR'].isin(s_mes)]
        
    pl_ef = set(df_base_ef['Planta'].dropna().astype(str).unique())
    pl_im = set(df_base_im[orig_col_pl].dropna().astype(str).unique()) if orig_col_pl and not df_base_im.empty else set()
    
    with f_pl: 
        s_pl = st.multiselect("🏭 Planta", sorted(list(pl_ef | pl_im)), placeholder="🏭 Planta...")
        
    if s_pl:
        norm_pl = normalizar_lista(s_pl)
        df_base_ef = df_base_ef[df_base_ef['Planta'].isin(s_pl)]
        if 'NORM_PLANTA' in df_base_im.columns and not df_base_im.empty: df_base_im = df_base_im[df_base_im['NORM_PLANTA'].isin(norm_pl)]
            
    li_ef = set(df_base_ef['Linea'].dropna().astype(str).unique())
    li_im = set(df_base_im[orig_col_li].dropna().astype(str).unique()) if orig_col_li and not df_base_im.empty else set()
    
    with f_li: 
        s_li = st.multiselect("⚙️ Línea", sorted(list(li_ef | li_im)), placeholder="⚙️ Línea...")
        
    if s_li:
        norm_li = normalizar_lista(s_li)
        df_base_ef = df_base_ef[df_base_ef['Linea'].isin(s_li)]
        if 'NORM_LINEA' in df_base_im.columns and not df_base_im.empty: df_base_im = df_base_im[df_base_im['NORM_LINEA'].isin(norm_li)]
            
    pu_ef = set(df_base_ef['Puesto_Trabajo'].dropna().astype(str).unique())
    pu_im = set(df_base_im[orig_col_pu].dropna().astype(str).unique()) if orig_col_pu and not df_base_im.empty else set()
    
    with f_pu: 
        s_pu = st.multiselect("🛠️ Puesto", sorted(list(pu_ef | pu_im)), placeholder="🛠️ Puesto...")

# APLICACIÓN DE FILTROS A DF FINALES
df_ef_f, df_im_f = df_ef.copy(), df_im.copy()
if s_pl: 
    norm_pl = normalizar_lista(s_pl)
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(s_pl)]
    if 'NORM_PLANTA' in df_im_f.columns: df_im_f = df_im_f[df_im_f['NORM_PLANTA'].isin(norm_pl)]
if s_li: 
    norm_li = normalizar_lista(s_li)
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(s_li)]
    if 'NORM_LINEA' in df_im_f.columns: df_im_f = df_im_f[df_im_f['NORM_LINEA'].isin(norm_li)]
if s_pu: 
    norm_pu = normalizar_lista(s_pu)
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(s_pu)]
    if 'NORM_PUESTO' in df_im_f.columns: df_im_f = df_im_f[df_im_f['NORM_PUESTO'].isin(norm_pu)]
if s_mes and "🎯 Acumulado YTD" not in s_mes: 
    df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(s_mes)]
    if 'MES_STR' in df_im_f.columns: df_im_f = df_im_f[df_im_f['MES_STR'].isin(s_mes)]

# REGLA DE ORO: df_plot_1 SIRVE EXCLUSIVAMENTE PARA EFICIENCIA DE LÍNEA
warn_linea = False
if s_pu: df_plot_1 = df_ef_f.copy()
else: 
    df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_salida.empty: df_plot_1 = df_salida
    else: df_plot_1 = df_ef_f.copy(); warn_linea = True if s_li else False

# BLINDAJE DE COLUMNA DE HH PRODUCTIVAS TOTALES
col_prod_tot = 'HH_Productivas_C/GAP'
for c in df_ef_f.columns:
    c_upper = str(c).upper()
    if 'PROD' in c_upper and 'GAP' in c_upper and 'UNID' not in c_upper and '/ U' not in c_upper:
        col_prod_tot = c
        break

# CÁLCULOS PONDERADOS UNIVERSALES PARA CARTELES
tot_costo = df_ef_f['Costo_Improd._$'].sum() if not df_ef_f.empty else 0
tot_hh_imp = df_im_f['HH_IMPRODUCTIVAS'].sum() if not df_im_f.empty else 0

tot_std = df_plot_1['HH_STD_TOTAL'].sum() if not df_plot_1.empty else 0
tot_disp_eficiencia = df_plot_1['HH_Disponibles'].sum() if not df_plot_1.empty else 0
tot_prod = df_plot_1[col_prod_tot].sum() if (col_prod_tot in df_plot_1.columns and not df_plot_1.empty) else 0

kpi_ef_real = (tot_std / tot_disp_eficiencia * 100) if tot_disp_eficiencia > 0 else 0
kpi_ef_prod = (tot_std / tot_prod * 100) if tot_prod > 0 else 0

top3_m1_html = "<div style='font-size:14px; color:#aaa; text-align:center;'>S/D</div>"
if not df_ef_f.empty:
    ag_p = df_ef_f.groupby('Puesto_Trabajo').agg({'HH_STD_TOTAL':'sum', 'HH_Disponibles':'sum'})
    ag_p['Ef'] = (ag_p['HH_STD_TOTAL'] / ag_p['HH_Disponibles'] * 100).fillna(0)
    top3_val = ag_p['Ef'].nlargest(3)
    if not top3_val.empty:
        filas = [f"<div style='display:flex; justify-content:space-between; margin-top:4px; font-size:13px;'><span style='white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:140px;' title='{p}'>{i}. {p}</span><strong style='color:#90CAF9; font-size:14px;'>{v:.1f}%</strong></div>" for i, (p, v) in enumerate(top3_val.items(), 1)]
        top3_m1_html = "".join(filas)

top3_imp_html = "<div style='font-size:14px; color:#aaa; text-align:center;'>S/D</div>"
if not df_im_f.empty:
    c_pu_im_top = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2] if len(df_im_f.columns)>2 else None)
    if c_pu_im_top:
        ag_imp_p = df_im_f.groupby(c_pu_im_top)['HH_IMPRODUCTIVAS'].sum().nlargest(3)
        if not ag_imp_p.empty:
            filas_imp = [f"<div style='display:flex; justify-content:space-between; margin-top:4px; font-size:13px;'><span style='white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:140px;' title='{p}'>{i}. {p}</span><strong style='color:#FFCDD2; font-size:14px;'>{v:.1f}</strong></div>" for i, (p, v) in enumerate(ag_imp_p.items(), 1)]
            top3_imp_html = "".join(filas_imp)

# =========================================================================
# CARTELES KPI 
# =========================================================================
st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-ef-real" style="background: linear-gradient(135deg, #e0e0e0, #f5f5f5); border: 1px solid #aaa; border-left: 6px solid #1E3A8A; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); padding: 10px;">
        <h4 style="color: #1E3A8A;">EFICIENCIA REAL</h4>
        <h2 style="color: #111;">{kpi_ef_real:.1f}%</h2>
    </div>
    <div class="kpi-ef-prod" style="background: linear-gradient(135deg, #2E7D32, #4CAF50); border: 1px solid #1B5E20; border-left: 6px solid #A5D6A7; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); padding: 10px;">
        <h4 style="color: white;">EFICIENCIA PROD.</h4>
        <h2 style="color: white;">{kpi_ef_prod:.1f}%</h2>
    </div>
    <div class="kpi-costo" style="background: linear-gradient(135deg, #D32F2F, #E53935); border: 1px solid #B71C1C; border-radius: 8px; display: flex; flex-direction: column; justify-content: center; text-align:center; box-shadow: 2px 4px 15px rgba(211,47,47,0.4); padding: 10px;">
        <h4 style="color: white;">COSTO HH IMPROD.</h4>
        <p style="color: #FFCDD2; margin: 0; font-size: 14px;">(Oportunidad Perdida)</p>
        <h2 style="color: #FFEB3B; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${tot_costo:,.0f}</h2>
        <h4 style="color: white;">{tot_hh_imp:,.1f} <span style="font-size:16px; font-weight:normal;">HH</span></h4>
    </div>
    <div class="kpi-top-ef" style="background: #0D47A1; color: white; border-radius: 6px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); padding: 10px;">
        <h4 style="font-size:14px; color:#BBDEFB; text-align:center; border-bottom: 1px solid rgba(255,255,255,0.2); margin:0 0 5px 0; padding-bottom:5px;">🏆 TOP EF. REAL</h4>
        {top3_m1_html}
    </div>
    <div class="kpi-top-imp" style="background: #B71C1C; color: white; border-radius: 6px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); padding: 10px;">
        <h4 style="font-size:14px; color:#FFCDD2; text-align:center; border-bottom: 1px solid rgba(255,255,255,0.2); margin:0 0 5px 0; padding-bottom:5px;">⚠️ TOP HH IMP.</h4>
        {top3_imp_html}
    </div>
</div>
""", unsafe_allow_html=True)

t_enc = f"Filtros >> Planta: {'+'.join(s_pl) if s_pl else 'Todas'} | Línea: {'+'.join(s_li) if s_li else 'Todas'} | Puesto: {'+'.join(s_pu) if s_pu else 'Todos'}"

# =========================================================================
# 5. BALANCE GERENCIAL PARA EL DUEÑO (MÉTRICA 7 ADAPTADA)
# =========================================================================
df_ef_h_res, df_im_h_res = df_ef.copy(), df_im.copy()
if s_pl: 
    df_ef_h_res = df_ef_h_res[df_ef_h_res['Planta'].isin(s_pl)]
    if 'NORM_PLANTA' in df_im_h_res.columns: df_im_h_res = df_im_h_res[df_im_h_res['NORM_PLANTA'].isin(norm_pl)]
if s_li: 
    df_ef_h_res = df_ef_h_res[df_ef_h_res['Linea'].isin(s_li)]
    if 'NORM_LINEA' in df_im_h_res.columns: df_im_h_res = df_im_h_res[df_im_h_res['NORM_LINEA'].isin(norm_li)]
if s_pu: 
    df_ef_h_res = df_ef_h_res[df_ef_h_res['Puesto_Trabajo'].isin(s_pu)]
    if 'NORM_PUESTO' in df_im_h_res.columns: df_im_h_res = df_im_h_res[df_im_h_res['NORM_PUESTO'].isin(norm_pu)]

fechas_ordenadas_res = sorted(df_ef_h_res['Fecha'].dropna().unique())
meses_seleccionados_res = [m for m in s_mes if m != "🎯 Acumulado YTD"] if s_mes else []
if meses_seleccionados_res: max_fecha_res = pd.to_datetime(meses_seleccionados_res, format="%b-%Y").max()
else: max_fecha_res = fechas_ordenadas_res[-1] if len(fechas_ordenadas_res) > 0 else pd.Timestamp.now()

fechas_previas_res = [f for f in fechas_ordenadas_res if f < max_fecha_res][-3:]

inc_hist_res = 0.0
if len(fechas_previas_res) > 0:
    ef_prev_res = df_ef_h_res[df_ef_h_res['Fecha'].isin(fechas_previas_res)]
    im_prev_res = df_im_h_res[df_im_h_res['FECHA'].isin(fechas_previas_res)]
    disp_prev_res = ef_prev_res['HH_Disponibles'].sum()
    imp_prev_res = im_prev_res['HH_IMPRODUCTIVAS'].sum()
    if disp_prev_res > 0: inc_hist_res = imp_prev_res / disp_prev_res

tot_disp_todas_res = df_ef_f['HH_Disponibles'].sum() if not df_ef_f.empty else 0
inc_act_res = (tot_hh_imp / tot_disp_todas_res) if tot_disp_todas_res > 0 else 0

ahorro_usd_res = 0
ahorro_hh = 0
if tot_disp_todas_res > 0 and len(fechas_previas_res) > 0:
    ahorro_hh = (inc_hist_res - inc_act_res) * tot_disp_todas_res
    costo_hh_res = (tot_costo / tot_hh_imp) if tot_hh_imp > 0 else 15000
    ahorro_usd_res = ahorro_hh * costo_hh_res

ag8_res = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL':'sum', col_prod_tot:'sum', 'Cant._Prod._A1':'sum'}).reset_index()
ag8_res = ag8_res[ag8_res['Cant._Prod._A1'] > 0]
tot_gp_crv26 = 0
if not ag8_res.empty:
    ag8_res['Horas_Ritmo'] = (ag8_res['HH_STD_TOTAL'] - ag8_res[col_prod_tot])
    tot_gp_crv26 = ag8_res['Horas_Ritmo'].sum() / 130.0

res_color_econ = "#1B5E20" if ahorro_usd_res >= 0 else "#B71C1C"
res_color_prod = "#1B5E20" if tot_gp_crv26 >= 0 else "#B71C1C"

st.markdown(f"""
<div style="background: #111; padding: 20px; border-radius: 10px; border: 2px solid #555; margin-top: 15px; margin-bottom: 20px; box-shadow: 0px 10px 20px rgba(0,0,0,0.5);">
    <h3 style="color: white; text-align: center; margin-top: 0; margin-bottom: 15px; font-size: 20px; white-space: normal; line-height: 1.3;">📊 BALANCE GERENCIAL<br><span style="font-size:14px; font-weight:normal; color:#ccc;">(EQUIVALENCIA MÁQUINAS CRV 26)</span></h3>
    <div style="display: flex; justify-content: space-around; flex-wrap: wrap; gap: 15px;">
        <div style="text-align:center; min-width: 250px; flex: 1;">
            <p style="color:#aaa; margin:0; font-size: 14px; font-weight: bold;">IMPACTO ECONÓMICO VS HISTORIA</p>
            <h2 style="color:{res_color_econ}; font-size:42px; margin:0; padding: 5px 0;">{"+" if ahorro_usd_res >= 0 else "-"}${abs(ahorro_usd_res):,.0f}</h2>
            <h4 style="color:{res_color_econ}; margin:0; font-size:18px;">{abs(ahorro_hh):,.0f} HH {'Recuperadas' if ahorro_hh >= 0 else 'Perdidas'}</h4>
            <p style="color:{res_color_econ}; font-size:13px; margin:0; padding-top:5px;">(Por gestión de ineficiencias)</p>
        </div>
        <div style="text-align:center; min-width: 250px; flex: 1; border-left: 1px dashed #444;">
            <p style="color:#aaa; margin:0; font-size: 14px; font-weight: bold;">CAPACIDAD DE PRODUCCIÓN EXTRA</p>
            <h2 style="color:{res_color_prod}; font-size:42px; margin:0; padding: 5px 0;">{abs(tot_gp_crv26):.1f} Máq.</h2>
            <p style="color:{res_color_prod}; font-size:13px; margin:0; padding-top:28px;">({'Ganadas' if tot_gp_crv26 >= 0 else 'Perdidas'} por ritmo de trabajo)</p>
        </div>
    </div>
    <p style="color: #666; font-size: 11px; text-align: center; margin-top: 10px;">* Cálculos de capacidad basados en estándar de 130 HH/Máquina CRV 26</p>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 6. GRÁFICOS MÉTRICAS 1 Y 2
# =========================================================================
col1, col2 = st.columns(2)

with col1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div class='sub-title'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    if warn_linea: st.warning("⚠️ Esta Línea NO registra un 'Último Puesto'. Seleccione un Puesto para análisis preciso.")
    
    if not df_plot_1.empty:
        ag1 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'}).reset_index()
        ag1['Ef_Real'] = (ag1['HH_STD_TOTAL'] / ag1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig1, ax1 = plt.subplots(figsize=(14, 10)); ax1_line = ax1.twinx()
        fig1.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag1))
        bs = ax1.bar(x_idx - 0.17, ag1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bd = ax1.bar(x_idx + 0.17, ag1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_escala_y(ax1, ag1['HH_Disponibles'].max(), 1.6)
        ax1.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax1.bar_label(bd, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax1, len(x_idx))
        
        for i, bar in enumerate(bs):
            val_prod = ag1['Cant._Prod._A1'].iloc[i]
            vu = int(float(val_prod)) if pd.notna(val_prod) else 0 
            if vu > 0: ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{vu} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)
        
        ax1_line.plot(x_idx, ag1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real', zorder=5)
        add_tendencia(ax1_line, x_idx, ag1['Ef_Real'])
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_idx[-1] if len(x_idx)>0 else 0, 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3)); ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, val in enumerate(ag1['Ef_Real']): ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
            
        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig1); st.pyplot(fig1, use_container_width=True)
    else: st.warning("⚠️ Sin datos para Eficiencia Real.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div class='sub-title'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    if not df_plot_1.empty:
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', col_prod_tot: 'sum', 'Cant._Prod._A1': 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2[col_prod_tot]).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2 = plt.subplots(figsize=(14, 10)); ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag2))
        bs = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bp = ax2.bar(x_idx + 0.17, ag2[col_prod_tot], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2[col_prod_tot].max()), 1.6)
        ax2.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax2.bar_label(bp, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax2, len(x_idx))
        
        for i, bar in enumerate(bs):
            val_prod = ag2['Cant._Prod._A1'].iloc[i]
            vu = int(float(val_prod)) if pd.notna(val_prod) else 0 
            if vu > 0: ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{vu} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)
        
        ax2_line.plot(x_idx, ag2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Prod.', zorder=5)
        add_tendencia(ax2_line, x_idx, ag2['Ef_Prod'])
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_idx[-1] if len(x_idx)>0 else 0, 101, 'META = 100%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        ax2_line.set_ylim(0, max(110, ag2['Ef_Prod'].max()*1.3)); ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, val in enumerate(ag2['Ef_Prod']): ax2_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
            
        ax2.set_xticks(x_idx); ax2.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
        ax2_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig2); st.pyplot(fig2, use_container_width=True)
    else: st.warning("⚠️ Sin datos para Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 7. GRÁFICOS MÉTRICAS 3 Y 4
# =========================================================================
col3, col4 = st.columns(2)
with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div class='sub-title'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag3 = df_ef_f.groupby('Fecha').agg({col_prod_tot: 'sum', 'HH_Disponibles': 'sum'}).reset_index()
        
        if not df_im_f.empty:
            ag_im = df_im_f.groupby('FECHA')['HH_IMPRODUCTIVAS'].sum().reset_index().rename(columns={'FECHA':'Fecha', 'HH_IMPRODUCTIVAS':'HH_Imp'})
            ag3 = pd.merge(ag3, ag_im, on='Fecha', how='left').fillna(0)
        else: ag3['HH_Imp'] = 0
            
        ag3['Total_Decl'] = ag3[col_prod_tot] + ag3['HH_Imp']
        
        fig3, ax3 = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.82, bottom=0.15, left=0.06, right=0.94)
        fig3.suptitle(t_enc, x=0.06, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag3))
        bp = ax3.bar(x_idx, ag3[col_prod_tot], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        bi = ax3.bar(x_idx, ag3['HH_Imp'], bottom=ag3[col_prod_tot], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax3.bar_label(bp, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.bar_label(bi, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        
        ax3.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='HH DISPONIBLES', zorder=5)
        set_escala_y(ax3, ag3['HH_Disponibles'].max(), 1.6); dibujar_meses(ax3, len(x_idx))

        tot_gap = ag3['HH_Disponibles'].sum() - ag3['Total_Decl'].sum()
        ax3.text(0.01, 0.98, f" GAP ACUMULADO: {int(tot_gap):,} HH ", transform=ax3.transAxes, fontsize=14, color='firebrick', bbox=dict(boxstyle="round,pad=0.4", fc="#FFEBEE", ec="firebrick", lw=2), fontweight='bold', zorder=20, va='top', ha='left')

        for i in range(len(x_idx)):
            hd, ht = ag3['HH_Disponibles'].iloc[i], ag3['Total_Decl'].iloc[i]; gap = hd - ht
            ax3.plot([i, i], [ht, hd], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            ax3.annotate(f"GAP:\n{int(gap)}", (i, ht + (gap/2)), color='firebrick', bbox=caja_b, ha='center', va='center', fontweight='bold', zorder=10)
            ax3.annotate(f"{int(hd)}", (i, hd + (ax3.get_ylim()[1]*0.05)), color='black', bbox=caja_b, ha='center', fontweight='bold', zorder=10)

        ax3.set_xticks(x_idx); ax3.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=3, frameon=True)
        agregar_sello_agua(fig3); st.pyplot(fig3, use_container_width=True)
    else: st.warning("⚠️ Sin datos para GAP.")

with col4:
    st.header("4. TENDENCIA COSTOS IMPRODUCTIVOS")
    st.markdown("<div class='sub-title'><i>Evolución de valorización económica de la ineficiencia</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({'Costo_Improd._$': 'sum'}).reset_index()
        if not df_im_f.empty:
            ag_im = df_im_f.groupby('FECHA')['HH_IMPRODUCTIVAS'].sum().reset_index().rename(columns={'FECHA':'Fecha', 'HH_IMPRODUCTIVAS':'HH_Imp'})
            ag4 = pd.merge(ag4, ag_im, on='Fecha', how='left').fillna(0)
        else: ag4['HH_Imp'] = 0

        fig4, ax4 = plt.subplots(figsize=(14, 10)); ax4_line = ax4.twinx()
        fig4.subplots_adjust(top=0.82, bottom=0.15, left=0.06, right=0.94)
        fig4.suptitle(t_enc, x=0.06, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag4))
        bi = ax4.bar(x_idx, ag4['HH_Imp'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4.bar_label(bi, padding=4, color='black', fontweight='bold', path_effects=efecto_b, zorder=4)
        set_escala_y(ax4, ag4['HH_Imp'].max(), 1.6) 
        
        ax4_line.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, label='COSTO ARS', zorder=5)
        add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'])
        ax4_line.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3)); ax4_line.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_line.get_yticks()], fontweight='bold')

        tot_imp, tot_cost = ag4['HH_Imp'].sum(), ag4['Costo_Improd._$'].sum()
        ax4.text(0.01, 0.98, f" HH IMPROD. ACUM.: {tot_imp:,.1f} HH  |  COSTO ACUM.: ${tot_cost:,.0f} ", transform=ax4.transAxes, fontsize=14, color='maroon', bbox=dict(boxstyle="round,pad=0.4", fc="#FFEBEE", ec="maroon", lw=2), fontweight='bold', zorder=20, va='top', ha='left')

        for i, val in enumerate(ag4['Costo_Improd._$']): ax4_line.annotate(f"${val:,.0f}", (x_idx[i], val + (ax4_line.get_ylim()[1]*0.04)), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax4.set_xticks(x_idx); ax4.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True); ax4_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig4); st.pyplot(fig4, use_container_width=True)
    else: st.warning("⚠️ No hay datos económicos.")

st.markdown("---")

# =========================================================================
# 8. GRÁFICOS MÉTRICAS 5 Y 6
# =========================================================================
col5, col6 = st.columns(2)
with col5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div class='sub-title'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)
    if not df_im_f.empty:
        ag5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        div = df_im_f['FECHA'].nunique() if df_im_f['FECHA'].nunique() > 0 else 1
        ag5['Prom_M'] = ag5['HH_IMPRODUCTIVAS'] / div; ag5 = ag5.sort_values(by='Prom_M', ascending=False)
        ag5['Pct_Acu'] = (ag5['Prom_M'].cumsum() / ag5['Prom_M'].sum()) * 100
        
        fig5, ax5 = plt.subplots(figsize=(14, 10)); ax5_line = ax5.twinx()
        fig5.subplots_adjust(top=0.80, bottom=0.35, left=0.06, right=0.92)
        fig5.suptitle(t_enc, x=0.06, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag5))
        bp = ax5.bar(x_idx, ag5['Prom_M'], color='maroon', edgecolor='white', zorder=2)
        
        set_escala_y(ax5, ag5['Prom_M'].max(), 3.5)
        ax5.bar_label(bp, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        ax5_line.plot(x_idx, ag5['Pct_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax5_line.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        ax5_line.set_ylim(0, 110); ax5_line.yaxis.set_major_formatter(mtick.PercentFormatter()) 
        
        lbls = [textwrap.fill(str(t), 20) for t in ag5['TIPO_PARADA']]
        ax5.set_xticks(x_idx); ax5.set_xticklabels(lbls, rotation=45, ha='right', fontsize=11, fontweight='bold')
        
        for i, val in enumerate(ag5['Pct_Acu']): ax5_line.annotate(f"{val:.1f}%", (x_idx[i], val + 6), color='white', bbox=caja_g, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
        
        tot_imp_acum = df_im_f['HH_IMPRODUCTIVAS'].sum()
        ax5.text(0.02, 0.95, f"HH IMP. ACUMULADAS: {tot_imp_acum:,.1f}", transform=ax5.transAxes, ha='left', va='top', bbox=dict(boxstyle="round,pad=0.5", fc="#f5f5f5", ec="gray", lw=1), fontsize=13, fontweight='bold', zorder=10)
        
        agrup_col = orig_col_pu if orig_col_pu else 'OPERARIO'
        if s_pu: agrup_col = 'OPERARIO'
        
        top3_causas = ag5['TIPO_PARADA'].head(3).tolist()
        for i, causa in enumerate(top3_causas):
            df_c = df_im_f[df_im_f['TIPO_PARADA'] == causa]
            if not df_c.empty and agrup_col in df_c.columns:
                top_name = df_c.groupby(agrup_col)['HH_IMPRODUCTIVAS'].sum().idxmax()
                top_val = df_c.groupby(agrup_col)['HH_IMPRODUCTIVAS'].sum().max()
                y_pos = ag5['Prom_M'].iloc[i] + (ag5['Prom_M'].max() * 0.1)
                ax5.text(i, y_pos, f"🚨 {top_name}\n({top_val:.1f}h)", rotation=90, ha='center', va='bottom', color='darkred', fontsize=11, fontweight='bold', bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.8), zorder=15)
                
        agregar_sello_agua(fig5); st.pyplot(fig5, use_container_width=True)
    else: st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div class='sub-title'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        df_ef_f['K_Mes'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp = df_ef_f.groupby('K_Mes', as_index=False)['HH_Disponibles'].sum()
        
        if not df_im_f.empty:
            df_im_f['K_Mes'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K_Mes', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6 = pd.merge(ag_disp, piv, on='K_Mes', how='left').fillna(0)
            list_c = [c for c in df6.columns if c not in ['HH_Disponibles', 'K_Mes']]
        else: df6, list_c = ag_disp.copy(), []
            
        df6['Suma_I'] = df6[list_c].sum(axis=1) if list_c else 0
        df6['Inc_%'] = (df6['Suma_I'] / df6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6['Fecha_O'] = pd.to_datetime(df6['K_Mes'] + '-01'); df6 = df6.sort_values(by='Fecha_O')
        
        fig6, ax6 = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.60, bottom=0.15, left=0.06, right=0.94) 
        fig6.suptitle(t_enc, x=0.06, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(df6))
        if list_c:
            base_st = np.zeros(len(df6)); paleta = plt.cm.tab20.colors
            for i, c_nom in enumerate(list_c):
                vals = df6[c_nom].values
                bar_stack = ax6.bar(x_idx, vals, bottom=base_st, label=textwrap.fill(c_nom, 15), color=paleta[i % 20], edgecolor='white', zorder=2)
                lbls_stk = [f"{int(v)}" if v > 0 else "" for v in vals]
                ax6.bar_label(bar_stack, labels=lbls_stk, label_type='center', color='white', fontsize=9, fontweight='bold', path_effects=efecto_n)
                base_st += vals
            ax6.legend(loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=4, framealpha=0.9, fontsize=12)
        else: ax6.bar(x_idx, np.zeros(len(df6)), color='white')
            
        set_escala_y(ax6, df6['Suma_I'].max(), 3.0) 
        
        ax6_line = ax6.twinx(); ax6_line.plot(x_idx, df6['Inc_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, label='% Incidencia', zorder=5)
        add_tendencia(ax6_line, x_idx, df6['Inc_%'])
        ax6_line.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_line.text(x_idx[-1] if len(x_idx)>0 else 0, 16, 'META = 15%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        
        for i, val in enumerate(df6['Inc_%']): 
            if df6['Suma_I'].iloc[i] > 0: 
                ax6_line.annotate(f"{val:.1f}%", (x_idx[i], val), textcoords="offset points", xytext=(0, 20), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_b, zorder=10)
        
        altura_fija_etiquetas = df6['Suma_I'].max() * 1.3
        for i in range(len(x_idx)):
            v_i, v_d = df6['Suma_I'].iloc[i], df6['HH_Disponibles'].iloc[i]
            if v_i > 0: ax6.annotate(f"Imp: {int(v_i)}\nDisp: {int(v_d)}", (i, altura_fija_etiquetas), ha='center', bbox=caja_o, fontweight='bold', zorder=10)
                
        ax6.set_xticks(x_idx); ax6.set_xticklabels(df6['K_Mes'], fontsize=14, fontweight='bold')
        ax6_line.set_ylim(0, max(30, df6['Inc_%'].max() * 1.5))
        agregar_sello_agua(fig6); st.pyplot(fig6, use_container_width=True)
    else: st.warning("⚠️ Sin datos históricos de eficiencia para cruzar.")

st.markdown("---")

col7, col8 = st.columns(2)

with col7:
    st.header("8. ESTABILIDAD DEL PROCESO")
    
    # INTELIGENCIA CUELLO DE BOTELLA: Solo se activa si eligieron una Planta y NINGUNA línea ni puesto.
    if not s_pl and not s_li and not s_pu:
        st.markdown("<div style='height: 32px;'></div>", unsafe_allow_html=True) # Espaciador para alinear con M9
        st.markdown("<div class='sub-title'><i>Desviación Tiempos Reales vs Estándar por Unidad</i></div>", unsafe_allow_html=True)
        st.info("🔒 Seleccione una **Planta**, **Línea** o **Puesto** en los Filtros Maestros para desbloquear el Análisis.")
    
    elif s_pl and not s_li and not s_pu:
        st.markdown("<div style='height: 32px;'></div>", unsafe_allow_html=True) # Espaciador para alinear con M9
        st.markdown("<div class='sub-title'><i>Análisis de Flujo y Cuello de Botella por Línea</i></div>", unsafe_allow_html=True)
        
        # Agrupamos TODO el universo de la Planta, ignorando el filtro de "Último Puesto"
        ag8_linea = df_ef_f.groupby('Linea').agg({
            'HH_STD_TOTAL':'sum', 
            col_prod_tot:'sum', 
            'HH_Disponibles':'sum',
            'Cant._Prod._A1':'sum'
        }).reset_index()
        
        ag8_linea = ag8_linea[ag8_linea['HH_Disponibles'] > 0]
        
        if not ag8_linea.empty:
            # A: Ef Real (en decimales)
            ag8_linea['A_val'] = np.where(ag8_linea['HH_Disponibles'] > 0, ag8_linea['HH_STD_TOTAL'] / ag8_linea['HH_Disponibles'], 0)
            # B: Ef Prod (en decimales)
            ag8_linea['B_val'] = np.where(ag8_linea[col_prod_tot] > 0, ag8_linea['HH_STD_TOTAL'] / ag8_linea[col_prod_tot], 0)
            # Diferencia % para pintar la barra (B - A)
            ag8_linea['Dif_pct'] = (ag8_linea['B_val'] - ag8_linea['A_val']) * 100
            # Factor C (Diferencia multiplicada por cantidad, la matemática pedida)
            ag8_linea['C_val'] = (ag8_linea['B_val'] - ag8_linea['A_val']) * ag8_linea['Cant._Prod._A1']
            
            # ORDENAR PURAMENTE POR FACTOR C (El mayor C arriba de todo -> sort_values ascending=True)
            ag8_linea = ag8_linea.sort_values('C_val', ascending=True).reset_index(drop=True)
                
            fig8, ax8 = plt.subplots(figsize=(14, 10))
            fig8.subplots_adjust(top=0.85, bottom=0.15, left=0.25, right=0.95)
            
            # Las barras ahora son una escala en % (Dif_pct)
            colors = ['firebrick' if val > 0 else 'darkgreen' for val in ag8_linea['Dif_pct']]
            bars = ax8.barh(ag8_linea['Linea'], ag8_linea['Dif_pct'], color=colors, edgecolor='white', height=0.6)
            ax8.axvline(0, color='black', linewidth=2, zorder=1)
            
            for i, (idx, row) in enumerate(ag8_linea.iterrows()):
                ef_a = row['A_val'] * 100
                ef_b = row['B_val'] * 100
                c_val = row['C_val']
                dif = row['Dif_pct']
                
                estado_c = "PERDIDAS" if c_val >= 0 else "GANADAS"
                c_100 = abs(c_val)
                c_85 = abs(c_val * 0.85)
                
                # Caja de texto DIVIDIDA EN MÚLTIPLES LÍNEAS PARA NO DESARMAR EL GRÁFICO
                txt_a = f"[A] HH STD / HH DISP: {ef_a:.1f}%"
                txt_b = f"[B] HH STD / HH PROD: {ef_b:.1f}%"
                txt_c = f"DIFERENCIA (C):\n➤ {c_100:.1f} U. {estado_c} AL 100% EF. REAL\n➤ {c_85:.1f} U. {estado_c} AL 85% EF. REAL"
                full_txt = f"{txt_a}\n{txt_b}\n{txt_c}"
                
                ha_align = 'left' if dif >= 0 else 'right'
                offset_x = 15 if dif >= 0 else -15
                
                ax8.annotate(full_txt, 
                             xy=(dif, i), xytext=(offset_x, 0), textcoords="offset points", 
                             va='center', ha=ha_align, color='gold', fontweight='bold', fontsize=10, 
                             path_effects=efecto_n, bbox=dict(boxstyle="round,pad=0.5", fc="black", ec="gold", lw=1.5, alpha=0.85), zorder=10)

            # Eje Y con Nombres y Cantidad en el extremo izquierdo
            yticklabels = [f"{textwrap.fill(str(row['Linea']), 20)}\n(Cant: {int(row['Cant._Prod._A1'])} U)" for _, row in ag8_linea.iterrows()]
            ax8.set_yticks(np.arange(len(ag8_linea)))
            ax8.set_yticklabels(yticklabels, fontsize=12, fontweight='bold')
            ax8.xaxis.set_major_formatter(mtick.PercentFormatter())
            
            # Detectar al CUELLO DE BOTELLA SEGÚN FÓRMULA C
            idx_max_cb = ag8_linea['C_val'].idxmax()
            peor_linea_cb = ag8_linea.loc[idx_max_cb, 'Linea']
            peor_cb_val = ag8_linea.loc[idx_max_cb, 'C_val']
            
            ax8.set_title(f"⚠️ CUELLO DE BOTELLA: {peor_linea_cb} (C: {peor_cb_val:.1f})", color="firebrick", fontweight="bold", fontsize=18, pad=20)
            
            max_abs = ag8_linea['Dif_pct'].abs().max()
            if max_abs == 0: max_abs = 10
            ax8.set_xlim(-max_abs*1.8 - 15, max_abs*1.8 + 15)
            
            agregar_sello_agua(fig8); st.pyplot(fig8, use_container_width=True)
        else:
            st.warning("⚠️ Sin datos de producción para comparar líneas.")
            
    else:
        # Selector para Línea/Puesto que alinea perfectamente con Métrica 9
        tipo_comp_8 = st.radio("Selector M8:", ["Eficiencia Real (HH Disp.)", "Eficiencia Productiva (HH Prod.)"], horizontal=True, label_visibility="collapsed", key="radio_m8")
        st.markdown("<div class='sub-title'><i>Desviación Tiempos Reales vs Estándar por Unidad</i></div>", unsafe_allow_html=True)
        
        c_std_u = next((c for c in df_plot_1.columns if 'STD' in str(c).upper() and ('UNID' in str(c).upper() or 'UNIT' in str(c).upper() or '/ U' in str(c).upper())), None)
        ag8 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL':'sum', col_prod_tot:'sum', 'HH_Disponibles':'sum', 'Cant._Prod._A1':'sum'}).reset_index()
        
        if c_std_u:
            ag_std_u = df_plot_1.groupby('Fecha')[c_std_u].mean().reset_index()
            ag8 = pd.merge(ag8, ag_std_u, on='Fecha', how='left')
            ag8['HH_Std_U'] = ag8[c_std_u]
        else:
            ag8['HH_Std_U'] = ag8['HH_STD_TOTAL'] / ag8['Cant._Prod._A1']
            
        ag8 = ag8[ag8['Cant._Prod._A1'] > 0]
        
        if not ag8.empty:
            ag8['HH_Prod_U'] = ag8[col_prod_tot] / ag8['Cant._Prod._A1']
            ag8['HH_Disp_U'] = ag8['HH_Disponibles'] / ag8['Cant._Prod._A1']
            
            fig8, ax8 = plt.subplots(figsize=(14, 10))
            fig8.subplots_adjust(top=0.85, bottom=0.15, left=0.08, right=0.92)
            fig8.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
            x_idx = np.arange(len(ag8))
            
            # DETERMINAR QUÉ LÍNEA GRAFICAR BASADO EN EL SELECTOR
            if "Disp" in tipo_comp_8:
                y_real = ag8['HH_Disp_U']
                real_color = 'black' # Línea NEGRA solicitada para HH Disponibles
                real_lbl = 'Tiempo DISP / Unidad (Real Global)'
                tot_real_horas = ag8['HH_Disponibles'].sum()
            else:
                y_real = ag8['HH_Prod_U']
                real_color = 'darkgreen'
                real_lbl = 'Tiempo PROD / Unidad (Trabajo Puro)'
                tot_real_horas = ag8[col_prod_tot].sum()
            
            # GRAFICAMOS SÓLO 2 LÍNEAS (MANO A MANO)
            ax8.plot(x_idx, y_real, color=real_color, marker='o', markersize=12, linewidth=5, path_effects=efecto_b, label=real_lbl, zorder=6)
            ax8.plot(x_idx, ag8['HH_Std_U'], color='midnightblue', linestyle='--', linewidth=4, label='Tiempo STD / Unidad (Meta)', zorder=4)
            
            ax8.fill_between(x_idx, ag8['HH_Std_U'], y_real, where=(y_real > ag8['HH_Std_U']), color='red', alpha=0.15, interpolate=True)
            ax8.fill_between(x_idx, ag8['HH_Std_U'], y_real, where=(y_real <= ag8['HH_Std_U']), color='green', alpha=0.15, interpolate=True)
            
            # MOTOR ANTI-COLISIONES DE ETIQUETAS
            for i in range(len(x_idx)): 
                val_r = y_real.iloc[i]
                val_s = ag8['HH_Std_U'].iloc[i]
                cant_u = int(ag8['Cant._Prod._A1'].iloc[i])
                
                # Lógica para separar etiquetas si están muy juntas (Ajustado)
                if val_r >= val_s:
                    off_r = 20; off_s = -25
                else:
                    off_r = -25; off_s = 20
                    
                ax8.annotate(f"{val_r:.2f}h\n({cant_u} Unid)", (x_idx[i], val_r), textcoords="offset points", xytext=(0,off_r), ha='center', fontweight='bold', fontsize=10, bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=real_color, lw=1.5), zorder=10)
                ax8.annotate(f"{val_s:.2f}h", (x_idx[i], val_s), textcoords="offset points", xytext=(0,off_s), ha='center', fontweight='bold', fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="midnightblue", lw=1.5), zorder=10)
                
                # Diferencia graficada a un lado
                diff_val = val_r - val_s
                mid_y = (val_r + val_s) / 2
                ax8.plot([x_idx[i], x_idx[i]], [val_s, val_r], color='dodgerblue', linestyle=':', linewidth=3, zorder=3)
                ax8.annotate(f"{diff_val:+.2f}h", (x_idx[i] + 0.08, mid_y), color='dodgerblue', fontweight='bold', fontsize=11, path_effects=efecto_b, zorder=4)

            # CÁLCULO INTELIGENTE DEL CARTEL SEGÚN EL SELECTOR
            tot_std = ag8['HH_STD_TOTAL'].sum()
            tot_cant = ag8['Cant._Prod._A1'].sum()
            std_u = tot_std / tot_cant if tot_cant > 0 else 0
            
            if std_u > 0:
                c_val = (tot_real_horas - tot_std) / std_u
            else:
                c_val = 0
                
            c_85 = c_val * 0.85

            estado_c = "PERDIDAS" if c_val >= 0 else "GANADAS"
            cartel_col = "#B71C1C" if c_val >= 0 else "#1B5E20"
            
            if "Disp" in tipo_comp_8:
                cartel_txt = f"⚠️ DIFERENCIA (EF. REAL): {abs(c_val):.1f} U. {estado_c} AL 100% / {abs(c_85):.1f} U. {estado_c} AL 85%"
            else:
                cartel_txt = f"⚠️ DIFERENCIA (EF. PROD): {abs(c_val):.1f} U. {estado_c} AL 100% / {abs(c_85):.1f} U. {estado_c} AL 85%"
                
            ax8.text(0.5, 0.95, cartel_txt, transform=ax8.transAxes, ha='center', va='top', bbox=dict(boxstyle="round,pad=0.5", fc=cartel_col, ec="white", lw=2), color="white", fontsize=15, fontweight='bold', zorder=20)
            
            min_y = min(y_real.min(), ag8['HH_Std_U'].min()) * 0.7
            max_y = max(y_real.max(), ag8['HH_Std_U'].max()) * 1.3
            ax8.set_ylim(min_y, max_y)
            
            # EXPANSIÓN DE BORDES PARA EVITAR QUE SE CORTEN LAS ETIQUETAS LATERALES
            ax8.set_xlim(-0.5, len(x_idx) - 0.5)

            ax8.set_xticks(x_idx); ax8.set_xticklabels(ag8['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
            ax8.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
            agregar_sello_agua(fig8); st.pyplot(fig8, use_container_width=True)
        else: st.warning("⚠️ Sin datos de producción (Cant. Producida) para esta selección.")

with col8:
    st.header("9. RANKING DE EFICIENCIA")
    
    # Switch/Toggle selector de tipo de métrica
    tipo_ranking = st.radio("Selector:", ["Eficiencia Real", "Eficiencia Productiva"], horizontal=True, label_visibility="collapsed")
    
    sub_txt = "Competencia Sectorial vs Meta 85%" if tipo_ranking == "Eficiencia Real" else "Competencia Sectorial vs Meta 100%"
    st.markdown(f"<div class='sub-title'><i>{sub_txt}</i></div>", unsafe_allow_html=True)
    
    agrupar_por = 'Puesto_Trabajo' if (s_li or s_pu) else 'Linea'
    if agrupar_por in df_ef_f.columns:
        if agrupar_por == 'Linea':
            lineas_con_si = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']['Linea'].unique()
            m1 = df_ef_f['Linea'].isin(lineas_con_si) & (df_ef_f['Es_Ultimo_Puesto'] == 'SI')
            m2 = ~df_ef_f['Linea'].isin(lineas_con_si)
            df_ranking = df_ef_f[m1 | m2]
        else:
            df_ranking = df_ef_f
            
        ag9 = df_ranking.groupby(agrupar_por).agg({'HH_STD_TOTAL':'sum', 'HH_Disponibles':'sum', col_prod_tot:'sum', 'Cant._Prod._A1':'sum'}).reset_index()
        
        if tipo_ranking == "Eficiencia Real":
            ag9 = ag9[ag9['HH_Disponibles'] > 0]
            ag9['Ef'] = (ag9['HH_STD_TOTAL'] / ag9['HH_Disponibles']) * 100
            meta_val = 85
        else:
            ag9 = ag9[ag9[col_prod_tot] > 0]
            ag9['Ef'] = (ag9['HH_STD_TOTAL'] / ag9[col_prod_tot]) * 100
            meta_val = 100
            
        ag9 = ag9.sort_values('Ef', ascending=True).reset_index(drop=True)
        
        if not ag9.empty:
            if tipo_ranking == "Eficiencia Real":
                colors = ['firebrick' if val < 60 else 'gold' if val < 70 else 'lightgreen' if val < 85 else 'darkgreen' for val in ag9['Ef']]
            else:
                colors = ['firebrick' if val < 80 else 'gold' if val < 90 else 'lightgreen' if val < 100 else 'darkgreen' for val in ag9['Ef']]
                
            fig9, ax9 = plt.subplots(figsize=(14, 10))
            fig9.subplots_adjust(top=0.85, bottom=0.15, left=0.25, right=0.95)
            fig9.suptitle(t_enc, x=0.06, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
            
            bars = ax9.barh(ag9[agrupar_por], ag9['Ef'], color=colors, edgecolor='white', height=0.6)
            
            ax9.bar_label(bars, fmt='%.1f%%', padding=5, color='black', fontweight='bold', path_effects=efecto_b)
            ax9.axvline(meta_val, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
            ax9.text(meta_val, len(ag9)-0.5, f'META {meta_val}%', color='darkgreen', fontweight='bold', ha='left', va='bottom', rotation=90)
            
            ax9.set_xlim(0, max(110, ag9['Ef'].max()*1.1)); ax9.xaxis.set_major_formatter(mtick.PercentFormatter())
            
            # Eje Y con Nombres y Cantidad Producida anclada
            yticklabels = [f"{textwrap.fill(str(row[agrupar_por]), 20)}\n(Cant: {int(row['Cant._Prod._A1'])} U)" for _, row in ag9.iterrows()]
            ax9.set_yticks(np.arange(len(ag9)))
            ax9.set_yticklabels(yticklabels, fontsize=12, fontweight='bold')
            
            agregar_sello_agua(fig9); st.pyplot(fig9, use_container_width=True)
        else: st.warning("⚠️ Sin datos suficientes para armar el ranking.")

st.markdown("---")

# =========================================================================
# 10. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)
# =========================================================================
st.header("10. DETALLES DE IMPRODUCTIVIDAD")
st.markdown("<div class='sub-title'><i>Apertura de registros detallados con fecha, operario y motor de sugerencia de acciones</i></div>", unsafe_allow_html=True)

if not df_im_f.empty and 'DETALLE' in df_im_f.columns:
    motivos_disp = sorted(df_im_f['TIPO_PARADA'].dropna().unique())
    col_sel_m, _ = st.columns([1, 2])
    motivo_sel = col_sel_m.selectbox("🔍 Filtrar Motivo a detallar:", ["Todos los motivos"] + list(motivos_disp))
    
    df_detalles = df_im_f[df_im_f['TIPO_PARADA'] == motivo_sel] if motivo_sel != "Todos los motivos" else df_im_f.copy()
    
    if not df_detalles.empty:
        df_detalles['FECHA_STR'] = df_detalles['FECHA_EXACTA'].dt.strftime('%d/%m/%Y').fillna('S/D')
        c_pu_det = next((c for c in df_detalles.columns if 'PUESTO' in c), None)
        if not c_pu_det: c_pu_det = 'PUESTO_X'
        if c_pu_det not in df_detalles.columns: df_detalles[c_pu_det] = "S/D"
        
        df_detalles['OPERARIO'] = df_detalles['OPERARIO'].fillna('S/D')
        df_detalles['DETALLE'] = df_detalles['DETALLE'].fillna('S/D')
        df_detalles[c_pu_det] = df_detalles[c_pu_det].fillna('S/D')
        
        ag_det = df_detalles.groupby(['FECHA_STR', 'OPERARIO', c_pu_det, 'DETALLE']).agg({'HH_IMPRODUCTIVAS': 'sum'}).reset_index()
        ag_det = ag_det.sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
        
        t_det = ag_det['HH_IMPRODUCTIVAS'].sum()
        ag_det['%'] = (ag_det['HH_IMPRODUCTIVAS'] / t_det) * 100 if t_det > 0 else 0
        ag_det['Acción Sugerida'] = ag_det['DETALLE'].apply(generar_accion_sugerida)
        
        ag_det = ag_det[['FECHA_STR', 'OPERARIO', c_pu_det, 'DETALLE', 'HH_IMPRODUCTIVAS', '%', 'Acción Sugerida']]
        ag_det.columns = ['Fecha', 'Operario', 'Puesto', 'Detalle Registrado', 'Subtotal HH', '%', 'Acción Sugerida']
        
        fila_tot = pd.DataFrame({'Fecha': ['---'], 'Operario': ['---'], 'Puesto': ['---'], 'Detalle Registrado': ['✅ TOTAL SUMATORIA'], 'Subtotal HH': [t_det], '%': [100.0], 'Acción Sugerida': ['🎯 ACCIÓN GLOBAL']})
        ag_det = pd.concat([ag_det, fila_tot], ignore_index=True)
        
        max_hh_value = ag_det['Subtotal HH'].iloc[:-1].max() 
        def highlight_max_cell(val): return 'background-color: rgba(211, 47, 47, 0.5); color: white; font-weight: bold;' if val == max_hh_value else ''
            
        styled_table = ag_det.style.map(highlight_max_cell, subset=['Subtotal HH'])
        
        st.dataframe(styled_table, use_container_width=True, hide_index=True, column_config={"Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), "%": st.column_config.NumberColumn(format="%.1f %%")})
        st.download_button(label="📥 Descargar Detalle Operativo (CSV)", data=ag_det.to_csv(index=False).encode('utf-8'), file_name="Detalles_Operativos.csv", mime="text/csv", use_container_width=True, type="primary")
    else: st.info("No hay registros detallados para el motivo seleccionado en este periodo.")
else: st.info("No hay horas improductivas reportadas con la configuración actual para analizar detalles.")
