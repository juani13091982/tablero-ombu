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
    [data-testid="stMultiSelect"] {margin-bottom: -15px !important;}

    /* --- REGLAS RESPONSIVAS PARA CELULARES --- */
    .kpi-grid {
        display: grid; 
        grid-template-columns: 1fr 1fr 1.3fr; 
        gap: 12px;
    }
    .kpi-costo {
        grid-row: span 2;
    }
    
    /* Clase exclusiva para tablas de celular */
    .mobile-only { display: none !important; }
    
    @media (max-width: 768px) {
        .mobile-only { display: block !important; }
        
        /* FUERZA A LAS COLUMNAS (GRAFICOS Y FILTROS) A APILARSE HACIA ABAJO */
        [data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
        }
        [data-testid="stHorizontalBlock"] > div {
            width: 100% !important;
            min-width: 100% !important;
        }
        /* APILA LOS CARTELES KPI */
        .kpi-grid {
            grid-template-columns: 1fr !important;
        }
        .kpi-costo {
            grid-row: span 1 !important;
        }
        .kpi-grid h2 { font-size: 32px !important; }
        .kpi-grid h4 { font-size: 16px !important; }
    }
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
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
            u_in, p_in = st.text_input("Usuario Corporativo"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")

if not st.session_state['autenticado']: mostrar_login(); st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE Y FUNCIONES
# =========================================================================
def set_escala_y(ax, vmax, factor=1.6): 
    ax.set_ylim(0, vmax * factor if vmax > 0 else 100)
    
def dibujar_meses(ax, n_meses):
    for i in range(n_meses): ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

# FUNCIÓN NUEVA: Limpia las listas de filtros rapidísimo
def normalizar_lista(s_list):
    return [re.sub(r'[^A-Z0-9]', '', str(s).upper()) for s in s_list]

def safe_match(s_list, val):
    if pd.isna(val): return False
    v_norm = re.sub(r'[^A-Z0-9]', '', str(val).upper())
    for s in s_list:
        s_norm = re.sub(r'[^A-Z0-9]', '', str(s).upper())
        if s_norm == v_norm and s_norm != "": return True
    return False

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

# =========================================================================
# 4. CARGA Y LIMPIEZA NUMÉRICA DE DATOS CACHEADA (SÚPER RÁPIDA)
# =========================================================================
@st.cache_data(ttl=300) # ESTO EVITA QUE DESCARGUE EL EXCEL DE GOOGLE DRIVE CADA VEZ
def cargar_datos():
    url_ef = "https://docs.google.com/spreadsheets/d/1_kO7GtjnlGHYgMnY7pkJGdRuFRAJB8xp8pGj8I0jq5E/export?format=xlsx"
    url_im = "https://docs.google.com/spreadsheets/d/1GwshvXAotIShBPcX69vlE8_juO6w5aAFBCaXEsgSdCY/export?format=xlsx"
    df_ef = pd.read_excel(url_ef)
    df_im = pd.read_excel(url_im)
    df_ef.columns = df_ef.columns.str.strip()
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    for col in ['HH_STD_TOTAL', 'HH_Disponibles', 'Cant._Prod._A1', 'HH_Productivas_C/GAP', 'Costo_Improd._$']:
        if col in df_ef.columns:
            df_ef[col] = pd.to_numeric(df_ef[col], errors='coerce').fillna(0)
            
    if 'TIPO_PARADA' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), df_im.columns[0]): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), df_im.columns[0]): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'DETALLE' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c or 'DESC' in c), df_im.columns[0]): 'DETALLE'}, inplace=True)
        
    df_im['HH_IMPRODUCTIVAS'] = pd.to_numeric(df_im['HH_IMPRODUCTIVAS'], errors='coerce').fillna(0).abs()
    c_pu_det = next((c for c in df_im.columns if 'PUESTO' in c), None)
    if not c_pu_det: c_pu_det = 'PUESTO_X'
    if c_pu_det not in df_im.columns: df_im[c_pu_det] = "S/D"
    
    c_nom = next((c for c in df_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in df_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: 
        df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '') + ' ' + df_im[c_ape].astype(str).replace('nan', '')
    elif c_nom: 
        df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '')
    else: 
        df_im['OPERARIO'] = "S/D"
    df_im['OPERARIO'] = df_im['OPERARIO'].str.strip().replace('', 'S/D')

    c_fec = None
    for c in df_im.columns:
        if 'A3' in str(c).upper() or 'INICIO' in str(c).upper():
            c_fec = c; break
    if not c_fec:
        for c in df_im.columns:
            if 'FECHA' in str(c).upper():
                c_fec = c; break

    df_im['FECHA_EXACTA'] = pd.to_datetime(df_im[c_fec], errors='coerce', dayfirst=True) if c_fec else pd.NaT
    if 'FECHA' in df_im.columns:
        df_im['FECHA'] = pd.to_datetime(df_im['FECHA'], errors='coerce', dayfirst=True).dt.to_period('M').dt.to_timestamp()
    else:
        df_im['FECHA'] = df_im['FECHA_EXACTA'].dt.to_period('M').dt.to_timestamp()
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce', dayfirst=True).dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_im['MES_STR'] = df_im['FECHA'].dt.strftime('%b-%Y') 

    # COLUMNAS INVISIBLES DE NORMALIZACIÓN (ACELERADOR RAM)
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

    return df_ef, df_im

try:
    df_ef, df_im = cargar_datos()
except Exception as e: 
    st.error(f"Error crítico cargando datos: {e}"); st.stop()

# =========================================================================
# 5. PANEL STICKY Y FILTROS EN CASCADA UNIFICADA
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h_l, h_t, h_s = st.columns([0.8, 3.5, 0.7])
    with h_l:
        try: st.image("LOGO OMBÚ.jpg", width=90)
        except: st.markdown("##### OMBÚ")
    with h_t: 
        st.markdown("<h3 style='margin:0; padding:0;'>TABLERO INTEGRADO C.G.P.</h3>", unsafe_allow_html=True)
    with h_s:
        if st.button("🚪 Salir", use_container_width=True): 
            st.session_state['autenticado'] = False; st.rerun()

    col_kpi, col_filtros = st.columns([3.5, 1], gap="large")
    with col_filtros:
        st.markdown("<div style='color:#4B8BBE; font-size:16px; font-weight:bold; margin-bottom:5px;'>🎛️ FILTROS MAESTROS</div>", unsafe_allow_html=True)
        meses_disp = sorted(list(set(df_ef['Mes_Str'].dropna().unique()) | set(df_im['MES_STR'].dropna().unique())))
        s_mes = st.multiselect("Mes", ["🎯 Acumulado YTD"] + meses_disp, label_visibility="collapsed", placeholder="📅 Seleccionar Mes")
        
        df_base_ef = df_ef.copy()
        df_base_im = df_im.copy()
        
        if s_mes and "🎯 Acumulado YTD" not in s_mes:
            df_base_ef = df_base_ef[df_base_ef['Mes_Str'].isin(s_mes)]
            df_base_im = df_base_im[df_base_im['MES_STR'].isin(s_mes)]
            
        c_pl_im = next((c for c in df_im.columns if 'PLANTA' in str(c).upper()), df_im.columns[0] if len(df_im.columns)>0 else None)
        pl_ef = set(df_base_ef['Planta'].dropna().astype(str).unique())
        pl_im = set(df_base_im[c_pl_im].dropna().astype(str).unique()) if c_pl_im and not df_base_im.empty else set()
        plantas_disp = sorted(list(pl_ef | pl_im))
        
        # FILTRO DE PLANTA (OPTIMIZADO)
        s_pl = st.multiselect("Planta", plantas_disp, label_visibility="collapsed", placeholder="🏭 Seleccionar Planta")
        if s_pl:
            norm_pl = normalizar_lista(s_pl)
            df_base_ef = df_base_ef[df_base_ef['Planta'].isin(s_pl)]
            if 'NORM_PLANTA' in df_base_im.columns and not df_base_im.empty: 
                df_base_im = df_base_im[df_base_im['NORM_PLANTA'].isin(norm_pl)]
                
        c_li_im = next((c for c in df_im.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), df_im.columns[1] if len(df_im.columns)>1 else None)
        li_ef = set(df_base_ef['Linea'].dropna().astype(str).unique())
        li_im = set(df_base_im[c_li_im].dropna().astype(str).unique()) if c_li_im and not df_base_im.empty else set()
        lineas_disp = sorted(list(li_ef | li_im))
        
        # FILTRO DE LÍNEA (OPTIMIZADO)
        s_li = st.multiselect("Línea", lineas_disp, label_visibility="collapsed", placeholder="⚙️ Seleccionar Línea")
        if s_li:
            norm_li = normalizar_lista(s_li)
            df_base_ef = df_base_ef[df_base_ef['Linea'].isin(s_li)]
            if 'NORM_LINEA' in df_base_im.columns and not df_base_im.empty: 
                df_base_im = df_base_im[df_base_im['NORM_LINEA'].isin(norm_li)]
                
        c_pu_im = next((c for c in df_im.columns if 'PUESTO' in str(c).upper()), df_im.columns[2] if len(df_im.columns)>2 else None)
        pu_ef = set(df_base_ef['Puesto_Trabajo'].dropna().astype(str).unique())
        pu_im = set(df_base_im[c_pu_im].dropna().astype(str).unique()) if c_pu_im and not df_base_im.empty else set()
        puestos_disp = sorted(list(pu_ef | pu_im))
        
        s_pu = st.multiselect("Puesto", puestos_disp, label_visibility="collapsed", placeholder="🛠️ Seleccionar Puesto")

    # APLICACIÓN DE FILTROS A DF FINALES (OPTIMIZADO)
    df_ef_f = df_ef.copy()
    df_im_f = df_im.copy()
    
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

    warn_linea = False
    if s_pu: 
        df_plot_1 = df_ef_f.copy()
    elif s_li:
        df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
        if not df_salida.empty: df_plot_1 = df_salida
        else: df_plot_1 = df_ef_f.copy(); warn_linea = True
    else: 
        df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
        if not df_salida.empty: df_plot_1 = df_salida
        else: df_plot_1 = df_ef_f.copy()

    # =========================================================================
    # PRE-CÁLCULOS EXCLUSIVOS PARA TABLAS MÓVILES
    # =========================================================================
    ag5_mobile = pd.DataFrame()
    if not df_im_f.empty:
        ag5_mobile = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        nm_m = df_im_f['FECHA'].nunique()
        div_m = nm_m if nm_m > 0 else 1
        ag5_mobile['Prom_M'] = ag5_mobile['HH_IMPRODUCTIVAS'] / div_m
        ag5_mobile = ag5_mobile.sort_values(by='Prom_M', ascending=False)
        ag5_mobile['Pct_Acu'] = (ag5_mobile['Prom_M'].cumsum() / ag5_mobile['Prom_M'].sum()) * 100

    df6_mobile = pd.DataFrame()
    if not df_ef_f.empty:
        df_ef_m = df_ef_f.copy()
        df_ef_m['K_Mes'] = df_ef_m['Fecha'].dt.strftime('%Y-%m')
        ag_disp_m = df_ef_m.groupby('K_Mes', as_index=False)['HH_Disponibles'].sum()
        
        if not df_im_f.empty:
            df_im_m = df_im_f.copy()
            df_im_m['K_Mes'] = df_im_m['FECHA'].dt.strftime('%Y-%m')
            piv_m = pd.pivot_table(df_im_m, values='HH_IMPRODUCTIVAS', index='K_Mes', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6_mobile = pd.merge(ag_disp_m, piv_m, on='K_Mes', how='left').fillna(0)
            list_c_m = [c for c in df6_mobile.columns if c not in ['HH_Disponibles', 'K_Mes']]
        else: 
            df6_mobile = ag_disp_m.copy()
            list_c_m = []
            
        df6_mobile['Suma_I'] = df6_mobile[list_c_m].sum(axis=1) if list_c_m else 0
        df6_mobile['Inc_%'] = (df6_mobile['Suma_I'] / df6_mobile['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6_mobile['Fecha_O'] = pd.to_datetime(df6_mobile['K_Mes'] + '-01')
        df6_mobile = df6_mobile.sort_values(by='Fecha_O')

    mobile_tables_html = "<div class='mobile-only' style='margin-top: 15px;'>"
    if not ag5_mobile.empty:
        mobile_tables_html += "<div style='background: #B71C1C; padding: 15px; border-radius: 8px; margin-bottom: 15px; color: white; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);'>"
        mobile_tables_html += "<h4 style='margin:0 0 10px 0; text-align:center; font-size:16px; color:#FFCDD2; border-bottom:1px solid rgba(255,255,255,0.2); padding-bottom:5px;'>📊 5. PARETO DE CAUSAS (MÓVIL)</h4>"
        mobile_tables_html += "<table style='width:100%; border-collapse: collapse; font-size:13px;'>"
        mobile_tables_html += "<tr style='border-bottom: 1px solid rgba(255,255,255,0.3);'> <th style='text-align:left; padding:6px;'>Motivo</th> <th style='text-align:right; padding:6px;'>HH</th> <th style='text-align:right; padding:6px;'>% Acum</th> </tr>"
        for _, row in ag5_mobile.iterrows():
            motivo = str(row['TIPO_PARADA'])[:20] + (".." if len(str(row['TIPO_PARADA'])) > 20 else "")
            mobile_tables_html += f"<tr style='border-bottom: 1px solid rgba(255,255,255,0.1);'> <td style='padding:6px;'>{motivo}</td> <td style='text-align:right; padding:6px;'>{row['Prom_M']:.1f}</td> <td style='text-align:right; padding:6px; font-weight:bold;'>{row['Pct_Acu']:.1f}%</td> </tr>"
        mobile_tables_html += "</table></div>"

    if not df6_mobile.empty:
        mobile_tables_html += "<div style='background: #0D47A1; padding: 15px; border-radius: 8px; color: white; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);'>"
        mobile_tables_html += "<h4 style='margin:0 0 10px 0; text-align:center; font-size:16px; color:#BBDEFB; border-bottom:1px solid rgba(255,255,255,0.2); padding-bottom:5px;'>📈 6. EVOLUCIÓN INCIDENCIA (MÓVIL)</h4>"
        mobile_tables_html += "<table style='width:100%; border-collapse: collapse; font-size:13px;'>"
        mobile_tables_html += "<tr style='border-bottom: 1px solid rgba(255,255,255,0.3);'> <th style='text-align:left; padding:6px;'>Mes</th> <th style='text-align:right; padding:6px;'>HH Imp</th> <th style='text-align:right; padding:6px;'>Incid.</th> </tr>"
        for _, row in df6_mobile.iterrows():
            mobile_tables_html += f"<tr style='border-bottom: 1px solid rgba(255,255,255,0.1);'> <td style='padding:6px;'>{row['K_Mes']}</td> <td style='text-align:right; padding:6px;'>{row['Suma_I']:.1f}</td> <td style='text-align:right; padding:6px; font-weight:bold; color:#FFCDD2;'>{row['Inc_%']:.1f}%</td> </tr>"
        mobile_tables_html += "</table></div>"
    mobile_tables_html += "</div>"

    # CÁLCULOS PONDERADOS UNIVERSALES PARA CARTELES
    tot_costo = df_ef_f['Costo_Improd._$'].sum() if not df_ef_f.empty else 0
    tot_hh_imp = df_im_f['HH_IMPRODUCTIVAS'].sum() if not df_im_f.empty else 0
    
    tot_std = df_plot_1['HH_STD_TOTAL'].sum() if not df_plot_1.empty else 0
    tot_disp = df_plot_1['HH_Disponibles'].sum() if not df_plot_1.empty else 0
    tot_prod = df_plot_1['HH_Productivas_C/GAP'].sum() if ('HH_Productivas_C/GAP' in df_plot_1.columns and not df_plot_1.empty) else 0
    
    kpi_ef_real = (tot_std / tot_disp * 100) if tot_disp > 0 else 0
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

    with col_kpi:
        st.markdown(f"""
        <div class="kpi-grid">
            <div style="background: linear-gradient(135deg, #e0e0e0, #f5f5f5); border: 1px solid #aaa; border-left: 6px solid #1E3A8A; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0; color: #1E3A8A; font-size:16px;">EFICIENCIA REAL</h4>
                <h2 style="margin:5px 0 0 0; color: #111; font-size:42px;">{kpi_ef_real:.1f}%</h2>
            </div>
            <div style="background: linear-gradient(135deg, #2E7D32, #4CAF50); border: 1px solid #1B5E20; border-left: 6px solid #A5D6A7; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0; color: white; font-size:16px;">EFICIENCIA PROD.</h4>
                <h2 style="margin:5px 0 0 0; color: white; font-size:42px;">{kpi_ef_prod:.1f}%</h2>
            </div>
            <div class="kpi-costo" style="background: linear-gradient(135deg, #D32F2F, #E53935); border: 1px solid #B71C1C; padding: 15px; border-radius: 8px; display: flex; flex-direction: column; justify-content: center; text-align:center; box-shadow: 2px 4px 15px rgba(211,47,47,0.4);">
                <h4 style="margin:0; color: white; font-size:22px;">COSTO HH IMPROD.</h4>
                <p style="margin:0; color: #FFCDD2; font-size:14px;">(Oportunidad Perdida)</p>
                <h2 style="margin:10px 0; color: #FFEB3B; font-size:48px; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${tot_costo:,.0f}</h2>
                <h4 style="margin:0; color: white; font-size:22px;">{tot_hh_imp:,.1f} <span style="font-size:16px; font-weight:normal;">HH</span></h4>
            </div>
            <div style="background: #0D47A1; color: white; padding: 15px; border-radius: 6px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0 0 8px 0; font-size:14px; color:#BBDEFB; text-align:center; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom:5px;">🏆 TOP EF. REAL (PUESTOS)</h4>
                {top3_m1_html}
            </div>
            <div style="background: #B71C1C; color: white; padding: 15px; border-radius: 6px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0 0 8px 0; font-size:14px; color:#FFCDD2; text-align:center; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom:5px;">⚠️ TOP MAYOR HH IMP.</h4>
                {top3_imp_html}
            </div>
        </div>
        {mobile_tables_html}
        """, unsafe_allow_html=True)

t_enc = f"Filtros >> Planta: {'+'.join(s_pl) if s_pl else 'Todas'} | Línea: {'+'.join(s_li) if s_li else 'Todas'} | Puesto: {'+'.join(s_pu) if s_pu else 'Todos'}"

# =========================================================================
# 6. GRÁFICOS MÉTRICAS 1 Y 2
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
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
            if vu > 0: 
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{vu} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)
        
        ax1_line.plot(x_idx, ag1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real', zorder=5)
        add_tendencia(ax1_line, x_idx, ag1['Ef_Real'])
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        
        last_x1 = x_idx[-1] if len(x_idx) > 0 else 0
        ax1_line.text(last_x1, 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3))
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        
        for i, val in enumerate(ag1['Ef_Real']): 
            ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
            
        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig1); st.pyplot(fig1, use_container_width=True)
    else: st.warning("⚠️ Sin datos para Eficiencia Real. Pruebe otra combinación de filtros.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    if not df_plot_1.empty:
        c_prod = next((c for c in df_plot_1.columns if 'GAP' in str(c).upper() and 'PROD' in str(c).upper()), 'HH_Productivas_C/GAP')
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', c_prod: 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2[c_prod]).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2 = plt.subplots(figsize=(14, 10)); ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag2))
        bs = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bp = ax2.bar(x_idx + 0.17, ag2[c_prod], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2[c_prod].max()), 1.6)
        ax2.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax2.bar_label(bp, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax2, len(x_idx))
        
        ax2_line.plot(x_idx, ag2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Prod.', zorder=5)
        add_tendencia(ax2_line, x_idx, ag2['Ef_Prod'])
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        
        last_x2 = x_idx[-1] if len(x_idx) > 0 else 0
        ax2_line.text(last_x2, 101, 'META = 100%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        
        ax2_line.set_ylim(0, max(110, ag2['Ef_Prod'].max()*1.3))
        ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        
        for i, val in enumerate(ag2['Ef_Prod']): 
            ax2_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
            
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
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH_Productivas_C/GAP'
        if c_prod not in df_ef_f.columns: c_prod = df_ef_f.columns[-1] # fallback
        ag3 = df_ef_f.groupby('Fecha').agg({c_prod: 'sum', 'HH_Disponibles': 'sum'}).reset_index()
        
        if not df_im_f.empty:
            ag_im = df_im_f.groupby('FECHA')['HH_IMPRODUCTIVAS'].sum().reset_index().rename(columns={'FECHA':'Fecha', 'HH_IMPRODUCTIVAS':'HH_Imp'})
            ag3 = pd.merge(ag3, ag_im, on='Fecha', how='left').fillna(0)
        else: 
            ag3['HH_Imp'] = 0
            
        ag3['Total_Decl'] = ag3[c_prod] + ag3['HH_Imp']
        
        fig3, ax3 = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag3))
        
        bp = ax3.bar(x_idx, ag3[c_prod], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        bi = ax3.bar(x_idx, ag3['HH_Imp'], bottom=ag3[c_prod], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax3.bar_label(bp, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.bar_label(bi, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        
        ax3.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='HH DISPONIBLES', zorder=5)
        set_escala_y(ax3, ag3['HH_Disponibles'].max(), 1.6); dibujar_meses(ax3, len(x_idx))

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
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Evolución de valorización económica de la ineficiencia</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({'Costo_Improd._$': 'sum'}).reset_index()
        
        if not df_im_f.empty:
            ag_im = df_im_f.groupby('FECHA')['HH_IMPRODUCTIVAS'].sum().reset_index().rename(columns={'FECHA':'Fecha', 'HH_IMPRODUCTIVAS':'HH_Imp'})
            ag4 = pd.merge(ag4, ag_im, on='Fecha', how='left').fillna(0)
        else: 
            ag4['HH_Imp'] = 0

        fig4, ax4 = plt.subplots(figsize=(14, 10)); ax4_line = ax4.twinx()
        fig4.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag4))
        bi = ax4.bar(x_idx, ag4['HH_Imp'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4.bar_label(bi, padding=4, color='black', fontweight='bold', path_effects=efecto_b, zorder=4)
        set_escala_y(ax4, ag4['HH_Imp'].max(), 1.6) 
        
        ax4_line.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, label='COSTO ARS', zorder=5)
        add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'])
        ax4_line.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3))
        ax4_line.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_line.get_yticks()], fontweight='bold')

        for i, val in enumerate(ag4['Costo_Improd._$']): 
            ax4_line.annotate(f"${val:,.0f}", (x_idx[i], val + (ax4_line.get_ylim()[1]*0.04)), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax4.set_xticks(x_idx); ax4.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
        ax4_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig4); st.pyplot(fig4, use_container_width=True)
    else: st.warning("⚠️ No hay datos económicos.")

st.markdown("---")

# =========================================================================
# 8. GRÁFICOS MÉTRICAS 5 Y 6
# =========================================================================
col5, col6 = st.columns(2)
with col5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)
    
    if not df_im_f.empty:
        ag5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        nm = df_im_f['FECHA'].nunique(); div = nm if nm > 0 else 1
        ag5['Prom_M'] = ag5['HH_IMPRODUCTIVAS'] / div; ag5 = ag5.sort_values(by='Prom_M', ascending=False)
        ag5['Pct_Acu'] = (ag5['Prom_M'].cumsum() / ag5['Prom_M'].sum()) * 100
        
        fig5, ax5 = plt.subplots(figsize=(14, 10)); ax5_line = ax5.twinx()
        fig5.subplots_adjust(top=0.75, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag5))
        bp = ax5.bar(x_idx, ag5['Prom_M'], color='maroon', edgecolor='white', zorder=2)
        set_escala_y(ax5, ag5['Prom_M'].max(), 2.8)
        ax5.bar_label(bp, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        ax5_line.plot(x_idx, ag5['Pct_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax5_line.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        ax5_line.set_ylim(0, 110); ax5_line.yaxis.set_major_formatter(mtick.PercentFormatter()) 
        
        lbls = [textwrap.fill(str(t), 12) for t in ag5['TIPO_PARADA']]
        ax5.set_xticks(x_idx); ax5.set_xticklabels(lbls, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val in enumerate(ag5['Pct_Acu']): 
            ax5_line.annotate(f"{val:.1f}%", (x_idx[i], val + 4), color='white', bbox=caja_g, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
            
        agregar_sello_agua(fig5); st.pyplot(fig5, use_container_width=True)
    else: st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col6:
    st.header("6. EVOL EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        df_ef_f['K_Mes'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp = df_ef_f.groupby('K_Mes', as_index=False)['HH_Disponibles'].sum()
        
        if not df_im_f.empty:
            df_im_f['K_Mes'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K_Mes', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6 = pd.merge(ag_disp, piv, on='K_Mes', how='left').fillna(0)
            list_c = [c for c in df6.columns if c not in ['HH_Disponibles', 'K_Mes']]
        else: 
            df6 = ag_disp.copy(); list_c = []
            
        df6['Suma_I'] = df6[list_c].sum(axis=1) if list_c else 0
        df6['Inc_%'] = (df6['Suma_I'] / df6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6['Fecha_O'] = pd.to_datetime(df6['K_Mes'] + '-01'); df6 = df6.sort_values(by='Fecha_O')
        
        fig6, ax6 = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.68, bottom=0.22, left=0.08, right=0.92)
        fig6.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(df6))
        
        if list_c:
            base_st = np.zeros(len(df6)); paleta = plt.cm.tab20.colors
            for i, c_nom in enumerate(list_c):
                vals = df6[c_nom].values
                bar_stack = ax6.bar(x_idx, vals, bottom=base_st, label=textwrap.fill(c_nom, 15), color=paleta[i % 20], edgecolor='white', zorder=2)
                lbls_stk = [f"{int(v)}" if v > 0 else "" for v in vals]
                ax6.bar_label(bar_stack, labels=lbls_stk, label_type='center', color='white', fontsize=9, fontweight='bold', path_effects=efecto_n)
                base_st += vals
            
            ax6.legend(loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=4, framealpha=0.9, fontsize=11)
        else: 
            ax6.bar(x_idx, np.zeros(len(df6)), color='white')
            
        set_escala_y(ax6, df6['Suma_I'].max(), 1.8) 
        
        for i in range(len(x_idx)):
            v_i, v_d = df6['Suma_I'].iloc[i], df6['HH_Disponibles'].iloc[i]
            if v_i > 0: 
                ax6.annotate(f"Imp: {int(v_i)}\nDisp: {int(v_d)}", (i, v_i + (ax6.get_ylim()[1]*0.02)), ha='center', bbox=caja_o, fontweight='bold', zorder=10)
                
        ax6_line = ax6.twinx(); ax6_line.plot(x_idx, df6['Inc_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, label='% Incidencia', zorder=5)
        add_tendencia(ax6_line, x_idx, df6['Inc_%'])
        ax6_line.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        
        last_x6 = x_idx[-1] if len(x_idx) > 0 else 0
        ax6_line.text(last_x6, 16, 'META = 15%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        
        for i, val in enumerate(df6['Inc_%']): 
            if df6['Suma_I'].iloc[i] > 0: 
                ax6_line.annotate(f"{val:.1f}%", (x_idx[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_b, zorder=10)
                
        ax6.set_xticks(x_idx); ax6.set_xticklabels(df6['K_Mes'], fontsize=14, fontweight='bold')
        ax6_line.set_ylim(0, max(30, df6['Inc_%'].max() * 1.5))
        agregar_sello_agua(fig6); st.pyplot(fig6, use_container_width=True)
    else: st.warning("⚠️ Sin datos históricos de eficiencia para cruzar.")

st.markdown("---")

# =========================================================================
# 9. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)
# =========================================================================
st.header("7. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)")
st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Apertura de registros detallados con fecha, operario y motor de sugerencia de acciones</i></div>", unsafe_allow_html=True)

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
        
        fila_tot = pd.DataFrame({
            'Fecha': ['---'], 'Operario': ['---'], 'Puesto': ['---'], 
            'Detalle Registrado': ['✅ TOTAL SUMATORIA'], 'Subtotal HH': [t_det], 
            '%': [100.0], 'Acción Sugerida': ['🎯 ACCIÓN GLOBAL']
        })
        ag_det = pd.concat([ag_det, fila_tot], ignore_index=True)
        
        max_hh_value = ag_det['Subtotal HH'].iloc[:-1].max() 
        def highlight_max_cell(val):
            return 'background-color: rgba(211, 47, 47, 0.5); color: white; font-weight: bold;' if val == max_hh_value else ''
            
        styled_table = ag_det.style.map(highlight_max_cell, subset=['Subtotal HH'])
        
        st.dataframe(styled_table, use_container_width=True, hide_index=True, column_config={"Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), "%": st.column_config.NumberColumn(format="%.1f %%")})
        st.download_button(label="📥 Descargar Detalle Operativo (CSV)", data=ag_det.to_csv(index=False).encode('utf-8'), file_name="Detalles_Operativos.csv", mime="text/csv", use_container_width=True, type="primary")
    else: st.info("No hay registros detallados para el motivo seleccionado en este periodo.")
else: st.info("No hay horas improductivas reportadas con la configuración actual para analizar detalles.")
