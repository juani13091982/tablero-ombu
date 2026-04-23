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
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o, caja_b = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}
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
# 3. MOTOR INTELIGENTE (VERSIÓN IGUALDAD ESTRICTA)
# =========================================================================
def set_escala_y(ax, vmax, factor=1.6): 
    ax.set_ylim(0, vmax * factor if vmax > 0 else 100)
    
def dibujar_meses(ax, n_meses):
    for i in range(n_meses): ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def safe_match(s_list, val):
    """Filtro ESTRICTO DEFINITIVO: Igualdad matemática 1 a 1"""
    if pd.isna(val): return False
    
    v_norm = " ".join(str(val).strip().upper().split())
    for a, b in zip("ÁÉÍÓÚ", "AEIOU"): v_norm = v_norm.replace(a, b)
    
    for s in s_list:
        s_norm = " ".join(str(s).strip().upper().split())
        for a, b in zip("ÁÉÍÓÚ", "AEIOU"): s_norm = s_norm.replace(a, b)
        
        if s_norm == v_norm: 
            return True
            
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
# 4. CARGA DE DATOS
# =========================================================================
try:
    df_ef = pd.read_excel("eficiencias.xlsx")
    df_im = pd.read_excel("improductivas.xlsx")
    df_ef.columns = df_ef.columns.str.strip()
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    if 'TIPO_PARADA' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), df_im.columns[0]): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), df_im.columns[0]): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'DETALLE' not in df_im.columns: 
        df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c or 'DESC' in c), df_im.columns[0]): 'DETALLE'}, inplace=True)
    
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

    # CORRECCIÓN DE LA LECTURA DE FECHAS (Preserva la FECHA original)
    c_fec_exacta = next((c for c in df_im.columns if 'A3' in str(c).upper() or 'INICIO' in str(c).upper()), None)
    c_fec_base = 'FECHA' if 'FECHA' in df_im.columns else df_im.columns[0]

    df_im['FECHA_EXACTA'] = pd.to_datetime(df_im[c_fec_exacta], errors='coerce') if c_fec_exacta else pd.NaT
    df_im['FECHA'] = pd.to_datetime(df_im[c_fec_base], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_im['Mes_Str'] = df_im['FECHA'].dt.strftime('%b-%Y') 
except Exception as e: 
    st.error(f"Error crítico cargando datos: {e}"); st.stop()

# =========================================================================
# 5. PANEL STICKY Y FILTROS EN CASCADA
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
        
        # --- 1. Filtro Mes ---
        meses_unicos = list(pd.Series(list(df_ef['Mes_Str'].dropna().unique()) + list(df_im['Mes_Str'].dropna().unique())).dropna().unique())
        s_mes = st.multiselect("Mes", ["🎯 Acumulado YTD"] + meses_unicos, label_visibility="collapsed", placeholder="📅 Seleccionar Mes")
        
        df_base_f = df_ef.copy()
        if s_mes and "🎯 Acumulado YTD" not in s_mes:
            df_base_f = df_base_f[df_base_f['Mes_Str'].isin(s_mes)]
            
        # --- 2. Filtro Planta (En cascada según Mes) ---
        c_pl_im = next((c for c in df_im.columns if 'PLANTA' in str(c).upper()), df_im.columns[0] if len(df_im.columns)>0 else None)
        pl_ef = set(df_base_f['Planta'].dropna().astype(str).unique())
        
        # Se obtiene también de improductivas filtradas por mes
        df_im_temp = df_im.copy()
        if s_mes and "🎯 Acumulado YTD" not in s_mes:
            df_im_temp = df_im_temp[df_im_temp['Mes_Str'].isin(s_mes)]
        pl_im = set(df_im_temp[c_pl_im].dropna().astype(str).unique()) if c_pl_im and not df_im_temp.empty else set()
        
        plantas_disp = sorted(list(pl_ef | pl_im))
        
        s_pl = st.multiselect("Planta", plantas_disp, label_visibility="collapsed", placeholder="🏭 Seleccionar Planta")
        
        df_tl = df_base_f[df_base_f['Planta'].isin(s_pl)] if s_pl else df_base_f
        
        # --- 3. Filtro Línea (En cascada) ---
        if s_pl and c_pl_im and not df_im_temp.empty:
            df_im_temp = df_im_temp[df_im_temp[c_pl_im].apply(lambda x: safe_match(s_pl, x))]
            
        c_li_im = next((c for c in df_im.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), df_im.columns[1] if len(df_im.columns)>1 else None)
        li_ef = set(df_tl['Linea'].dropna().astype(str).unique())
        li_im = set(df_im_temp[c_li_im].dropna().astype(str).unique()) if c_li_im and not df_im_temp.empty else set()
        lineas_disp = sorted(list(li_ef | li_im))
        
        s_li = st.multiselect("Línea", lineas_disp, label_visibility="collapsed", placeholder="⚙️ Seleccionar Línea")
        
        df_tp = df_tl[df_tl['Linea'].isin(s_li)] if s_li else df_tl
        
        # --- 4. Filtro Puesto (En cascada) ---
        if s_li and c_li_im and not df_im_temp.empty:
            df_im_temp = df_im_temp[df_im_temp[c_li_im].apply(lambda x: safe_match(s_li, x))]
            
        c_pu_im = next((c for c in df_im.columns if 'PUESTO' in str(c).upper()), df_im.columns[2] if len(df_im.columns)>2 else None)
        pu_ef = set(df_tp['Puesto_Trabajo'].dropna().astype(str).unique())
        pu_im = set(df_im_temp[c_pu_im].dropna().astype(str).unique()) if c_pu_im and not df_im_temp.empty else set()
        puestos_disp = sorted(list(pu_ef | pu_im))
        
        s_pu = st.multiselect("Puesto", puestos_disp, label_visibility="collapsed", placeholder="🛠️ Seleccionar Puesto")

    df_ef_f = df_ef.copy()
    df_im_f = df_im.copy()
    
    if s_pl: df_ef_f = df_ef_f[df_ef_f['Planta'].isin(s_pl)]
    if s_li: df_ef_f = df_ef_f[df_ef_f['Linea'].isin(s_li)]
    if s_pu: df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(s_pu)]
    if s_mes and "🎯 Acumulado YTD" not in s_mes: df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(s_mes)]

    if not df_im_f.empty:
        if s_pl:
            col_pl = next((c for c in df_im_f.columns if 'PLANTA' in str(c).upper()), df_im_f.columns[0])
            df_im_f = df_im_f[df_im_f[col_pl].apply(lambda x: safe_match(s_pl, x))]
        if s_li:
            col_li = next((c for c in df_im_f.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), df_im_f.columns[1])
            df_im_f = df_im_f[df_im_f[col_li].apply(lambda x: safe_match(s_li, x))]
        if s_pu:
            col_pu = next((c for c in df_im_f.columns if 'PUESTO' in str(c).upper()), df_im_f.columns[2])
            df_im_f = df_im_f[df_im_f[col_pu].apply(lambda x: safe_match(s_pu, x))]
        if s_mes and "🎯 Acumulado YTD" not in s_mes: 
            df_im_f = df_im_f[df_im_f['Mes_Str'].isin(s_mes)]

    warn_linea = False
    
    if s_pu: 
        df_plot_1 = df_ef_f.copy()
    elif s_li:
        df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
        if not df_salida.empty: 
            df_plot_1 = df_salida
        else: 
            df_plot_1 = df_ef_f.copy(); warn_linea = True
    else: 
        df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
        if not df_salida.empty: 
            df_plot_1 = df_salida
        else: 
            df_plot_1 = df_ef_f.copy()

    tot_costo = df_ef_f['Costo_Improd._$'].sum() if not df_ef_f.empty else 0
    tot_hh_imp = df_im_f['HH_IMPRODUCTIVAS'].sum() if not df_im_f.empty else 0
    
    # RESTAURACIÓN EXACTA DE LA LÓGICA PARA EL CARTEL (GARANTIZA EL 52.1%)
    if not any([s_pl, s_li, s_pu, s_mes]) and not df_plot_1.empty:
        ag_global = df_plot_1.groupby(['Planta', 'Linea', 'Puesto_Trabajo']).agg({'HH_STD_TOTAL':'sum', 'HH_Disponibles':'sum', 'HH_Productivas_C/GAP':'sum'})
        ef_r_arr = (ag_global['HH_STD_TOTAL'] / ag_global['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        ef_p_arr = (ag_global['HH_STD_TOTAL'] / ag_global['HH_Productivas_C/GAP'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        kpi_ef_real = ef_r_arr.mean() if not ef_r_arr.empty else 0
        kpi_ef_prod = ef_p_arr.mean() if not ef_p_arr.empty else 0
    else:
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
        <div style="display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 12px;">
            <div style="background: linear-gradient(135deg, #e0e0e0, #f5f5f5); border: 1px solid #aaa; border-left: 6px solid #1E3A8A; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0; color: #1E3A8A; font-size:16px;">EFICIENCIA REAL</h4>
                <h2 style="margin:5px 0 0 0; color: #111; font-size:42px;">{kpi_ef_real:.1f}%</h2>
            </div>
            <div style="background: linear-gradient(135deg, #2E7D32, #4CAF50); border: 1px solid #1B5E20; border-left: 6px solid #A5D6A7; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
                <h4 style="margin:0; color: white; font-size:16px;">EFICIENCIA PROD.</h4>
                <h2 style="margin:5px 0 0 0; color: white; font-size:42px;">{kpi_ef_prod:.1f}%</h2>
            </div>
            <div style="background: linear-gradient(135deg, #D32F2F, #E53935); border: 1px solid #B71C1C; padding: 15px; border-radius: 8px; grid-row: span 2; display: flex; flex-direction: column; justify-content: center; text-align:center; box-shadow: 2px 4px 15px rgba(211,47,47,0.4);">
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
        agregar_sello_agua(fig1); st.pyplot(fig1)
    else: st.warning("⚠️ Sin datos para Eficiencia Real. Pruebe otra combinación de filtros.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    if not df_plot_1.empty:
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Productivas_C/GAP': 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2 = plt.subplots(figsize=(14, 10)); ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.80, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag2))
        bs = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bp = ax2.bar(x_idx + 0.17, ag2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 1.6)
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
        agregar_sello_agua(fig2); st.pyplot(fig2)
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
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
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
        agregar_sello_agua(fig3); st.pyplot(fig3)
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
        agregar_sello_agua(fig4); st.pyplot(fig4)
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
            
        agregar_sello_agua(fig5); st.pyplot(fig5)
    else: st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
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
        agregar_sello_agua(fig6); st.pyplot(fig6)
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
