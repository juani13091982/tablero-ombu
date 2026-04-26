import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.subplots as plt_subplots
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import matplotlib.image as mpimg
import textwrap
import re

# =========================================================================
# 1. CONFIGURACIÓN, ESCUDO VISUAL Y CSS DINÁMICO
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")

def inyectar_estilos_integrados(color_per):
    st.markdown(f"""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="manage-app-button"], 
        #MainMenu, footer, .stAppDeployButton, .viewerBadge_container {{display: none !important; visibility: hidden !important;}}
        .block-container {{padding-top: 1rem !important; padding-bottom: 1.5rem !important; background-color: #0f172a;}}
        
        div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {{
            position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
            background-color: rgba(14, 17, 23, 0.98) !important; z-index: 99999 !important;
            padding: 5px 10px 15px 10px !important; border-bottom: 3px solid {color_per} !important;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
        }}
        
        .kpi-grid {{ display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 12px; }}
        .kpi-costo {{ grid-row: span 2; }}
        .mobile-only {{ display: none !important; }}
        
        @media (max-width: 768px) {{
            .mobile-only {{ display: block !important; }}
            [data-testid="stHorizontalBlock"] {{ flex-direction: column !important; }}
            .kpi-grid {{ grid-template-columns: 1fr !important; }}
            .kpi-costo {{ grid-row: span 1 !important; }}
        }}

        .card-oee {{
            background-color: #1e293b; border-radius: 20px; padding: 25px; 
            margin-bottom: 20px; border: 1px solid #334155;
            box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
        }}
        
        .pillar-item {{ 
            background: #0f172a; border-radius: 15px; padding: 18px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 4px solid {color_per} !important;
        }}
        .p-label {{ font-size: 13px; font-weight: 700; color: #94a3b8; text-transform: uppercase; margin-bottom: 5px; }}
        .p-val {{ font-size: 32px; font-weight: 900; color: white; margin: 0; }}
        .p-formula {{ font-size: 11px; color: {color_per}; margin-top: 8px; border-top: 1px solid #334155; padding-top: 8px; font-weight: bold; }}
        .p-desc {{ font-size: 11px; color: #cbd5e1; margin-top: 5px; line-height: 1.3; font-style: italic; }}

        .kpi-card-custom {{
            background: #1e293b; border-radius: 12px; padding: 20px; text-align: center;
            border-top: 4px solid {color_per}; box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        }}
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

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True); c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='text-align:center;'><h2 style='color:white;'>GESTIÓN INDUSTRIAL OMBÚ</h2></div>", unsafe_allow_html=True)
        with st.form("login"):
            u_in, p_in = st.text_input("Usuario Corporativo"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE
# =========================================================================
def safe_match(s_list, val):
    if pd.isna(val): return False
    v_norm = re.sub(r'[^A-Z0-9]', '', str(val).upper())
    for s in s_list:
        s_norm = re.sub(r'[^A-Z0-9]', '', str(s).upper())
        if s_norm == v_norm and s_norm != "": return True
    return False

def add_tendencia(ax, x, y, color_line):
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color=color_line, linestyle=':', linewidth=4, alpha=0.6, zorder=4)

def agregar_sello_agua(fig):
    try:
        img = mpimg.imread("LOGO OMBÚ.jpg")
        ax_logo = fig.add_axes([0.88, 0.02, 0.08, 0.08], zorder=1)
        ax_logo.imshow(img, alpha=0.15); ax_logo.axis('off')
    except: pass

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if '✅' in d: return "🎯 ACCIÓN GLOBAL"
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO', 'MECANICO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'LOGISTICA']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'ERROR']): return "🔎 Ajustar Calidad"
    return "⚡ Investigar Causa"

def set_escala_y(ax, vmax, factor=1.6): 
    ax.set_ylim(0, vmax * factor if vmax > 0 else 100)

def dibujar_meses(ax, n_meses):
    for i in range(n_meses): ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

# =========================================================================
# 4. CARGA Y LIMPIEZA EXTREMA
# =========================================================================
@st.cache_data(ttl=300)
def cargar_datos_drive():
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    d_ef = pd.read_excel(url_ef)
    d_im = pd.read_excel(url_im)
    
    d_ef = d_ef.loc[:, ~d_ef.columns.duplicated()]
    d_im = d_im.loc[:, ~d_im.columns.duplicated()]
    
    d_ef.columns = d_ef.columns.str.strip()
    
    m_ef = {}
    for c in d_ef.columns:
        c_up = str(c).strip().upper()
        if 'FECHA' in c_up: m_ef[c] = 'Fecha'
        elif 'PLANTA' in c_up: m_ef[c] = 'Planta'
        elif 'LÍNEA' in c_up or 'LINEA' in c_up: m_ef[c] = 'Linea'
        elif 'PUESTO' in c_up: m_ef[c] = 'Puesto_Trabajo'
        elif 'ULTIMO' in c_up: m_ef[c] = 'Es_Ultimo_Puesto'
        
    d_ef.rename(columns=m_ef, inplace=True)
    d_ef = d_ef.loc[:, ~d_ef.columns.duplicated()]
    
    for col in ['HH_STD_TOTAL', 'HH_Disponibles', 'Cant._Prod._A1', 'HH_Productivas_C/GAP', 'Costo_Improd._$']:
        if col in d_ef.columns: d_ef[col] = pd.to_numeric(d_ef[col], errors='coerce').fillna(0)
    
    if 'Fecha' in d_ef.columns:
        d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce', dayfirst=True)
        d_ef = d_ef.sort_values('Fecha')
        d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    else:
        d_ef['Mes_Str'] = "S/D"
        
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    m_im = {}
    for c in d_im.columns:
        c_up = str(c).strip().upper()
        if 'HH' in c_up and 'IMP' in c_up: m_im[c] = 'HH_IMPRODUCTIVAS'
        elif 'TIPO' in c_up or 'MOTIVO' in c_up: m_im[c] = 'TIPO_PARADA'
        elif 'DETALLE' in c_up: m_im[c] = 'DETALLE'
        elif 'PLANTA' in c_up: m_im[c] = 'PLANTA'
        elif 'LINEA' in c_up or 'LÍNEA' in c_up: m_im[c] = 'LINEA'
        elif 'NOMBRE' in c_up: m_im[c] = 'NOMBRE'
        elif 'APELLIDO' in c_up: m_im[c] = 'APELLIDO'
        elif 'PUESTO' in c_up: m_im[c] = 'PUESTO'
        
    d_im.rename(columns=m_im, inplace=True)
    d_im = d_im.loc[:, ~d_im.columns.duplicated()]
    
    if 'HH_IMPRODUCTIVAS' not in d_im.columns: d_im['HH_IMPRODUCTIVAS'] = 0.0
    d_im['HH_IMPRODUCTIVAS'] = pd.to_numeric(d_im['HH_IMPRODUCTIVAS'], errors='coerce').fillna(0).abs()
    
    if 'NOMBRE' in d_im.columns and 'APELLIDO' in d_im.columns: 
        d_im['OPERARIO'] = d_im['NOMBRE'].astype(str).replace('nan','') + ' ' + d_im['APELLIDO'].astype(str).replace('nan','')
    elif 'NOMBRE' in d_im.columns: 
        d_im['OPERARIO'] = d_im['NOMBRE'].astype(str).replace('nan','')
    else: 
        d_im['OPERARIO'] = "S/D"
        
    if 'TIPO_PARADA' not in d_im.columns: d_im['TIPO_PARADA'] = "S/D"
    if 'DETALLE' not in d_im.columns: d_im['DETALLE'] = "S/D"
    if 'PUESTO' not in d_im.columns: d_im['PUESTO'] = "S/D"
    
    return d_ef, d_im

try:
    df_ef, df_im = cargar_datos_drive()
except Exception as e:
    st.error(f"Error crítico: {e}"); st.stop()

# =========================================================================
# 5. FILTROS MAESTROS
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h1, h2, h3 = st.columns([0.8, 3.5, 0.7])
    with h1: 
        try: st.image("LOGO OMBÚ.jpg", width=90)
        except: st.markdown("##### OMBÚ")
    with h2: st.markdown("<h3 style='margin:0; color:white;'>TABLERO INTEGRADO C.G.P.</h3>", unsafe_allow_html=True)
    with h3:
        if st.button("🚪 Salir", use_container_width=True): 
            st.session_state['autenticado'] = False; st.rerun()

with st.sidebar:
    st.header("🎛️ FILTROS")
    lista_meses = df_ef['Mes_Str'].dropna().unique().tolist() if 'Mes_Str' in df_ef.columns else []
    sel_mes = st.multiselect("📅 Mes", lista_meses, placeholder="Seleccionar Mes")
    
    lista_plantas = sorted(df_ef['Planta'].dropna().unique()) if 'Planta' in df_ef.columns else []
    sel_planta = st.multiselect("🏭 Planta", lista_plantas)
    
    lista_lineas = sorted(df_ef['Linea'].dropna().unique()) if 'Linea' in df_ef.columns else []
    sel_linea = st.multiselect("⚙️ Línea", lista_lineas)
    
    lista_puestos = sorted(df_ef['Puesto_Trabajo'].dropna().unique()) if 'Puesto_Trabajo' in df_ef.columns else []
    sel_puesto = st.multiselect("🛠️ Puesto", lista_puestos)

# APLICACIÓN CRUZADA
df_ef_f = df_ef.copy()
df_im_f = df_im.copy()

if sel_mes and 'Mes_Str' in df_ef_f.columns: 
    df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(sel_mes)]
if sel_planta: 
    if 'Planta' in df_ef_f.columns: df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
    if 'PLANTA' in df_im_f.columns: df_im_f = df_im_f[df_im_f['PLANTA'].astype(str).apply(lambda x: safe_match(sel_planta, x))]
if sel_linea and 'Linea' in df_ef_f.columns: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto and 'Puesto_Trabajo' in df_ef_f.columns: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]

warn_linea = False
if sel_puesto: 
    df_plot_1 = df_ef_f.copy()
elif sel_linea:
    df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_salida.empty: df_plot_1 = df_salida
    else: df_plot_1 = df_ef_f.copy(); warn_linea = True
else: 
    df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_salida.empty: df_plot_1 = df_salida
    else: df_plot_1 = df_ef_f.copy()

# =========================================================================
# 6. CÁLCULOS OEE (CON LÓGICA ESTRICTA 'ÚLTIMO PUESTO' Y COLUMNA C/GAP)
# =========================================================================
def safe_sum(ser): 
    try:
        val = ser.sum()
        if isinstance(val, (pd.Series, np.ndarray)): return float(val.iloc[0])
        return float(val)
    except: return 0.0

# VARIABLES GLOBALES PARA CARTELES
tot_costo = safe_sum(df_ef_f['Costo_Improd._$'])
tot_hh_imp_global = safe_sum(df_im_f['HH_IMPRODUCTIVAS'])

# VARIABLES ESPECÍFICAS PARA EL OEE (SOLO ÚLTIMO PUESTO)
t_std_oee = safe_sum(df_plot_1['HH_STD_TOTAL'])
t_disp_oee = safe_sum(df_plot_1['HH_Disponibles'])

# EL RENDIMIENTO TOMA ESTRICTAMENTE LA COLUMNA HH_Productivas_C/GAP DE ÚLTIMO PUESTO
c_prod_gap = next((c for c in df_plot_1.columns if 'GAP' in str(c).upper() and 'PROD' in str(c).upper()), 'HH_Productivas_C/GAP')
t_prod_oee = safe_sum(df_plot_1[c_prod_gap]) if c_prod_gap in df_plot_1.columns else 0.0

# LA DISPONIBILIDAD TOMA ESTRICTAMENTE LAS HH_IMPRODUCTIVAS DE LOS PUESTOS CONSIDERADOS "ÚLTIMOS"
if not df_im_f.empty and 'Puesto_Trabajo' in df_plot_1.columns and 'PUESTO' in df_im_f.columns:
    ultimos_puestos = df_plot_1['Puesto_Trabajo'].unique()
    t_imp_oee = df_im_f[df_im_f['PUESTO'].isin(ultimos_puestos)]['HH_IMPRODUCTIVAS'].sum()
else:
    t_imp_oee = tot_hh_imp_global

# Fórmulas Finales Exactas
disponibilidad = ((t_disp_oee - t_imp_oee) / t_disp_oee * 100) if t_disp_oee > 0 else 0.0
rendimiento = (t_std_oee / t_prod_oee * 100) if t_prod_oee > 0 else 0.0
calidad_sim = 98.0
oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Sincronizado
if oee_final >= 85: color_main = "#22c55e" # Verde
elif oee_final >= 65: color_main = "#eab308" # Amarillo
else: color_main = "#ef4444" # Rojo

inyectar_estilos_integrados(color_main)

# =========================================================================
# 7. PORTADA: VELOCÍMETRO OEE PRO CON EXPLICACIONES
# =========================================================================
st.markdown("<div class='card-oee'>", unsafe_allow_html=True)
fig_oee = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "OEE GLOBAL (%)", 'font': {'color': 'white', 'size': 16}},
    number = {'suffix': "%", 'font': {'color': color_main, 'size': 75}},
    gauge = {
        'axis': {'range': [None, 100], 'tickcolor': "white"},
        'bar': {'color': "white"}, 'bgcolor': "rgba(0,0,0,0)",
        'steps': [{'range': [0, 65], 'color': '#ef4444'},{'range': [65, 85], 'color': '#eab308'},{'range': [85, 100], 'color': '#22c55e'}],
        'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': 85}
    }
))
fig_oee.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=320, margin=dict(l=30, r=30, t=50, b=0))
st.plotly_chart(fig_oee, use_container_width=True)

st.markdown(f"""
    <div style='display: flex; gap: 15px;'>
        <div class='pillar-item'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <h2 style='color: white; margin:0;'>{min(disponibilidad, 100.0):.1f}%</h2>
            <p class='p-formula'>Fórmula: (HH Disp - HH Imp) / HH Disp</p>
            <p class='p-desc'>Mide el tiempo real de producción de la planta frente al tiempo que estuvo disponible, penalizado por las horas improductivas reportadas en la última estación.</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>🚀 Rendimiento</p>
            <h2 style='color: white; margin:0;'>{min(rendimiento, 100.0):.1f}%</h2>
            <p class='p-formula'>Fórmula: HH Standard / HH Productivas_C_GAP</p>
            <p class='p-desc'>Evalúa la velocidad de producción real comparada con el tiempo estándar establecido en las últimas estaciones.</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>💎 Calidad</p>
            <h2 style='color: white; margin:0;'>{calidad_sim:.1f}%</h2>
            <p class='p-formula'>Fórmula: Piezas OK / Total <span style='color:#ef4444'>(SIM)</span></p>
            <p class='p-desc'>Representa el porcentaje de piezas producidas correctamente sin necesidad de descartes ni reprocesos.</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 8. KPIs ORIGINALES
# =========================================================================
kpi_ef_real = (t_std_oee / t_disp_oee * 100) if t_disp_oee > 0 else 0
kpi_ef_prod = (t_std_oee / t_prod_oee * 100) if t_prod_oee > 0 else 0

st.markdown(f"""
<div class="kpi-grid">
    <div style="background: linear-gradient(135deg, #e0e0e0, #f5f5f5); border: 1px solid #aaa; border-left: 6px solid {color_main}; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
        <h4 style="margin:0; color: {color_main}; font-size:16px;">EFICIENCIA REAL</h4>
        <h2 style="margin:5px 0 0 0; color: #111; font-size:42px;">{kpi_ef_real:.1f}%</h2>
    </div>
    <div style="background: linear-gradient(135deg, #2E7D32, #4CAF50); border: 1px solid #1B5E20; border-left: 6px solid #A5D6A7; padding: 15px; border-radius: 6px; text-align:center; box-shadow: 2px 4px 10px rgba(0,0,0,0.3);">
        <h4 style="margin:0; color: white; font-size:16px;">EFICIENCIA PROD.</h4>
        <h2 style="margin:5px 0 0 0; color: white; font-size:42px;">{kpi_ef_prod:.1f}%</h2>
    </div>
    <div class="kpi-costo" style="background: linear-gradient(135deg, #D32F2F, #E53935); border: 1px solid #B71C1C; padding: 15px; border-radius: 8px; display: flex; flex-direction: column; justify-content: center; text-align:center; box-shadow: 2px 4px 15px rgba(211,47,47,0.4);">
        <h4 style="margin:0; color: white; font-size:22px;">COSTO HH IMPROD.</h4>
        <p style="margin:0; color: #FFCDD2; font-size:14px;">(Oportunidad Perdida Global)</p>
        <h2 style="margin:10px 0; color: #FFEB3B; font-size:48px; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${tot_costo:,.0f}</h2>
        <h4 style="margin:0; color: white; font-size:22px;">{tot_hh_imp_global:,.1f} <span style="font-size:16px; font-weight:normal;">HH</span></h4>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 9. GRÁFICOS M1 A M6 INTACTOS
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
col1, col2 = st.columns(2)
t_enc = f"Filtros >> Planta: {'+'.join(sel_planta) if sel_planta else 'Todas'} | Línea: {'+'.join(sel_linea) if sel_linea else 'Todas'} | Puesto: {'+'.join(sel_puesto) if sel_puesto else 'Todos'}"

with col1:
    st.markdown(f"<h3 style='color:white; border-left: 5px solid {color_main}; padding-left:10px;'>1. TENDENCIA EFICIENCIA REAL</h3>", unsafe_allow_html=True)
    if warn_linea: st.warning("⚠️ Esta Línea NO registra un 'Último Puesto'. Seleccione un Puesto para análisis preciso.")
    
    if not df_plot_1.empty and 'Mes_Str' in df_plot_1.columns:
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
        
        ax1_line.plot(x_idx, ag1['Ef_Real'], color=color_main, marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real', zorder=5)
        add_tendencia(ax1_line, x_idx, ag1['Ef_Real'], color_main)
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        
        last_x1 = x_idx[-1] if len(x_idx) > 0 else 0
        ax1_line.text(last_x1, 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10, ha='right', va='bottom')
        
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3))
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        
        for i, val in enumerate(ag1['Ef_Real']): 
            ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
            
        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold', color='white')
        ax1.tick_params(axis='y', colors='white'); ax1_line.tick_params(axis='y', colors='white')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.05), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.05), frameon=True)
        agregar_sello_agua(fig1); st.pyplot(fig1, use_container_width=True)
    else: st.warning("⚠️ Sin datos para Eficiencia Real. Pruebe otra combinación de filtros.")

with col2:
    st.markdown(f"<h3 style='color:white; border-left: 5px solid #ef4444; padding-left:10px;'>2. PARETO DE CAUSAS</h3>", unsafe_allow_html=True)
    if not df_im_f.empty and 'TIPO_PARADA' in df_im_f.columns:
        res_p = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(14, 10), facecolor='#0f172a')
        res_p.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis(); ax2.tick_params(colors='white'); ax2.set_xlabel("Horas Perdidas", color='white', fontweight='bold')
        st.pyplot(fig2)

st.markdown("---")

col3, col4 = st.columns(2)
with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='font-size:14px; color:#aaa; margin-top:-15px; margin-bottom:10px;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH_Productivas_C/GAP'
        if c_prod not in df_ef_f.columns: c_prod = 'HH_Productivas_C/GAP'
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
        add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'], 'darkgray')
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
# 10. TABLA DE DETALLES
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
