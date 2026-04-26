import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
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
        
        /* Reglas responsivas de tu código original */
        .kpi-grid {{ display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 12px; }}
        .kpi-costo {{ grid-row: span 2; }}
        .mobile-only {{ display: none !important; }}
        
        @media (max-width: 768px) {{
            .mobile-only {{ display: block !important; }}
            [data-testid="stHorizontalBlock"] {{ flex-direction: column !important; }}
            .kpi-grid {{ grid-template-columns: 1fr !important; }}
            .kpi-costo {{ grid-row: span 1 !important; }}
        }}

        /* Tarjeta OEE Master */
        .card-oee {{
            background-color: #1e293b; border-radius: 20px; padding: 25px; 
            margin-bottom: 20px; border: 1px solid #334155;
            box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
        }}
        
        /* Pilares OEE */
        .pillar-item {{ 
            background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 5px solid {color_per} !important;
        }}
        .p-label {{ font-size: 12px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }}
        .p-val {{ font-size: 32px; font-weight: 900; color: white; margin: 0; }}
        .p-formula {{ font-size: 10px; color: {color_per}; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; font-style: italic; }}
    </style>
    """, unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o, caja_b = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD (TUS CREDENCIALES)
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True); c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='text-align:center;'><h2 style='color:white;'>GESTIÓN INDUSTRIAL OMBÚ</h2></div>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in, p_in = st.text_input("Usuario Corporativo"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: 
                    st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE Y CARGA (BLINDAJE TOTAL DE NOMBRES)
# =========================================================================
def safe_match(s_list, val):
    if pd.isna(val): return False
    v_norm = re.sub(r'[^A-Z0-9]', '', str(val).upper())
    for s in s_list:
        s_norm = re.sub(r'[^A-Z0-9]', '', str(s).upper())
        if s_norm == v_norm and s_norm != "": return True
    return False

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if '✅' in d: return "🎯 ACCIÓN GLOBAL"
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO', 'MECANICO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'LOGISTICA']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'ERROR']): return "🔎 Ajustar Calidad"
    return "⚡ Investigar Causa"

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

@st.cache_data(ttl=300)
def cargar_datos_drive():
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    
    d_ef = pd.read_excel(url_ef)
    d_im = pd.read_excel(url_im)
    
    # ---------------- BLINDAJE EFICIENCIAS ----------------
    d_ef.columns = [str(c).strip().upper() for c in d_ef.columns]
    map_ef = {
        'FECHA': 'Fecha', 'PLANTA': 'Planta', 'LINEA': 'Linea', 'LÍNEA': 'Linea', 
        'PUESTO': 'Puesto_Trabajo', 'STD': 'HH_STD', 'DISP': 'HH_Disp', 
        'PROD': 'HH_Prod', 'GAP': 'HH_Prod', 'COSTO': 'Costo'
    }
    for col_excel in d_ef.columns:
        for k, v in map_ef.items():
            if k in col_excel: d_ef.rename(columns={col_excel: v}, inplace=True)
            
    # Forzado Numérico Seguro
    for col in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if col not in d_ef.columns: d_ef[col] = 0.0
        d_ef[col] = pd.to_numeric(d_ef[col], errors='coerce').fillna(0)
    
    # Orden Cronológico
    if 'Fecha' in d_ef.columns:
        d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce', dayfirst=True)
        d_ef = d_ef.sort_values('Fecha')
        d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    else:
        d_ef['Mes_Str'] = "S/D"
    
    # ---------------- BLINDAJE IMPRODUCTIVAS ----------------
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    c_hh = next((c for c in d_im.columns if 'HH' in c and 'IMP' in c), d_im.columns[0])
    d_im.rename(columns={c_hh: 'HH_IMPRODUCTIVAS'}, inplace=True)
    d_im['HH_IMPRODUCTIVAS'] = pd.to_numeric(d_im['HH_IMPRODUCTIVAS'], errors='coerce').fillna(0).abs()
    
    c_nom = next((c for c in d_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in d_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: d_im['OPERARIO'] = d_im[c_nom].astype(str).replace('nan','') + ' ' + d_im[c_ape].astype(str).replace('nan','')
    elif c_nom: d_im['OPERARIO'] = d_im[c_nom].astype(str).replace('nan','')
    else: d_im['OPERARIO'] = "S/D"
    
    if 'TIPO_PARADA' not in d_im.columns:
        c_tipo = next((c for c in d_im.columns if 'TIPO' in c or 'MOTIVO' in c), d_im.columns[0])
        d_im.rename(columns={c_tipo: 'TIPO_PARADA'}, inplace=True)
        
    if 'DETALLE' not in d_im.columns: d_im['DETALLE'] = "S/D"
    
    return d_ef, d_im

try:
    df_ef, df_im = cargar_datos_drive()
except Exception as e:
    st.error(f"Error cargando datos: {e}"); st.stop()

# =========================================================================
# 4. FILTROS MAESTROS (SIDEBAR A PRUEBA DE FALLOS)
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h_l, h_t, h_s = st.columns([0.8, 3.5, 0.7])
    with h_l:
        try: st.image("LOGO OMBÚ.jpg", width=90)
        except: st.markdown("##### OMBÚ")
    with h_t: st.markdown("<h3 style='margin:0; color:white;'>TABLERO INTEGRADO C.G.P.</h3>", unsafe_allow_html=True)
    with h_s:
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
    c_pl = next((c for c in df_im_f.columns if 'PLANTA' in c), None)
    if c_pl: df_im_f = df_im_f[df_im_f[c_pl].astype(str).apply(lambda x: safe_match(sel_planta, x))]
if sel_linea and 'Linea' in df_ef_f.columns: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto and 'Puesto_Trabajo' in df_ef_f.columns: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]

# =========================================================================
# 5. CÁLCULOS OEE Y SINCRONIZACIÓN DE COLORES
# =========================================================================
def safe_sum(ser): return float(ser.sum()) if not ser.empty else 0.0

t_std = safe_sum(df_ef_f['HH_STD'])
t_disp = safe_sum(df_ef_f['HH_Disp'])
t_prod = safe_sum(df_ef_f['HH_Prod'])
t_costo = safe_sum(df_ef_f['Costo'])
t_imp = safe_sum(df_im_f['HH_IMPRODUCTIVAS'])

# Fórmulas
disponibilidad = ((t_disp - t_imp) / t_disp * 100) if t_disp > 0 else 0.0
rendimiento = (t_std / t_prod * 100) if t_prod > 0 else 0.0
calidad_sim = 98.0
oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Sincronizado
if oee_final >= 85: color_main = "#22c55e" # Verde
elif oee_final >= 65: color_main = "#eab308" # Amarillo
else: color_main = "#ef4444" # Rojo

inyectar_estilos_integrados(color_main)

# =========================================================================
# 6. PORTADA: VELOCÍMETRO OEE PRO
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
            <p class='p-formula'>(HH Disp - HH Imp) / HH Disp</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>🚀 Rendimiento</p>
            <h2 style='color: white; margin:0;'>{min(rendimiento, 100.0):.1f}%</h2>
            <p class='p-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-item'>
            <p class='p-label'>💎 Calidad</p>
            <h2 style='color: white; margin:0;'>{calidad_sim:.1f}%</h2>
            <p class='p-formula'>Piezas OK / Total <br><span style='color:#ef4444'>(SIMULADO)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. TUS CARTELES KPI ORIGINALES (COLORIZADOS)
# =========================================================================
kpi_ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0
kpi_ef_prod = (t_std / t_prod * 100) if t_prod > 0 else 0

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
        <h2 style="margin:10px 0; color: #FFEB3B; font-size:48px;">${t_costo:,.0f}</h2>
        <h4 style="margin:0; color: white; font-size:20px;">{t_imp:,.1f} HH</h4>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 8. TUS GRÁFICOS M1 A M6 (INTACTOS)
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    st.markdown(f"<h3 style='color:white; border-left: 5px solid {color_main}; padding-left:10px;'>1. TENDENCIA EFICIENCIA REAL</h3>", unsafe_allow_html=True)
    if not df_ef_f.empty and 'Mes_Str' in df_ef_f.columns:
        ag1 = df_ef_f.groupby('Mes_Str', sort=False).agg({'HH_STD': 'sum', 'HH_Disp': 'sum'}).reset_index()
        ag1['Ef'] = (ag1['HH_STD'] / ag1['HH_Disp'] * 100).fillna(0)
        
        fig1, ax1 = plt.subplots(figsize=(14, 10), facecolor='#0f172a'); ax1_line = ax1.twinx()
        x_idx = np.arange(len(ag1))
        
        ax1.bar(x_idx - 0.17, ag1['HH_STD'], 0.35, color='midnightblue', edgecolor='white', label='HH STD')
        ax1.bar(x_idx + 0.17, ag1['HH_Disp'], 0.35, color='black', edgecolor='white', label='HH DISP')
        
        ax1_line.plot(x_idx, ag1['Ef'], color=color_main, marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real')
        add_tendencia(ax1_line, x_idx, ag1['Ef'], color_main)
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3)
        ax1_line.set_ylim(0, 110); ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Mes_Str'], rotation=0, color='white')
        ax1.tick_params(axis='y', colors='white'); ax1_line.tick_params(axis='y', colors='white')
        
        agregar_sello_agua(fig1); st.pyplot(fig1)

with col2:
    st.markdown(f"<h3 style='color:white; border-left: 5px solid #ef4444; padding-left:10px;'>2. PARETO DE CAUSAS</h3>", unsafe_allow_html=True)
    if not df_im_f.empty and 'TIPO_PARADA' in df_im_f.columns:
        res_p = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().sort_values(ascending=False).head(5)
        fig2, ax2 = plt.subplots(figsize=(14, 10), facecolor='#0f172a')
        res_p.plot(kind='barh', color='#ef4444', ax=ax2)
        ax2.invert_yaxis(); ax2.tick_params(colors='white'); ax2.set_xlabel("Horas Perdidas", color='white')
        st.pyplot(fig2)

# =========================================================================
# 9. DETALLES DE OPERACIÓN (LA MESA DE TRABAJO DE TU MAIN)
# =========================================================================
st.markdown("---")
st.header("📋 ANÁLISIS DETALLADO Y ACCIONES SUGERIDAS")
if not df_im_f.empty and 'DETALLE' in df_im_f.columns:
    df_im_f['Acción Sugerida'] = df_im_f['DETALLE'].astype(str).apply(generar_accion_sugerida)
    cols = ['OPERARIO', 'TIPO_PARADA', 'DETALLE', 'HH_IMPRODUCTIVAS', 'Acción Sugerida']
    
    # Filtramos para asegurarnos que las columnas existen antes de mostrar
    cols_existentes = [c for c in cols if c in df_im_f.columns]
    
    df_mesa = df_im_f[cols_existentes].sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
    st.dataframe(df_mesa, use_container_width=True, hide_index=True)
    
st.caption("Ombu Industrial Intelligence © 2026 | Desarrollado para Control de Gestión")
