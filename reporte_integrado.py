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
# 1. CONFIGURACIÓN Y ESTILO SINCRONIZADO (CSS)
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")

def inject_theme_cgp(color_per):
    st.markdown(f"""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"], footer {{display: none !important;}}
        .block-container {{padding-top: 1rem !important; background-color: #0f172a;}}
        
        /* Sticky Header */
        div[data-testid="stVerticalBlock"] > div:has(#sticky-header) {{
            position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
            background-color: rgba(15, 23, 42, 0.98) !important; z-index: 99999 !important;
            padding: 5px 10px 15px 10px !important; border-bottom: 2px solid {color_per} !important;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.5);
        }}

        /* Tarjeta OEE Master */
        .card-oee {{
            background-color: #1e293b; border-radius: 20px; padding: 25px; 
            margin-bottom: 20px; border: 1px solid #334155;
            box-shadow: 0px 15px 25px rgba(0,0,0,0.2);
        }}

        /* Pilares Sincronizados */
        .pillar-box {{ 
            background: #0f172a; border-radius: 15px; padding: 20px; flex: 1;
            border: 1px solid #334155; text-align: center;
            border-top: 5px solid {color_per} !important;
        }}
        .p-label {{ font-size: 12px; font-weight: 700; color: #94a3b8; text-transform: uppercase; }}
        .p-val {{ font-size: 32px; font-weight: 900; color: white; margin: 0; }}
        .p-formula {{ font-size: 10px; color: {color_per}; margin-top: 10px; border-top: 1px solid #334155; padding-top: 8px; font-style: italic; }}

        /* KPIs de Soporte (Originales del Main) */
        .kpi-integrated {{
            background: #1e293b; padding: 20px; border-radius: 15px;
            text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-top: 4px solid {color_per};
        }}
        .kpi-integrated h2 {{ margin: 5px 0; font-size: 35px; color: white; }}
        .kpi-integrated p {{ margin: 0; font-size: 12px; color: #94a3b8; font-weight: bold; text-transform: uppercase; }}

        /* Responsive */
        @media (max-width: 768px) {{
            [data-testid="stHorizontalBlock"] {{ flex-direction: column !important; }}
            .pillar-box {{ margin-bottom: 10px; }}
        }}
    </style>
    """, unsafe_allow_html=True)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
if 'auth' not in st.session_state: st.session_state['auth'] = False
if not st.session_state['auth']:
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='text-align:center; margin-top:100px;'><h1 style='color:white;'>🏭 SISTEMA OMBÚ</h1><p style='color:#94a3b8;'>Acceso Control de Gestión</p></div>", unsafe_allow_html=True)
        with st.form("login"):
            u, p = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("INGRESAR AL TABLERO", use_container_width=True):
                if u == "acceso.ombu" and p == "Gestion2026":
                    st.session_state['auth'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
    st.stop()

# =========================================================================
# 3. MOTOR DE DATOS Y LÓGICA DE NEGOCIO (THE BEAST)
# =========================================================================
def safe_val(series):
    try:
        res = series.sum()
        return float(res.iloc[0]) if hasattr(res, 'iloc') else float(res)
    except: return 0.0

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if any(x in d for x in ['ROTURA', 'FALLA', 'MANTENIMIENTO', 'MECANICO']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'LOGISTICA', 'ABASTECIMIENTO']): return "📦 Apurar Logística"
    if any(x in d for x in ['CALIDAD', 'ERROR', 'REPROCESO', 'DEFECTO']): return "🔎 Ajustar Calidad"
    if any(x in d for x in ['LIMPIEZA', 'ORDEN', '5S']): return "🧹 Optimizar 5S"
    return "⚡ Investigar Causa"

@st.cache_data(ttl=300)
def load_all_data():
    url_ef = "https://drive.google.com/uc?export=download&id=14kmjYqzkgRs0V2pFGMaEc6ebZc9tcK_V"
    url_im = "https://drive.google.com/uc?export=download&id=1LdemtoOSyetVgXCxDrYsL7tNUZKqiK9P"
    d_ef, d_im = pd.read_excel(url_ef), pd.read_excel(url_im)
    
    # --- LIMPIEZA EFICIENCIAS ---
    d_ef.columns = [str(c).strip().upper() for c in d_ef.columns]
    map_ef = {'FECHA':'Fecha', 'PLANTA':'Planta', 'LINEA':'Linea', 'STD':'HH_STD', 'DISP':'HH_Disp', 'PROD':'HH_Prod', 'COSTO':'Costo', 'ULTIMO':'Es_Ultimo', 'PUESTO':'Puesto'}
    for c in d_ef.columns:
        for k, v in map_ef.items():
            if k in c: d_ef.rename(columns={c: v}, inplace=True)
    
    # Orden Cronológico
    d_ef['Fecha'] = pd.to_datetime(d_ef['Fecha'], errors='coerce')
    d_ef = d_ef.sort_values('Fecha')
    d_ef['Mes_Str'] = d_ef['Fecha'].dt.strftime('%b-%Y')
    
    for c in ['HH_STD', 'HH_Disp', 'HH_Prod', 'Costo']:
        if c in d_ef.columns: d_ef[c] = pd.to_numeric(d_ef[c], errors='coerce').fillna(0)

    # --- LIMPIEZA IMPRODUCTIVAS ---
    d_im.columns = [str(c).strip().upper() for c in d_im.columns]
    map_im = {'HH':'HH_Imp', 'TIPO':'Motivo', 'MOTIVO':'Motivo', 'DETALLE':'Detalle', 'PLANTA':'Planta', 'LINEA':'Linea', 'PUESTO':'Puesto'}
    for c in d_im.columns:
        for k, v in map_im.items():
            if k in c: d_im.rename(columns={c: v}, inplace=True)
    
    d_im['HH_Imp'] = pd.to_numeric(d_im['HH_Imp'], errors='coerce').fillna(0).abs()
    
    # Operarios
    c_nom = next((c for c in d_im.columns if 'NOMBRE' in c), None)
    c_ape = next((c for c in d_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: d_im['OPERARIO'] = d_im[c_nom].astype(str).replace('nan','') + ' ' + d_im[c_ape].astype(str).replace('nan','')
    else: d_im['OPERARIO'] = "S/D"
    
    return d_ef, d_im

try:
    df_ef, df_im = load_all_data()
except Exception as e:
    st.error(f"Error cargando datos: {e}"); st.stop()

# =========================================================================
# 4. HEADER Y FILTROS MAESTROS (CASCADA)
# =========================================================================
with st.container():
    st.markdown('<div id="sticky-header"></div>', unsafe_allow_html=True)
    h_col1, h_col2 = st.columns([4, 1])
    with h_col1:
        st.markdown("<h2 style='color:white; margin:0;'>REPORTING INDUSTRIAL CGP - OMBÚ</h2>", unsafe_allow_html=True)
    with h_col2:
        if st.button("🚪 Salir", use_container_width=True): 
            st.session_state['auth'] = False; st.rerun()

    # Filtros en el Sidebar
    with st.sidebar:
        st.header("⚙️ FILTROS")
        # Meses Cronológicos
        lista_meses = df_ef['Mes_Str'].unique().tolist()
        sel_mes = st.multiselect("📅 Mes", lista_meses, placeholder="Seleccionar Mes")
        
        sel_planta = st.multiselect("🏭 Planta", sorted(df_ef['Planta'].dropna().unique()))
        sel_linea = st.multiselect("⚙️ Línea", sorted(df_ef['Linea'].dropna().unique()))
        sel_puesto = st.multiselect("🛠️ Puesto", sorted(df_ef['Puesto'].dropna().unique()))

# Lógica de Filtrado Cruzado
df_ef_f = df_ef.copy()
df_im_f = df_im.copy()

if sel_mes: 
    df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(sel_mes)]
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
    df_im_f = df_im_f[df_im_f['Planta'].astype(str).isin(sel_planta)]
if sel_linea:
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
    df_im_f = df_im_f[df_im_f['Linea'].astype(str).isin(sel_linea)]
if sel_puesto:
    df_ef_f = df_ef_f[df_ef_f['Puesto'].isin(sel_puesto)]
    df_im_f = df_im_f[df_im_f['Puesto'].astype(str).isin(sel_puesto)]

# =========================================================================
# 5. CÁLCULOS OEE (FORMULACIÓN OMBÚ)
# =========================================================================
t_std = safe_val(df_ef_f['HH_STD'])
t_disp = safe_val(df_ef_f['HH_Disp'])
t_prod = safe_val(df_ef_f['HH_Prod'])
t_costo = safe_val(df_ef_f['Costo'])
t_imp = safe_val(df_im_f['HH_Imp'])

# 1. DISPONIBILIDAD = (HH Disponibles - HH Improductivas) / HH Disponibles
disponibilidad = ((t_disp - t_imp) / t_disp * 100) if t_disp > 0 else 0.0

# 2. RENDIMIENTO = Eficiencia Productiva (HH Standard / HH Productivas)
rendimiento = (t_std / t_prod * 100) if t_prod > 0 else 0.0

# 3. CALIDAD = 98.0% (Simulado hasta carga real)
calidad_sim = 98.0

oee_final = (max(0, disponibilidad)/100 * max(0, rendimiento)/100 * calidad_sim/100) * 100

# Color Maestro Sincronizado
color_status = "#22c55e" if oee_final >= 85 else "#eab308" if oee_final >= 65 else "#ef4444"
inject_theme_cgp(color_status)

# =========================================================================
# 6. VELOCÍMETRO OEE PRO
# =========================================================================
st.markdown("<div class='card-oee'>", unsafe_allow_html=True)

fig_oee = go.Figure(go.Indicator(
    mode = "gauge+number", value = oee_final,
    title = {'text': "INDICADOR OEE GLOBAL", 'font': {'color': 'white', 'size': 16}},
    number = {'suffix': "%", 'font': {'color': color_status, 'size': 80}},
    gauge = {
        'axis': {'range': [None, 100], 'tickcolor': "white"},
        'bar': {'color': "white"},
        'bgcolor': "rgba(0,0,0,0)",
        'steps': [
            {'range': [0, 65], 'color': '#ef4444'},
            {'range': [65, 85], 'color': '#eab308'},
            {'range': [85, 100], 'color': '#22c55e'}
        ],
        'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': 85}
    }
))
fig_oee.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=380, margin=dict(l=30, r=30, t=50, b=0))
st.plotly_chart(fig_oee, use_container_width=True)

# Pilares OEE con Fórmulas
st.markdown(f"""
    <div style='display: flex; gap: 15px; margin-top: 10px;'>
        <div class='pillar-box'>
            <p class='p-label'>⏱️ Disponibilidad</p>
            <p class='p-val'>{min(disponibilidad, 100.0):.1f}%</p>
            <p class='p-formula'>(HH Disp - HH Imp) / HH Disp</p>
        </div>
        <div class='pillar-box'>
            <p class='p-label'>🚀 Rendimiento</p>
            <p class='p-val'>{min(rendimiento, 100.0):.1f}%</p>
            <p class='p-formula'>HH Standard / HH Productivas</p>
        </div>
        <div class='pillar-box'>
            <p class='p-label'>💎 Calidad</p>
            <p class='p-val'>{calidad_sim:.1f}%</p>
            <p class='p-formula'>Piezas OK / Total Producido <br><span style='color:#ef4444'>(SIMULADO)</span></p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 7. MÉTRICAS DE SOPORTE (INTEGRACIÓN MAIN)
# =========================================================================
st.markdown("<h4 style='color:white; margin-bottom:15px; font-size:18px;'>📊 MÉTRICAS DE GESTIÓN Y COSTOS</h4>", unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)

with c1:
    ef_real = (t_std / t_disp * 100) if t_disp > 0 else 0.0
    st.markdown(f"<div class='kpi-integrated'><p>Eficiencia Real</p><h2 style='color:{color_status};'>{ef_real:.1f}%</h2></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='kpi-integrated'><p>HH Standard Producidas</p><h2>{t_std:,.0f}</h2></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='kpi-integrated' style='border-top-color:#ef4444;'><p>Costo Ineficiencia</p><h2 style='color:#ef4444;'>${t_costo:,.0f}</h2></div>", unsafe_allow_html=True)
with c4:
    gap_h = t_disp - t_prod
    st.markdown(f"<div class='kpi-integrated' style='border-top-color:#3b82f6;'><p>Gap de Horas (Sin Cargar)</p><h2 style='color:#3b82f6;'>{int(gap_h)}h</h2></div>", unsafe_allow_html=True)

# =========================================================================
# 8. GRÁFICOS Y PARETO
# =========================================================================
st.markdown("<br>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns([2, 1])

with g_col1:
    st.markdown("<h4 style='color:white;'>📈 Evolución Cronológica de Eficiencia Real</h4>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        fig_trend, ax_trend = plt.subplots(figsize=(10, 4.5), facecolor='#0f172a')
        ag_trend = df_ef_f.groupby('Mes_Str', sort=False).agg({'HH_STD':'sum', 'HH_Disp':'sum'})
        ag_trend['Ef'] = (ag_trend['HH_STD'] / ag_trend['HH_Disp'] * 100)
        ax_trend.plot(ag_trend.index, ag_trend['Ef'], marker='o', color=color_status, linewidth=4, markersize=10, label='% Eficiencia Real')
        ax_trend.fill_between(ag_trend.index, ag_trend['Ef'], color=color_status, alpha=0.1)
        ax_trend.axhline(85, color='#22c55e', linestyle='--', alpha=0.6, label="Meta (85%)")
        ax_trend.set_ylim(0, 110); ax_trend.tick_params(colors='white'); ax_trend.grid(axis='y', alpha=0.1)
        ax_trend.legend(facecolor='#1e293b', labelcolor='white')
        st.pyplot(fig_trend)

with g_col2:
    st.markdown("<h4 style='color:white;'>⚠️ Top Causas Improductivas</h4>", unsafe_allow_html=True)
    if not df_im_f.empty:
        res_pareto = df_im_f.groupby('Motivo')['HH_Imp'].sum().sort_values(ascending=False).head(5)
        fig_p, ax_p = plt.subplots(figsize=(5, 8.5), facecolor='#0f172a')
        res_pareto.plot(kind='barh', color='#ef4444', ax=ax_p)
        ax_p.invert_yaxis(); ax_p.tick_params(colors='white'); ax_p.set_xlabel("HH Perdidas", color='white')
        st.pyplot(fig_p)

# =========================================================================
# 9. TABLA DE DETALLES (EL CORAZÓN DEL ANÁLISIS)
# =========================================================================
st.markdown("---")
st.header("📋 ANÁLISIS DE OPERACIÓN Y ACCIONES SUGERIDAS")

if not df_im_f.empty:
    df_im_f['Acción Sugerida'] = df_im_f['Detalle'].apply(generar_accion_sugerida)
    # Formateo para visualización
    cols_display = ['OPERARIO', 'Motivo', 'Detalle', 'HH_Imp', 'Acción Sugerida']
    df_mesa = df_im_f[cols_display].sort_values(by='HH_Imp', ascending=False)
    
    # Tabla interactiva con búsqueda
    st.dataframe(df_mesa, use_container_width=True, hide_index=True)
    
    # Descarga directa
    st.download_button(
        label="📥 Descargar Reporte de Operación (CSV)",
        data=df_mesa.to_csv(index=False).encode('utf-8'),
        file_name="Reporte_Industrial_Ombu.csv",
        mime="text/csv",
        use_container_width=True
    )
else:
    st.info("No hay registros de improductividad para los filtros seleccionados.")

st.caption("Ombu Industrial Intelligence © 2026 | Desarrollado para Control de Gestión de Planta")
