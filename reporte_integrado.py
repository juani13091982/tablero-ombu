import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as path_effects
import textwrap
import datetime
import os
import re

# =========================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y MARCO CORPORATIVO
# =========================================================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================================
# 2. SISTEMA DE SEGURIDAD (ACCESO RESTRINGIDO)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    """Dibuja la pantalla de acceso con logo pequeño y nombre oficial solicitado."""
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_v1, col_login, col_v2 = st.columns([1, 1.8, 1])
    
    with col_login:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        
        # Logo institucional pequeño centrado
        l_c1, l_c2, l_c3 = st.columns([1, 1, 1])
        with l_c2:
            try:
                st.image("LOGO OMBÚ.jpg", width=160)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_login_seguro"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            u_digitado = st.text_input("Usuario Corporativo")
            p_digitado = st.text_input("Contraseña", type="password")
            
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_digitado in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_digitado] == p_digitado:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos.")

if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES Y FILTROS STICKY (CSS)
# =========================================================================
st.markdown("""
<style>
    /* Ocultar elementos estándar para profesionalismo */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* PANEL DE FILTROS STICKY (FIJO ARRIBA) */
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 0px !important;
        background-color: #0E1117 !important; 
        z-index: 99999 !important;
        padding: 15px;
        border-bottom: 3px solid #1E3A8A !important; 
    }
</style>
""", unsafe_allow_html=True)

# Configuración Maestra de Matplotlib (Fuentes Grandes y Negrita)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Efectos de legibilidad (Contornos)
c_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
c_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cajas de anotación
bbox_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
bbox_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_oro = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. MOTOR INTELIGENTE DE CRUCE DE DATOS (MOTOR FUZZY)
# =========================================================================
def set_margen_superior(ax_plot, v_max, multiplicador=2.6):
    """Asegura que el gráfico tenga 'techo' para que las etiquetas no se pisen."""
    if v_max > 0: 
        ax_plot.set_ylim(0, v_max * multiplicador)
    else: 
        ax_plot.set_ylim(0, 100)

def dibujar_guias_meses(ax_plot, n_fechas):
    """Dibuja las líneas verticales grises para separar visualmente los meses."""
    for i in range(n_fechas):
        ax_plot.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def motor_robusto_cruce(sel_usuario, val_excel_impro):
    """
    MOTOR DE CRUCE INDUSTRIAL:
    Busca coincidencias entre Eficiencias e Improductivas aunque los textos varíen.
    Indispensable para recuperar las 54 HH perdidas en el filtrado.
    """
    if pd.isna(val_excel_impro) or pd.isna(sel_usuario): 
        return False
    
    # 1. Normalización profunda
    s1 = str(sel_usuario).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(val_excel_impro).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # 2. Solo alfanuméricos
    l1 = re.sub(r'[^A-Z0-9]', '', s1)
    l2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not l1 or not l2: 
        return False
        
    # 3. Coincidencia directa
    if l1 in l2 or l2 in l1: 
        return True
    
    # 4. Coincidencia por código de estación (3+ dígitos)
    n1 = set(re.findall(r'\d{3,}', s1))
    n2 = set(re.findall(r'\d{3,}', s2))
    if n1 and n2 and n1.intersection(n2): 
        return True
        
    # 5. Coincidencia por raíces de palabras
    w1 = set(re.findall(r'[A-Z]{4,}', s1))
    w2 = set(re.findall(r'[A-Z]{4,}', s2))
    
    excluir = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1_limpio = w1 - excluir
    v2_limpio = w2 - excluir
    
    for palabra in v1_limpio:
        if any(palabra in x for x in v2_limpio): 
            return True
                
    return False

# =========================================================================
# 5. CABECERA Y BOTÓN DE SALIDA
# =========================================================================
h_l, h_t, h_s = st.columns([1, 3, 1])

with h_l:
    try: st.image("LOGO OMBÚ.jpg", width=120)
    except: st.markdown("### OMBÚ")

with h_t:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Gerencia de Control de Gestión</p>", unsafe_allow_html=True)

with h_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA DE DATOS (EXCEL)
# =========================================================================
try:
    # Carga de archivos maestros
    df_ef_base = pd.read_excel("eficiencias.xlsx")
    df_im_base = pd.read_excel("improductivas.xlsx")
    
    # Limpieza inicial
    df_ef_base.columns = df_ef_base.columns.str.strip()
    df_im_base.columns = [str(c).strip().upper() for c in df_im_base.columns]
    
    # Auto-mapeo para improductivas
    if 'TIPO_PARADA' not in df_im_base.columns:
        c_motivo = next((c for c in df_im_base.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if c_motivo: df_im_base.rename(columns={c_motivo: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im_base.columns:
        c_hs = next((c for c in df_im_base.columns if 'HH' in c and 'IMP' in c), None)
        if c_hs: df_im_base.rename(columns={c_hs: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im_base.columns:
        c_fec = next((c for c in df_im_base.columns if 'FECHA' in c), None)
        if c_fec: df_im_base.rename(columns={c_fec: 'FECHA'}, inplace=True)
    
    # Estandarización temporal
    df_ef_base['Fecha'] = pd.to_datetime(df_ef_base['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im_base['FECHA'] = pd.to_datetime(df_im_base['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Atributos adicionales
    df_ef_base['Es_Ultimo_Puesto'] = df_ef_base['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef_base['Filtro_Mes'] = df_ef_base['Fecha'].dt.strftime('%b-%Y')
    df_im_base['Filtro_Mes'] = df_im_base['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e_carga:
    st.error(f"Error cargando los archivos Excel: {e_carga}")
    st.stop()

# =========================================================================
# 7. PANEL DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    fl1, fl2, fl3, fl4 = st.columns(4)
    
    with fl1: 
        p_list = list(df_ef_base['Planta'].dropna().unique())
        sel_planta = st.multiselect("🏭 Planta", p_list, placeholder="Todas")
        
    df_t_lineas = df_ef_base[df_ef_base['Planta'].isin(sel_planta)] if sel_planta else df_ef_base
    
    with fl2: 
        l_list = list(df_t_lineas['Linea'].dropna().unique())
        sel_linea = st.multiselect("⚙️ Línea", l_list, placeholder="Todas")
        
    df_t_puestos = df_t_lineas[df_t_lineas['Linea'].isin(sel_linea)] if sel_linea else df_t_lineas
    
    with fl3: 
        ps_list = list(df_t_puestos['Puesto_Trabajo'].dropna().unique())
        sel_puesto = st.multiselect("🛠️ Puesto de Trabajo", ps_list, placeholder="Todos")
        
    with fl4: 
        m_list = ["🎯 Acumulado YTD"] + list(df_ef_base['Filtro_Mes'].unique())
        sel_mes = st.multiselect("📅 Mes", m_list, placeholder="Todos")

# =========================================================================
# 8. LÓGICA DE FILTRADO FINAL (RECUPERACIÓN DE VARIABLES)
# =========================================================================
df_ef_f = df_ef_base.copy()
df_im_f = df_im_base.copy()

# Filtrar Eficiencias
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
if sel_linea: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]
if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_ef_f = df_ef_f[df_ef_f['Filtro_Mes'].isin(sel_mes)]

# Filtrar Improductivas (Motor Fuzzy)
if sel_planta:
    m_pl = df_im_f.iloc[:,0].apply(lambda x: any(motor_robusto_cruce(p, x) for p in sel_planta))
    df_im_f = df_im_f[m_pl]

if sel_linea:
    c_l_idx = next((c for c in df_im_f.columns if 'LINEA' in c), df_im_f.columns[1])
    m_li = df_im_f[c_l_idx].apply(lambda x: any(motor_robusto_cruce(l, x) for l in sel_linea))
    df_im_f = df_im_f[m_li]

if sel_puesto:
    c_p_idx = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2])
    m_ps = df_im_f[c_p_idx].apply(lambda x: any(motor_robusto_cruce(ps, x) for ps in sel_puesto))
    df_im_f = df_im_f[m_ps]

if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_im_f = df_im_f[df_im_f['Filtro_Mes'].isin(sel_mes)]

txt_header_graficos = f"Filtros: {sel_planta if sel_planta else 'Todas'} | {sel_linea if sel_linea else 'Todas'} | {sel_puesto if sel_puesto else 'Todos'}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (PRODUCTIVIDAD)
# =========================================================================
c1, c2 = st.columns(2)

with c1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    df_m1_build = df_ef_f.copy() if sel_puesto else df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_build.empty:
        agrup_1 = df_m1_build.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        agrup_1['Efic_Real'] = (agrup_1['HH_STD_TOTAL'] / agrup_1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig1, ax1_b = plt.subplots(figsize=(14, 10))
        ax1_l = ax1_b.twinx()
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_header_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices_1 = np.arange(len(agrup_1))
        
        b1 = ax1_b.bar(x_indices_1 - 0.17, agrup_1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b2 = ax1_b.bar(x_indices_1 + 0.17, agrup_1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_margen_superior(ax1_b, agrup_1['HH_Disponibles'].max(), 2.6)
        ax1_b.bar_label(b1, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax1_b.bar_label(b2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        dibujar_guias_meses(ax1_b, len(x_indices_1))

        for i, bar in enumerate(b1):
            val_u = int(agrup_1['Cant._Prod._A1'].iloc[i])
            if val_u > 0: 
                ax1_b.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{val_u} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=c_negro, zorder=4)

        ax1_l.plot(x_indices_1, agrup_1['Efic_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=c_blanco, label='% Eficiencia Real', zorder=5)
        ax1_l.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_l.text(x_indices_1[0], 86, 'META = 85%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_l.set_ylim(0, max(120, agrup_1['Efic_Real'].max()*1.8))
        ax1_l.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(agrup_1['Efic_Real']):
            ax1_l.annotate(f"{val:.1f}%", (x_indices_1[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax1_b.set_xticks(x_indices_1); ax1_b.set_xticklabels(agrup_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1_b.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_l.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        st.pyplot(fig1)
    else: st.warning("⚠️ Sin registros para Eficiencia Real.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    if not df_m1_build.empty:
        agrup_2 = df_m1_build.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Productivas_C/GAP': 'sum'}).reset_index()
        agrup_2['Ef_Prod'] = (agrup_2['HH_STD_TOTAL'] / agrup_2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2_b = plt.subplots(figsize=(14, 10))
        ax2_l = ax2_b.twinx()
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        
        x_indices_2 = np.arange(len(agrup_2))
        b_s_2 = ax2_b.bar(x_indices_2 - 0.17, agrup_2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_p_2 = ax2_b.bar(x_indices_2 + 0.17, agrup_2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_margen_superior(ax2_b, max(agrup_2['HH_STD_TOTAL'].max(), agrup_2['HH_Productivas_C/GAP'].max()), 2.6)
        ax2_b.bar_label(b_s_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax2_b.bar_label(b_p_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        ax2_l.plot(x_indices_2, agrup_2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=c_blanco, label='% Efic. Prod.', zorder=5)
        ax2_l.axhline(100, color='darkgreen', linestyle='--', linewidth=3)
        ax2_l.text(x_indices_2[0], 101, 'META = 100%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        ax2_l.set_ylim(0, max(150, agrup_2['Ef_Prod'].max()*1.8))
        ax2_l.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(agrup_2['Ef_Prod']):
            ax2_l.annotate(f"{val:.1f}%", (x_indices_2[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax2_b.set_xticks(x_indices_2); ax2_b.set_xticklabels(agrup_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_b.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        st.pyplot(fig2)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (ANÁLISIS DE BRECHA Y COSTOS)
# =========================================================================
col3, col4 = st.columns(2)

with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        c_prod_maestra = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        agrup_3 = df_ef_f.groupby('Fecha').agg({c_prod_maestra: 'sum', 'HH_Improductivas': 'sum', 'HH_Disponibles': 'sum'}).reset_index()
        agrup_3['Suma_Declaradas'] = agrup_3[c_prod_maestra] + agrup_3['HH_Improductivas']
        
        fig3, ax3 = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        x_v3 = np.arange(len(agrup_3))
        
        ax3.bar(x_v3, agrup_3[c_prod_maestra], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3.bar(x_v3, agrup_3['HH_Improductivas'], bottom=agrup_3[c_prod_maestra], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax3.plot(x_v3, agrup_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=c_blanco, label='HH DISPONIBLES', zorder=5)
        
        set_margen_superior(ax3, agrup_3['HH_Disponibles'].max(), 2.6)
        for i in range(len(x_v3)):
            h_dis, h_dec = agrup_3['HH_Disponibles'].iloc[i], agrup_3['Suma_Declaradas'].iloc[i]
            v_gap = h_dis - h_dec
            ax3.plot([i, i], [h_dec, h_dis], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            ax3.annotate(f"GAP:\n{int(v_gap)}", (i, h_dec + 5), color='firebrick', bbox=bbox_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            ax3.annotate(f"{int(h_dis)}", (i, h_dis + (ax3.get_ylim()[1]*0.08)), color='black', bbox=bbox_blanco, ha='center', fontweight='bold', zorder=10)

        ax3.set_xticks(x_v3); ax3.set_xticklabels(agrup_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        st.pyplot(fig3)
    else: st.warning("⚠️ Sin datos para análisis de GAP.")

with col4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización del impacto económico de las improductividades</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        agrup_4 = df_ef_f.groupby('Fecha').agg({'HH_Improductivas': 'sum', 'Costo_Improd._$': 'sum'}).reset_index()
        fig4, ax4_i = plt.subplots(figsize=(14, 10)); ax4_d = ax4_i.twinx()
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)

        x_v4 = np.arange(len(agrup_4))
        bar_imp4 = ax4_i.bar(x_v4, agrup_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_i.bar_label(bar_imp4, padding=4, color='black', fontweight='bold', path_effects=c_blanco, zorder=4)
        
        set_margen_superior(ax4_i, agrup_4['HH_Improductivas'].max(), 2.6)
        ax4_d.plot(x_v4, agrup_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=c_blanco, label='COSTO ARS', zorder=5)
        ax4_d.set_ylim(0, max(1000, agrup_4['Costo_Improd._$'].max() * 1.8))
        ax4_d.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_d.get_yticks()], fontweight='bold')
        
        total_p = agrup_4['Costo_Improd._$'].sum()
        ax4_i.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${total_p:,.0f}", transform=ax4_i.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=bbox_oro, weight='bold', zorder=10)
        for i, v in enumerate(agrup_4['Costo_Improd._$']): ax4_d.annotate(f"${v:,.0f}", (x_v4[i], v + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)
        ax4_i.set_xticks(x_v4); ax4_i.set_xticklabels(agrup_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_i.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        st.pyplot(fig4)
    else: st.warning("⚠️ Sin datos de costos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (PARETO E INCIDENCIA)
# =========================================================================
col5, col6 = st.columns(2)

with col5:
    st.header("5. PARETO DE CAUSAS")
    if not df_im_f.empty:
        agrup_5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        n_m = df_im_f['FECHA'].nunique()
        div_m = n_m if n_m > 0 else 1
        agrup_5['Prom_M'] = agrup_5['HH_IMPRODUCTIVAS'] / div_m
        agrup_5 = agrup_5.sort_values(by='Prom_M', ascending=False)
        agrup_5['%_Acu'] = (agrup_5['Prom_M'].cumsum() / agrup_5['Prom_M'].sum()) * 100
        fig5, ax5 = plt.subplots(figsize=(14, 10)); ax5b = ax5.twinx()
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        pos_x = np.arange(len(agrup_5))
        b_p5 = ax5.bar(pos_x, agrup_5['Prom_M'], color='maroon', edgecolor='white', zorder=2)
        set_margen_superior(ax5, agrup_5['Prom_M'].max(), 2.8)
        ax5.bar_label(b_p5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        ax5b.plot(pos_x, agrup_5['%_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=c_blanco, zorder=5)
        ax5b.axhline(80, color='gray', linestyle='--'); ax5b.set_ylim(0, 200); ax5b.yaxis.set_major_formatter(mtick.PercentFormatter())
        ax5.set_xticks(pos_x); ax5.set_xticklabels([textwrap.fill(str(t), 12) for t in agrup_5['TIPO_PARADA']], rotation=90, fontsize=12, fontweight='bold')
        for i, val in enumerate(agrup_5['%_Acu']): ax5b.annotate(f"{val:.1f}%", (pos_x[i], val + 4), color='white', bbox=bbox_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
        ax5.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL: {agrup_5['Prom_M'].sum():.1f} HH", transform=ax5.transAxes, bbox=bbox_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        st.pyplot(fig5)
        st.markdown("### 🛠️ Mesa de Trabajo")
        df_tbl = agrup_5.copy(); t_hs = df_tbl['HH_IMPRODUCTIVAS'].sum()
        df_tbl['% sobre Selección'] = (df_tbl['HH_IMPRODUCTIVAS'] / t_hs) * 100
        f_tot = pd.DataFrame({'TIPO_PARADA': ['✅ TOTAL'], 'HH_IMPRODUCTIVAS': [t_hs], 'Prom_M': [df_tbl['Prom_M'].sum()], '%_Acu': [100.0], '% sobre Selección': [100.0]})
        df_tbl = pd.concat([df_tbl, f_tot], ignore_index=True)
        st.dataframe(df_tbl.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH'}), use_container_width=True, hide_index=True)
        st.download_button(label="📥 Descargar Plan (CSV)", data=df_tbl.to_csv(index=False).encode('utf-8'), file_name="Plan_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else: st.success("✅ Sin registros.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    if not df_ef_f.empty:
        df_ef_f['K_Cru'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_d6 = df_ef_f.groupby('K_Cru', as_index=False)['HH_Disponibles'].sum()
        if not df_im_f.empty:
            df_im_f['K_Cru'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K_Cru', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6_f = pd.merge(ag_d6, piv6, on='K_Cru', how='left').fillna(0)
            list_6 = [c for c in df6_f.columns if c not in ['HH_Disponibles', 'K_Cru']]
        else: df6_f = ag_d6.copy(); list_6 = []
        df6_f['Suma_I'] = df6_f[list_6].sum(axis=1) if list_6 else 0
        df6_f['Inc_Pct'] = (df6_f['Suma_I'] / df6_f['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6_f['Sort'] = pd.to_datetime(df6_f['K_Cru'] + '-01')
        df6_f = df6_f.sort_values(by='Sort')
        fig6, ax6 = plt.subplots(figsize=(14, 10)); fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        pos_6 = np.arange(len(df6_f))
        if list_6:
            piso = np.zeros(len(df6_f)); pal = plt.cm.tab20.colors
            for i, nom in enumerate(list_6):
                val_c = df6_f[nom].values
                ax6.bar(pos_6, val_c, bottom=piso, label=nom, color=pal[i % 20], edgecolor='white', zorder=2)
                piso += val_c
        else: ax6.bar(pos_6, np.zeros(len(df6_f)), color='white')
        set_margen_superior(ax6, df6_f['Suma_I'].max(), 2.2)
        for i in range(len(pos_6)):
            iv, dv = df6_f['Suma_I'].iloc[i], df6_f['HH_Disponibles'].iloc[i]
            if iv > 0: ax6.annotate(f"Imp: {int(iv)}\nDisp: {int(dv)}", (i, iv + (ax6.get_ylim()[1]*0.05)), ha='center', bbox=bbox_oro, fontweight='bold', zorder=10)
        ax6t = ax6.twinx()
        ax6t.plot(pos_6, df6_f['Inc_Pct'], color='red', marker='o', markersize=12, linewidth=6, path_effects=contorno_blanco, label='% Incidencia', zorder=5)
        ax6t.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6t.text(pos_6[0], 16, 'META = 15%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        for i, val in enumerate(df6_f['Inc_Pct']): ax6t.annotate(f"{val:.1f}%", (pos_6[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=contorno_blanco, zorder=10)
        ax6.set_xticks(pos_6); ax6.set_xticklabels(df6_f['K_Cru'], fontsize=14, fontweight='bold')
        ax6t.set_ylim(0, max(30, df6_f['Inc_Pct'].max() * 1.8))
        st.pyplot(fig6)
