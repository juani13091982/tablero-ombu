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
# 1. CONFIGURACIÓN DE PÁGINA Y MARCO GERENCIAL
# =========================================================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================================
# 2. SISTEMA DE SEGURIDAD CORPORATIVA (LLAVE MAESTRA)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    """Dibuja la pantalla institucional de acceso con el diseño corporativo."""
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_v1, col_log, col_v2 = st.columns([1, 1.8, 1])
    
    with col_log:
        # Franja Azul Ombú Superior
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGOTIPO PEQUEÑO Y CENTRADO
        s_l, s_c, s_r = st.columns([1, 1, 1])
        with s_c:
            try:
                st.image("LOGO OMBÚ.jpg", width=160)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        # TEXTO INSTITUCIONAL OFICIAL
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Formulario de entrada
        with st.form("form_login_seguro"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            u_digitado = st.text_input("Usuario Corporativo")
            p_digitado = st.text_input("Contraseña", type="password")
            
            entrar = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if entrar:
                if u_digitado in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_digitado] == p_digitado:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos.")

if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES Y PANEL DE FILTROS FIJO (STICKY)
# =========================================================================
st.markdown("""
<style>
    /* Ocultar elementos estándar para profesionalismo */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* PANEL DE FILTROS STICKY */
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 0px !important;
        background-color: #0E1117 !important; 
        z-index: 99999 !important;
        padding-top: 15px !important;
        padding-bottom: 15px !important;
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
box_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
box_oro = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
box_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
box_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. MOTOR INTELIGENTE DE CRUCE DE DATOS (MOTOR FUZZY)
# =========================================================================
def escala_superior(ax_plot, v_max, multiplicador=2.6):
    """Asegura que el gráfico tenga 'techo' para que las etiquetas no se pisen."""
    if v_max > 0: 
        ax_plot.set_ylim(0, v_max * multiplicador)
    else: 
        ax_plot.set_ylim(0, 100)

def motor_robusto_cruce(sel_actual, val_excel):
    """
    Busca coincidencias aunque los textos no sean idénticos.
    Fundamental para no perder registros de improductividades.
    """
    if pd.isna(val_excel) or pd.isna(sel_actual): 
        return False
    
    # Normalización profunda
    s1 = str(sel_actual).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(val_excel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # Solo alfanuméricos
    l1 = re.sub(r'[^A-Z0-9]', '', s1)
    l2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not l1 or not l2: 
        return False
        
    if l1 in l2 or l2 in l1: 
        return True
    
    # Coincidencia por código de estación (3+ dígitos)
    n1 = set(re.findall(r'\d{3,}', s1))
    n2 = set(re.findall(r'\d{3,}', s2))
    if n1 and n2 and n1.intersection(n2): 
        return True
        
    # Coincidencia por raíces de palabras clave
    w1 = set(re.findall(r'[A-Z]{4,}', s1))
    w2 = set(re.findall(r'[A-Z]{4,}', s2))
    
    excluir = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1_limpio = w1 - excluir
    v2_limpio = w2 - excluir
    
    for word in v1_limpio:
        if any(word in x for x in v2_limpio): 
            return True
                
    return False

# =========================================================================
# 5. HEADER Y BOTÓN DE SALIDA
# =========================================================================
h_l, h_t, h_s = st.columns([1, 3, 1])

with h_l:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with h_t:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Control de Gestión Productiva</p>", unsafe_allow_html=True)

with h_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA DE DATOS (EXCEL)
# =========================================================================
try:
    # Carga de archivos
    df_ef_base = pd.read_excel("eficiencias.xlsx")
    df_im_base = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de columnas
    df_ef_base.columns = df_ef_base.columns.str.strip()
    df_im_base.columns = [str(c).strip().upper() for c in df_im_base.columns]
    
    # Auto-corrector de mapeo para improductivas
    if 'TIPO_PARADA' not in df_im_base.columns:
        c_alt = next((c for c in df_im_base.columns if 'TIPO' in c or 'MOTIVO' in c), None)
        if c_alt: df_im_base.rename(columns={c_alt: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im_base.columns:
        c_hs = next((c for c in df_im_base.columns if 'HH' in c and 'IMP' in c), None)
        if c_hs: df_im_base.rename(columns={c_hs: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im_base.columns:
        c_fec = next((c for c in df_im_base.columns if 'FECHA' in c), None)
        if c_fec: df_im_base.rename(columns={c_fec: 'FECHA'}, inplace=True)
    
    # Estandarización de Fechas
    df_ef_base['Fecha'] = pd.to_datetime(df_ef_base['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im_base['FECHA'] = pd.to_datetime(df_im_base['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación técnica
    df_ef_base['Es_Ultimo_Puesto'] = df_ef_base['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Etiquetas de filtros
    df_ef_base['Filtro_Mes'] = df_ef_base['Fecha'].dt.strftime('%b-%Y')
    df_im_base['Filtro_Mes'] = df_im_base['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e_error:
    st.error(f"Error cargando los archivos Excel: {e_error}")
    st.stop()

# =========================================================================
# 7. FILTROS EN CASCADA
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    fc1, fc2, fc3, fc4 = st.columns(4)
    
    with fc1: 
        list_p = list(df_ef_base['Planta'].dropna().unique())
        sel_planta = st.multiselect("🏭 Planta", list_p, placeholder="Todas")
        
    df_temp_l = df_ef_base[df_ef_base['Planta'].isin(sel_planta)] if sel_planta else df_ef_base
    
    with fc2: 
        list_l = list(df_temp_l['Linea'].dropna().unique())
        sel_linea = st.multiselect("⚙️ Línea", list_l, placeholder="Todas")
        
    df_temp_p = df_temp_l[df_temp_l['Linea'].isin(sel_linea)] if sel_linea else df_temp_l
    
    with fc3: 
        list_ps = list(df_temp_p['Puesto_Trabajo'].dropna().unique())
        sel_puesto = st.multiselect("🛠️ Puesto de Trabajo", list_ps, placeholder="Todos")
        
    with fc4: 
        list_m = ["🎯 Acumulado YTD"] + list(df_ef_base['Filtro_Mes'].unique())
        sel_mes = st.multiselect("📅 Mes", list_m, placeholder="Todos")

# =========================================================================
# 8. LÓGICA DE FILTRADO FINAL (RECUPERACIÓN DE VARIABLES)
# =========================================================================
df_ef_f = df_ef_base.copy()
df_im_f = df_im_base.copy()

# Filtros Eficiencias
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
if sel_linea: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]
if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_ef_f = df_ef_f[df_ef_f['Filtro_Mes'].isin(sel_mes)]

# Filtros Improductivas (Motor Fuzzy)
if sel_planta:
    mask_pl = df_im_f.iloc[:,0].apply(lambda x: any(motor_robusto_cruce(p, x) for p in sel_planta))
    df_im_f = df_im_f[mask_pl]

if sel_linea:
    c_l_search = next((c for c in df_im_f.columns if 'LINEA' in c), df_im_f.columns[1])
    mask_li = df_im_f[c_l_search].apply(lambda x: any(motor_robusto_cruce(l, x) for l in sel_linea))
    df_im_f = df_im_f[mask_li]

if sel_puesto:
    c_p_search = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2])
    mask_ps = df_im_f[c_p_search].apply(lambda x: any(motor_robusto_cruce(ps, x) for ps in sel_puesto))
    df_im_f = df_im_f[mask_ps]

if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_im_f = df_im_f[df_im_f['Filtro_Mes'].isin(sel_mes)]

txt_h_graficos = f"Filtros: {sel_planta if sel_planta else 'Todas'} > {sel_linea if sel_linea else 'Todas'} > {sel_puesto if sel_puesto else 'Todos'}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (PRODUCTIVIDAD)
# =========================================================================
c1, c2 = st.columns(2)

with c1:
    st.header("1. EFICIENCIA REAL")
    # FORMULA SOLICITADA EXPLICITAMENTE
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    # Lógica de dibujo: Si no hay puesto, mostramos Salida de Línea.
    df_m1_plot = df_ef_f.copy() if sel_puesto else df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_plot.empty:
        # Agrupación Mensual
        ag_1 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        ag_1['Ef_Real'] = (ag_1['HH_STD_TOTAL'] / ag_1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Gráfico
        fig1, ax1_b = plt.subplots(figsize=(14, 10))
        ax1_l = ax1_b.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices = np.arange(len(ag_1))
        
        # Barras de Horas
        bar_std = ax1_b.bar(x_indices - 0.17, ag_1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_dis = ax1_b.bar(x_indices + 0.17, ag_1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        escala_superior(ax1_b, ag_1['HH_Disponibles'].max(), 2.6)
        ax1_b.bar_label(bar_std, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax1_b.bar_label(bar_dis, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        for i in range(len(x_indices)):
            ax1_b.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

        # Unidades producidas (Vertical)
        for i, bar in enumerate(bar_std):
            u_val = int(ag_1['Cant._Prod._A1'].iloc[i])
            if u_val > 0: 
                ax1_b.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{u_val} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=c_negro, zorder=4)

        # Línea de Eficiencia
        ax1_l.plot(x_indices, ag_1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=c_blanco, label='% Efic. Real', zorder=5)
        
        # Meta 85%
        ax1_l.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_l.text(x_indices[0], 86, 'META = 85%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_l.set_ylim(0, max(120, ag_1['Ef_Real'].max()*1.8))
        ax1_l.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag_1['Ef_Real']):
            ax1_l.annotate(f"{val:.1f}%", (x_indices[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax1_b.set_xticks(x_indices)
        ax1_b.set_xticklabels(ag_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1_b.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_l.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin datos para Eficiencia Real.")

with c2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_plot.empty:
        # Agrupación temporal
        ag_2 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        ag_2['Ef_P'] = (ag_2['HH_STD_TOTAL'] / ag_2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Gráfico
        fig2, ax2_b = plt.subplots(figsize=(14, 10))
        ax2_l = ax2_b.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx2 = np.arange(len(ag_2))
        
        # Barras
        b_s_2 = ax2_b.bar(x_idx2 - 0.17, ag_2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_p_2 = ax2_b.bar(x_idx2 + 0.17, ag_2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        escala_superior(ax2_b, max(ag_2['HH_STD_TOTAL'].max(), ag_2['HH_Productivas_C/GAP'].max()), 2.6)
        ax2_b.bar_label(b_s_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax2_b.bar_label(b_p_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        for i in range(len(x_idx2)):
            ax2_b.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

        # Línea Eficiencia Productiva
        ax2_l.plot(x_idx2, ag_2['Ef_P'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=c_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta 100%
        ax2_l.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_l.text(x_idx2[0], 101, 'META = 100%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_l.set_ylim(0, max(150, ag_2['Ef_P'].max()*1.8))
        ax2_l.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag_2['Ef_P']):
            ax2_l.annotate(f"{val:.1f}%", (x_idx2[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax2_b.set_xticks(x_idx2)
        ax2_b.set_xticklabels(ag_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_b.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ Sin datos para Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (ANÁLISIS DE GAP Y COSTOS)
# =========================================================================
c3, c4 = st.columns(2)

with c3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        # CORRECCIÓN DE VARIABLE: Asegurando nombre exacto
        c_p_maestra = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        
        ag_3 = df_ef_f.groupby('Fecha').agg({
            c_p_maestra: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        ag_3['Suma_Declaradas'] = ag_3[c_p_maestra] + ag_3['HH_Improductivas']
        
        # Gráfico
        fig3, ax3 = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_vals_3 = np.arange(len(ag_3))
        
        # Barras Apiladas
        ax3.bar(x_vals_3, ag_3[c_p_maestra], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3.bar(x_vals_3, ag_3['HH_Improductivas'], bottom=ag__3[c_p_maestra] if 'ag_3' in locals() else 0, color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        # Fix del nombre de la variable de abajo para evitar el error anterior
        ax3.bar(x_vals_3, ag_3['HH_Improductivas'], bottom=ag_3[c_p_maestra], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante
        ax3.plot(x_vals_3, ag_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=c_blanco, label='HH DISPONIBLES', zorder=5)
        
        escala_superior(ax3, ag_3['HH_Disponibles'].max(), 2.6)

        for i in range(len(x_vals_3)):
            v_dis = ag_3['HH_Disponibles'].iloc[i]
            v_dec = ag_3['Suma_Declaradas'].iloc[i]
            v_brecha = v_dis - v_dec
            
            # Flecha GAP
            ax3.plot([i, i], [v_dec, v_dis], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            # Etiqueta GAP
            ax3.annotate(f"GAP:\n{int(v_brecha)}", (i, v_dec + 5), color='firebrick', bbox=box_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            # Valor Disponibilidad
            ax3.annotate(f"{int(v_dis)}", (i, v_dis + (ax3.get_ylim()[1]*0.08)), color='black', bbox=box_blanco, ha='center', fontweight='bold', zorder=10)

        ax3.set_xticks(x_vals_3)
        ax3.set_xticklabels(ag_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ Sin datos para análisis de GAP.")

with c4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización del impacto de las improductividades</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag_4 = df_ef_f.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        # Gráfico
        fig4, ax4_i = plt.subplots(figsize=(14, 10))
        ax4_d = ax4_i.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_vals_4 = np.arange(len(ag_4))
        
        # Barras Horas Perdidas
        b_lost_4 = ax4_i.bar(x_vals_4, ag_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_i.bar_label(b_lost_4, padding=4, color='black', fontweight='bold', path_effects=c_blanco, zorder=4)
        
        escala_superior(ax4_i, ag_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Pesos
        ax4_d.plot(x_vals_4, ag_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=c_blanco, label='COSTO ARS', zorder=5)
        
        ax4_d.set_ylim(0, max(1000, ag_4['Costo_Improd._$'].max() * 1.8))
        ax4_d.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_d.get_yticks()], fontweight='bold')

        # Cartelera Acumulada
        tot_p = ag_4['Costo_Improd._$'].sum()
        ax4_i.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${tot_p:,.0f}", transform=ax4_i.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=box_oro, weight='bold', zorder=10)

        for i, v in enumerate(ag_4['Costo_Improd._$']):
            ax4_d.annotate(f"${v:,.0f}", (x_vals_4[i], v + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax4_i.set_xticks(x_vals_4)
        ax4_i.set_xticklabels(ag_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_i.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ Sin datos de costos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (ANÁLISIS CAUSA RAÍZ)
# =========================================================================
c5, c6 = st.columns(2)

with c5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        # Lógica de Pareto
        df_p_data = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de periodos
        n_meses_u = df_im_f['FECHA'].nunique()
        divisor_p = n_meses_u if n_meses_u > 0 else 1
        
        df_p_data['Promedio_M'] = df_p_data['HH_IMPRODUCTIVAS'] / divisor_p
        df_p_data = df_p_data.sort_values(by='Promedio_M', ascending=False)
        df_p_data['%_Acum'] = (df_p_data['Promedio_M'].cumsum() / df_p_data['Promedio_M'].sum()) * 100

        # Gráfico
        fig5, ax5_i = plt.subplots(figsize=(14, 10))
        ax5_d = ax5_i.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_x_5 = np.arange(len(df_p_data))
        
        # Barras
        bar_p_5 = ax5_i.bar(pos_x_5, df_p_data['Promedio_M'], color='maroon', edgecolor='white', zorder=2)
        escala_superior(ax5_i, df_p_data['Promedio_M'].max(), 2.8)
        ax5_i.bar_label(bar_p_5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva Pareto
        ax5_d.plot(pos_x_5, df_p_data['%_Acum'], color='red', marker='D', markersize=10, linewidth=4, path_effects=c_blanco, zorder=5)
        ax5_d.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax5_d.set_ylim(0, 200)
        ax5_d.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Etiquetas envolventes
        labels_w = [textwrap.fill(str(t), 12) for t in df_p_data['TIPO_PARADA']]
        ax5_i.set_xticks(pos_x_5)
        ax5_i.set_xticklabels(labels_w, rotation=90, fontsize=12, fontweight='bold')
        
        for i, v in enumerate(df_p_data['%_Acum']):
            ax5_d.annotate(f"{v:.1f}%", (pos_x_5[i], v + 4), color='white', bbox=box_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        s_promedio = df_p_data['Promedio_M'].sum()
        ax5_i.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{s_promedio:.1f} HH", transform=ax5_i.transAxes, bbox=box_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl = df_p_data.copy()
        t_hs_total = df_tbl['HH_IMPRODUCTIVAS'].sum()
        df_tbl['% sobre Selección'] = (df_tbl['HH_IMPRODUCTIVAS'] / t_hs_total) * 100
        
        # FILA TOTAL MAESTRA
        f_total = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [t_hs_total], 
            'Promedio_M': [df_tbl['Promedio_M'].sum()],
            '%_Acum': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tbl = pd.concat([df_tbl, f_total], ignore_index=True)
        
        st.dataframe(
            df_tbl.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Botón de Descarga CSV
        c_data = df_tbl.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=c_data, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Felicidades! Cero horas improductivas en este periodo.")

with c6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de Horas Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        # Cruce de datos por Mes
        df_ef_f['Mes_Key'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp_6 = df_ef_f.groupby('Mes_Key', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['Mes_Key'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            pivote_6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='Mes_Key', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_6_final = pd.merge(ag_disp_6, pivote_6, on='Mes_Key', how='left').fillna(0)
            list_causas = [c for c in df_6_final.columns if c not in ['HH_Disponibles', 'Mes_Key']]
        else:
            df_6_final = ag_disp_6.copy()
            list_causas = []
            
        # Indicadores finales de incidencia
        df_6_final['Total_Imp_Hs'] = df_6_final[list_causas].sum(axis=1) if list_causas else 0
        df_6_final['Inc_Pct'] = (df_6_final['Total_Imp_Hs'] / df_6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden Cronológico
        df_6_final['Fecha_O'] = pd.to_datetime(df_6_final['Mes_Key'] + '-01')
        df_6_final = df_6_final.sort_values(by='Fecha_O')

        # Gráfico
        fig6, ax6_b = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(txt_h_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_6 = np.arange(len(df_6_final))
        
        # Barras Apiladas
        if list_causas:
            fondo_6 = np.zeros(len(df_6_final))
            pal_6 = plt.cm.tab20.colors
            for idx, c_nom in enumerate(list_causas):
                vals_c = df_6_final[c_nom].values
                ax6_b.bar(pos_6, vals_c, bottom=fondo_6, label=c_nom, color=pal_6[idx % 20], edgecolor='white', zorder=2)
                fondo_6 += vals_c
        else:
            ax6_b.bar(pos_6, np.zeros(len(df_6_final)), color='white')

        escala_superior(ax6_b, df_6_final['Total_Imp_Hs'].max(), 2.2)
        
        # Anotaciones de volumen
        for i in range(len(pos_6)):
            iv = df_6_final['Total_Imp_Hs'].iloc[i]
            dv = df_6_final['HH_Disponibles'].iloc[i]
            if iv > 0: 
                ax6_b.annotate(f"Imp: {int(iv)}\nDisp: {int(dv)}", (i, iv + (ax6_b.get_ylim()[1]*0.05)), ha='center', bbox=box_oro, fontweight='bold', zorder=10)

        # Línea de Porcentaje de Incidencia
        ax6_l = ax6_b.twinx()
        ax6_l.plot(pos_6, df_6_final['Inc_Pct'], color='red', marker='o', markersize=12, linewidth=6, path_effects=c_blanco, label='% Incidencia', zorder=5)
        
        # Meta 15%
        ax6_l.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_l.text(pos_6[0], 16, 'META = 15%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(df_6_final['Inc_Pct']):
            ax6_l.annotate(f"{val:.1f}%", (pos_6[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=c_blanco, zorder=10)

        ax6_b.set_xticks(pos_6)
        ax6_b.set_xticklabels(df_6_final['Mes_Key'], fontsize=14, fontweight='bold')
        ax6_l.set_ylim(0, max(30, df_6_final['Inc_Pct'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia para este sector.")
