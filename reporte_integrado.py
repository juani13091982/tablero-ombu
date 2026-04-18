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

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS GLOBALES
# ==========================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 🚨 ACCESO GENERAL ÚNICO - SEGURIDAD CORPORATIVA
# ==========================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        # CONTENEDOR ESTÉTICO SUPERIOR
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGO PEQUEÑO Y TEXTO INSTITUCIONAL SOLICITADO
        inner_col1, inner_col2, inner_col3 = st.columns([1, 1, 1])
        with inner_col2:
            try:
                st.image("LOGO OMBÚ.jpg", width=150)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            # Inputs de usuario
            user_input = st.text_input("Usuario Corporativo")
            pass_input = st.text_input("Contraseña", type="password")
            
            # Botón de ingreso
            submit_login = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if submit_login:
                if user_input in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[user_input] == pass_input:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos e intente nuevamente.")

# Verificación de Seguridad
if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# ==========================================
# DISEÑO VISUAL Y ESTILOS (STICKY PANEL Y FUENTES)
# ==========================================
css_styles = """
<style>
    /* Ocultar elementos estándar de Streamlit para limpieza visual */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* PANEL DE FILTROS FIJO (STICKY) */
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 0px !important;
        background-color: #0E1117 !important; 
        z-index: 99999 !important;
        padding-top: 10px !important;
        padding-bottom: 10px !important;
        border-bottom: 3px solid #1E3A8A !important; 
    }
</style>
"""
st.markdown(css_styles, unsafe_allow_html=True)

# Configuración Maestra de Matplotlib (Fuentes Grandes y Negrita para reportes gerenciales)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Definición de efectos de contorno para legibilidad de etiquetas sobre barras
outline_white = [path_effects.withStroke(linewidth=3, foreground='white')]
outline_black = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cuadros de texto (Bboxes)
bbox_gray = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_yellow = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_green = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
bbox_white = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# ==========================================
# FUNCIONES AUXILIARES DE LÓGICA Y CRUCE
# ==========================================
def aplicar_anti_overlap(ax, max_val, multiplier=2.6):
    """
    Función crucial: Escala el eje Y para dejar un 'cielo' despejado arriba 
    y evitar que las etiquetas choquen con los bordes o títulos.
    """
    if max_val > 0: 
        ax.set_ylim(0, max_val * multiplier)
    else: 
        ax.set_ylim(0, 100)

def dibujar_meses(ax, fechas):
    """Líneas de referencia vertical para separar los periodos mensuales."""
    for x in range(len(fechas)):
        ax.axvline(x=x, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_seleccion(lista_sel, default_str):
    """Formatea el texto de los filtros seleccionados para los títulos de los gráficos."""
    if not lista_sel: 
        return default_str
    if len(lista_sel) > 2: 
        return f"Varios ({len(lista_sel)})"
    return " + ".join(lista_sel)

def robust_match(val_sel, val_imp):
    """
    MOTOR FUZZY INTELIGENTE: Cruza Eficiencias con Improductivas.
    Resuelve el problema de que los nombres de los puestos no coinciden exactamente.
    """
    if pd.isna(val_imp) or pd.isna(val_sel): 
        return False
    
    # Limpieza profunda de strings (mayúsculas, tildes, caracteres especiales)
    s1 = str(val_sel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(val_imp).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # Solo caracteres alfanuméricos para comparación base
    c1 = re.sub(r'[^A-Z0-9]', '', s1)
    c2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not c1 or not c2: 
        return False
        
    # 1. Coincidencia exacta de texto continuo
    if c1 in c2 or c2 in c1: 
        return True
    
    # 2. Coincidencia por códigos numéricos (si ambos tienen el mismo código de 3+ dígitos)
    num1 = set(re.findall(r'\d{3,}', s1))
    num2 = set(re.findall(r'\d{3,}', s2))
    if num1 and num2 and num1.intersection(num2): 
        return True
        
    # 3. Búsqueda por palabras raíz (evitando términos genéricos)
    words1 = set(re.findall(r'[A-Z]{4,}', s1))
    words2 = set(re.findall(r'[A-Z]{4,}', s2))
    
    exclusion = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1 = words1 - exclusion
    v2 = words2 - exclusion
    
    for w1 in v1:
        for w2 in v2:
            if w1 in w2 or w2 in w1: 
                return True
                
    return False

# ==========================================
# HEADER PRINCIPAL Y SALIDA
# ==========================================
col_logo_h, col_title_h, col_exit_h = st.columns([1, 3, 1])

with col_logo_h:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with col_title_h:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")

with col_exit_h:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# ==========================================
# CARGA DE DATOS Y AUTO-CORRECTOR DE COLUMNAS
# ==========================================
try:
    # Carga de archivos Excel
    df_ef = pd.read_excel("eficiencias.xlsx")
    df_imp = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de encabezados
    df_ef.columns = df_ef.columns.str.strip()
    df_imp.columns = [str(c).strip().upper() for c in df_imp.columns]
    
    # Mapeo inteligente de columnas para evitar errores de carga
    if 'TIPO_PARADA' not in df_imp.columns:
        col_motivo = next((c for c in df_imp.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if col_motivo: 
            df_imp.rename(columns={col_motivo: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_imp.columns:
        col_horas = next((c for c in df_imp.columns if 'HH' in c and 'IMP' in c), None)
        if col_horas: 
            df_imp.rename(columns={col_horas: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_imp.columns:
        col_fecha_imp = next((c for c in df_imp.columns if 'FECHA' in c), None)
        if col_fecha_imp: 
            df_imp.rename(columns={col_fecha_imp: 'FECHA'}, inplace=True)
    
    # Procesamiento de Fechas a primer día del mes
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_imp['FECHA'] = pd.to_datetime(df_imp['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación técnica de puestos de salida
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Creación de etiquetas de filtro legibles
    df_ef['Mes_Filtro'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_imp['Mes_Filtro'] = df_imp['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e:
    st.error(f"Error crítico en la carga de datos: {e}")
    st.stop()

# ==========================================
# FILTROS SUPERIORES (CASCADA DINÁMICA)
# ==========================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    f1, f2, f3, f4 = st.columns(4)
    
    with f1: 
        list_plantas = list(df_ef['Planta'].dropna().unique())
        planta_sel = st.multiselect("🏭 Planta", list_plantas, placeholder="Todas")
        
    df_l_tmp = df_ef[df_ef['Planta'].isin(planta_sel)] if planta_sel else df_ef
    
    with f2: 
        list_lineas = list(df_l_tmp['Linea'].dropna().unique())
        linea_sel = st.multiselect("⚙️ Línea", list_lineas, placeholder="Todas")
        
    df_p_tmp = df_l_tmp[df_l_tmp['Linea'].isin(linea_sel)] if linea_sel else df_l_tmp
    
    with f3: 
        list_puestos = list(df_p_tmp['Puesto_Trabajo'].dropna().unique())
        puesto_sel = st.multiselect("🛠️ Puesto de Trabajo", list_puestos, placeholder="Todos")
        
    with f4: 
        list_meses = ["🎯 Acumulado YTD"] + list(df_ef['Mes_Filtro'].unique())
        mes_sel = st.multiselect("📅 Mes", list_meses, placeholder="Todos")

# ==========================================
# LÓGICA DE FILTRADO FINAL
# ==========================================
df_ef_final = df_ef.copy()
df_imp_final = df_imp.copy()

# Aplicar filtros a Eficiencias
if planta_sel: 
    df_ef_final = df_ef_final[df_ef_final['Planta'].isin(planta_sel)]
if linea_sel: 
    df_ef_final = df_ef_final[df_ef_final['Linea'].isin(linea_sel)]
if puesto_sel: 
    df_ef_final = df_ef_final[df_ef_final['Puesto_Trabajo'].isin(puesto_sel)]
if mes_sel and "🎯 Acumulado YTD" not in mes_sel:
    df_ef_final = df_ef_final[df_ef_final['Mes_Filtro'].isin(mes_sel)]

# Aplicar filtros a Improductivas (Uso del motor inteligente para Planta y Puesto)
if planta_sel:
    mask_pl = df_imp_final.iloc[:,0].apply(lambda x: any(robust_match(s, x) for s in planta_sel))
    df_imp_final = df_imp_final[mask_pl]

if linea_sel:
    col_linea_search = next((c for c in df_imp_final.columns if 'LINEA' in c), df_imp_final.columns[1])
    mask_ln = df_imp_final[col_linea_search].apply(lambda x: any(robust_match(s, x) for s in linea_sel))
    df_imp_final = df_imp_final[mask_ln]

if puesto_sel:
    col_puesto_search = next((c for c in df_imp_final.columns if 'PUESTO' in c), df_imp_final.columns[2])
    mask_ps = df_imp_final[col_puesto_search].apply(lambda x: any(robust_match(s, x) for s in puesto_sel))
    df_imp_final = df_imp_final[mask_ps]

if mes_sel and "🎯 Acumulado YTD" not in mes_sel:
    df_imp_final = df_imp_final[df_imp_final['Mes_Filtro'].isin(mes_sel)]

# Texto dinámico para los encabezados de los gráficos
texto_header_graficos = f"Filtro: {formatear_seleccion(planta_sel, 'Todas')} > {formatear_seleccion(linea_sel, 'Todas')} > {formatear_seleccion(puesto_sel, 'Todos')}"

st.markdown("---")

# =========================================================================
# FILA 1: MÉTRICAS 1 Y 2
# =========================================================================
c1, c2 = st.columns(2)

with c1:
    st.header("1. EFICIENCIA REAL")
    
    # Si no hay puesto seleccionado, mostramos solo "Salida de Línea" (Es_Ultimo_Puesto == 'SI')
    df_m1_plot = df_ef_final.copy() if puesto_sel else df_ef_final[df_ef_final['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_plot.empty:
        # Agrupación por mes
        agrup_m1 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo de Eficiencia
        agrup_m1['Ef'] = (agrup_m1['HH_STD_TOTAL'] / agrup_m1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Preparación de gráfico
        fig1, ax_bars = plt.subplots(figsize=(14, 10))
        ax_line = ax_bars.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')
        
        x_vals = np.arange(len(agrup_m1))
        ancho_barra = 0.35
        
        # Barras
        bar_std = ax_bars.bar(x_vals - ancho_barra/2, agrup_m1['HH_STD_TOTAL'], ancho_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_disp = ax_bars.bar(x_vals + ancho_barra/2, agrup_m1['HH_Disponibles'], ancho_barra, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        aplicar_anti_overlap(ax_bars, agrup_m1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas de barras
        ax_bars.bar_label(bar_std, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        ax_bars.bar_label(bar_disp, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        
        dibujar_meses(ax_bars, x_vals)

        # Cantidad de unidades producidas (Vertical dentro de la barra)
        for i, b in enumerate(bar_std):
            uni = int(agrup_m1['Cant._Prod._A1'].iloc[i])
            if uni > 0:
                ax_bars.text(b.get_x() + b.get_width()/2, b.get_height()*0.05, f"{uni} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=outline_black, zorder=4)

        # Línea de Eficiencia
        ax_line.plot(x_vals, agrup_m1['Ef'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=outline_white, label='% Eficiencia Real', zorder=5)
        
        # Línea de Meta
        ax_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax_line.text(x_vals[0], 86, 'META = 85%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', zorder=10)
        
        # Formateo de eje Y secundario (%)
        ax_line.set_ylim(0, max(120, agrup_m1['Ef'].max()*1.8))
        ax_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Etiquetas de la línea
        for i, val in enumerate(agrup_m1['Ef']):
            ax_line.annotate(f"{val:.1f}%", (x_vals[i], val + 5), color='white', bbox=bbox_gray, ha='center', fontweight='bold', zorder=10)

        # Configuración final
        ax_bars.set_xticks(x_vals)
        ax_bars.set_xticklabels(agrup_m1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin datos para el filtro seleccionado.")

with c2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_plot.empty:
        agrup_m2 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        agrup_m2['Ef'] = (agrup_m2['HH_STD_TOTAL'] / agrup_m2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax_bars2 = plt.subplots(figsize=(14, 10))
        ax_line2 = ax_bars2.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')
        
        x_vals2 = np.arange(len(agrup_m2))
        
        # Barras
        bar_std2 = ax_bars2.bar(x_vals2 - ancho_barra/2, agrup_m2['HH_STD_TOTAL'], ancho_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_prod2 = ax_bars2.bar(x_vals2 + ancho_barra/2, agrup_m2['HH_Productivas_C/GAP'], ancho_barra, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        aplicar_anti_overlap(ax_bars2, max(agrup_m2['HH_STD_TOTAL'].max(), agrup_m2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax_bars2.bar_label(bar_std2, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        ax_bars2.bar_label(bar_prod2, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        
        dibujar_meses(ax_bars2, x_vals2)

        # Línea de Eficiencia Productiva
        ax_line2.plot(x_vals2, agrup_m2['Ef'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=outline_white, label='% Efic. Prod.', zorder=5)
        
        # Línea de Meta (100%)
        ax_line2.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax_line2.text(x_vals2[0], 101, 'META = 100%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', zorder=10)
        
        ax_line2.set_ylim(0, max(150, agrup_m2['Ef'].max()*1.8))
        ax_line2.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotaciones
        for i, val in enumerate(agrup_m2['Ef']):
            ax_line2.annotate(f"{val:.1f}%", (x_vals2[i], val + 5), color='white', bbox=bbox_gray, ha='center', fontweight='bold', zorder=10)

        ax_bars2.set_xticks(x_vals2)
        ax_bars2.set_xticklabels(agrup_m2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax_bars2.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ Sin datos evaluables.")

st.markdown("---")

# =========================================================================
# FILA 2: MÉTRICAS 3 Y 4
# =========================================================================
c3, c4 = st.columns(2)

with c3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_final.empty:
        # Columna de productivas real
        c_prod_ pura = 'HH_Productivas' if 'HH_Productivas' in df_ef_final.columns else 'HH Productivas'
        
        agrup_m3 = df_ef_final.groupby('Fecha').agg({
            c_prod_pura: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        agrup_m3['Total_Decl'] = agrup_m3[c_prod_pura] + agrup_m3['HH_Improductivas']
        
        fig3, ax_gap = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')
        
        x_m3 = np.arange(len(agrup_m3))
        
        # Barras Apiladas (Productivas e Improductivas)
        ax_gap.bar(x_m3, agrup_m3[c_prod_pura], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax_gap.bar(x_m3, agrup_m3['HH_Improductivas'], bottom=agrup_m3[c_prod_pura], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea de Disponibilidad
        ax_gap.plot(x_m3, agrup_m3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=outline_white, label='HH DISPONIBLES', zorder=5)
        
        aplicar_anti_overlap(ax_gap, agrup_m3['HH_Disponibles'].max(), 2.6)
        dibujar_meses(ax_gap, x_m3)

        # Cálculo y visualización del GAP
        for i in range(len(x_m3)):
            d_val = agrup_m3['HH_Disponibles'].iloc[i]
            t_val = agrup_m3['Total_Decl'].iloc[i]
            g_val = d_val - t_val
            
            # Dibujar flecha/línea de conexión del GAP
            ax_gap.plot([i, i], [t_val, d_val], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Etiqueta de GAP
            ax_gap.annotate(f"GAP:\n{int(g_val)}", (i, t_val + 5), color='firebrick', bbox=bbox_white, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Etiqueta de Disponibilidad Final
            ax_gap.annotate(f"{int(d_val)}", (i, d_val + (ax_gap.get_ylim()[1]*0.08)), color='black', bbox=bbox_white, ha='center', fontweight='bold', zorder=10)

        ax_gap.set_xticks(x_m3)
        ax_gap.set_xticklabels(agrup_m3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax_gap.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ Sin datos evaluables.")

with c4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Valorización del impacto económico de las improductividades</i></div>", unsafe_allow_html=True)
    
    if not df_ef_final.empty:
        agrup_m4 = df_ef_final.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        fig4, ax_c1 = plt.subplots(figsize=(14, 10))
        ax_c2 = ax_c1.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')

        x_m4 = np.arange(len(agrup_m4))
        
        # Barras de horas perdidas
        b_lost = ax_c1.bar(x_m4, agrup_m4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax_c1.bar_label(b_lost, padding=4, color='black', fontweight='bold', path_effects=outline_white, zorder=4)
        
        aplicar_anti_overlap(ax_c1, agrup_m4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Costo en Pesos
        ax_c2.plot(x_m4, agrup_m4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=outline_white, label='COSTO ARS', zorder=5)
        
        # Formato de dinero
        ax_c2.set_ylim(0, max(1000, agrup_m4['Costo_Improd._$'].max() * 1.8))
        ax_c2.set_yticklabels([f'${int(x/1000000)}M' for x in ax_c2.get_yticks()], fontweight='bold')

        # Cartelera de Costo Total Acumulado
        total_acum_costo = agrup_m4['Costo_Improd._$'].sum()
        total_acum_hh = agrup_m4['HH_Improductivas'].sum()
        ax_c1.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${total_acum_costo:,.0f}\n(Pérdida: {total_acum_hh:,.0f} HH)", transform=ax_c1.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=bbox_yellow, weight='bold', zorder=10)

        for i, val in enumerate(agrup_m4['Costo_Improd._$']):
            ax_c2.annotate(f"${val:,.0f}", (x_m4[i], val + 5), color='white', bbox=bbox_gray, ha='center', fontweight='bold', zorder=10)

        ax_c1.set_xticks(x_m4)
        ax_c1.set_xticklabels(agrup_m4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax_c1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ Sin datos económicos disponibles.")

st.markdown("---")

# =========================================================================
# FILA 3: MÉTRICAS 5 Y 6 - ANALISIS DE CAUSA RAÍZ
# =========================================================================
c5, c6 = st.columns(2)

with c5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_imp_final.empty:
        # Agrupación de causas
        df_p_data = df_imp_final.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de meses para el promedio
        n_meses = df_imp_final['FECHA'].nunique()
        div_m = n_meses if n_meses > 0 else 1
        
        df_p_data['Prom_Mes'] = df_p_data['HH_IMPRODUCTIVAS'] / div_m
        df_p_data = df_p_data.sort_values(by='Prom_Mes', ascending=False)
        df_p_data['%_Acumulado'] = (df_p_data['Prom_Mes'].cumsum() / df_p_data['Prom_Mes'].sum()) * 100

        fig5, ax_p1 = plt.subplots(figsize=(14, 10))
        ax_p2 = ax_p1.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')

        x_pos_p = np.arange(len(df_p_data))
        
        # Barras del Pareto
        b_p_m5 = ax_p1.bar(x_pos_p, df_p_data['Prom_Mes'], color='maroon', edgecolor='white', zorder=2)
        aplicar_anti_overlap(ax_p1, df_p_data['Prom_Mes'].max(), 2.8)
        ax_p1.bar_label(b_p_m5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva de Lorenz
        ax_p2.plot(x_pos_p, df_p_data['%_Acumulado'], color='red', marker='D', markersize=10, linewidth=4, path_effects=outline_white, zorder=5)
        ax_p2.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax_p2.set_ylim(0, 200)
        ax_p2.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Wrap de texto para causas largas
        causas_wrap = [textwrap.fill(str(l), 12) for l in df_p_data['TIPO_PARADA']]
        ax_p1.set_xticks(x_pos_p)
        ax_p1.set_xticklabels(causas_wrap, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val_p in enumerate(df_p_data['%_Acumulado']):
            ax_p2.annotate(f"{val_p:.1f}%", (x_pos_p[i], val_p + 4), color='white', bbox=bbox_gray, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        # Información Resumida en el gráfico
        suma_p = df_p_data['Prom_Mes'].sum()
        ax_p1.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{suma_p:.1f} HH", transform=ax_p1.transAxes, bbox=bbox_gray, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_resumen_tbl = df_p_data.copy()
        total_p_hh = df_resumen_tbl['HH_IMPRODUCTIVAS'].sum()
        df_resumen_tbl['% sobre Selección'] = (df_resumen_tbl['HH_IMPRODUCTIVAS'] / total_p_hh) * 100
        
        # Inyección de la fila de TOTAL
        fila_total_pareto = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [total_p_hh], 
            'Prom_Mes': [df_resumen_tbl['Prom_Mes'].sum()],
            '%_Acumulado': [100.0],
            '% sobre Selección': [100.0]
        })
        df_resumen_tbl = pd.concat([df_resumen_tbl, fila_total_pareto], ignore_index=True)
        
        # Despliegue de la Tabla
        st.dataframe(
            df_resumen_tbl.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Causa de Improductividad'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Exportación a CSV
        csv_file = df_resumen_tbl.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Plan de Acción (CSV)", 
            data=csv_file, 
            file_name="Plan_Accion_Ombu.csv", 
            mime="text/csv", 
            use_container_width=True, 
            type="primary"
        )
    else:
        st.success("✅ ¡Excelente! Sin registros de Horas Improductivas para este periodo.")

with c6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Porcentaje histórico de Horas Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_final.empty:
        # Preparación de cruce temporal
        df_ef_final['C_Aux'] = df_ef_final['Fecha'].dt.strftime('%Y-%m')
        disp_m6 = df_ef_final.groupby('C_Aux', as_index=False)['HH_Disponibles'].sum()

        if not df_imp_final.empty:
            df_imp_final['C_Aux'] = df_imp_final['FECHA'].dt.strftime('%Y-%m')
            piv_m6 = pd.pivot_table(df_imp_final, values='HH_IMPRODUCTIVAS', index='C_Aux', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_m6_final = pd.merge(disp_m6, piv_m6, on='C_Aux', how='left').fillna(0)
            cols_paradas = [c for c in df_m6_final.columns if c not in ['HH_Disponibles', 'C_Aux']]
        else:
            df_m6_final = disp_m6.copy()
            cols_paradas = []
            
        # Calculo de indicadores
        df_m6_final['Suma_Imp'] = df_m6_final[cols_paradas].sum(axis=1) if cols_paradas else 0
        df_m6_final['%_Incidencia'] = (df_m6_final['Suma_Imp'] / df_m6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden cronológico
        df_m6_final['Fecha_Sort'] = pd.to_datetime(df_m6_final['C_Aux'] + '-01')
        df_m6_final = df_m6_final.sort_values(by='Fecha_Sort')

        fig6, ax_m6_1 = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(texto_header_graficos, x=0.08, y=0.98, ha='left', fontsize=9, color='dimgray', fontweight='bold')

        x_pos_m6 = np.arange(len(df_m6_final))
        
        # Barras de causas (si existen)
        if cols_paradas:
            fondo_m6 = np.zeros(len(df_m6_final))
            paleta_col = plt.cm.tab20.colors
            for idx_m6, col_nom in enumerate(cols_paradas):
                val_causa = df_m6_final[col_nom].values
                ax_m6_1.bar(x_pos_m6, val_causa, bottom=fondo_m6, label=col_nom, color=paleta_col[idx_m6 % 20], edgecolor='white', zorder=2)
                fondo_m6 += val_causa
        else:
            ax_m6_1.bar(x_pos_m6, np.zeros(len(df_m6_final)), color='white')

        aplicar_anti_overlap(ax_m6_1, df_m6_final['Suma_Imp'].max(), 2.2)
        
        # Etiquetas de volumen
        for i in range(len(x_pos_m6)):
            iv_val = df_m6_final['Suma_Imp'].iloc[i]
            dv_val = df_m6_final['HH_Disponibles'].iloc[i]
            if iv_val > 0: 
                ax_m6_1.annotate(f"Imp: {int(iv_val)}\nDisp: {int(dv_val)}", (i, iv_val + (ax_m6_1.get_ylim()[1]*0.05)), ha='center', bbox=bbox_yellow, fontweight='bold', zorder=10)

        # Línea de Incidencia %
        ax_m6_2 = ax_m6_1.twinx()
        ax_m6_2.plot(x_pos_m6, df_m6_final['%_Incidencia'], color='red', marker='o', markersize=12, linewidth=6, path_effects=outline_white, label='% Incidencia', zorder=5)
        
        # Meta 15%
        ax_m6_2.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax_m6_2.text(x_pos_m6[0], 16, 'META = 15%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val_inc in enumerate(df_m6_final['%_Incidencia']):
            ax_m6_2.annotate(f"{val_inc:.1f}%", (x_pos_m6[i], val_inc + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=outline_white, zorder=10)

        ax_m6_1.set_xticks(x_pos_m6)
        ax_m6_1.set_xticklabels(df_m6_final['C_Aux'], fontsize=14, fontweight='bold')
        ax_m6_2.set_ylim(0, max(30, df_m6_final['%_Incidencia'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia para este sector.")
