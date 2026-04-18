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
# 1. CONFIGURACIÓN DE LA ESTRUCTURA DE PÁGINA
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
    """Dibuja la pantalla institucional de acceso con el diseño oficial."""
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_v1, col_login, col_v2 = st.columns([1, 1.8, 1])
    
    with col_login:
        # Franja Azul Ombú superior
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGOTIPO PEQUEÑO CENTRADO (160px)
        inner_l, inner_c, inner_r = st.columns([1, 1, 1])
        with inner_c:
            try:
                st.image("LOGO OMBÚ.jpg", width=160)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        # NOMBRE OFICIAL DE LA COMPAÑÍA
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Formulario de entrada
        with st.form("form_seguridad_acceso"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            usuario_input = st.text_input("Usuario Corporativo")
            clave_input = st.text_input("Contraseña", type="password")
            
            btn_entrar = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if btn_entrar:
                if usuario_input in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[usuario_input] == clave_input:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos.")

# Bloqueo si no hay login exitoso
if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES AVANZADOS (CSS Y PANEL FIJO)
# =========================================================================
st.markdown("""
<style>
    /* Ocultar elementos nativos de Streamlit */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* PANEL DE FILTROS STICKY (FIJO EN LA PARTE SUPERIOR) */
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

# Configuración Maestra de Matplotlib (Fuentes en Negrita para Gerencia)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Efectos de contorno para etiquetas de alta legibilidad
c_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
c_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cuadros decorativos (Bboxes)
bbox_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
bbox_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_oro = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. MOTOR INTELIGENTE DE CRUCE DE DATOS (MOTOR FUZZY)
# =========================================================================
def set_escala_superior(ax_plot, valor_max_eje, multiplicador=2.6):
    """Deja espacio en el gráfico arriba para que las etiquetas no se pisen."""
    if valor_max_eje > 0: 
        ax_plot.set_ylim(0, valor_max_eje * multiplicador)
    else: 
        ax_plot.set_ylim(0, 100)

def dibujar_guias_meses(ax_plot, n_total_fechas):
    """Líneas verticales grises para separar visualmente los meses."""
    for i in range(n_total_fechas):
        ax_plot.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_labels_filtros(lista_sel, string_default):
    """Prepara el texto de filtros activos para los títulos."""
    if not lista_sel: 
        return string_default
    if len(lista_sel) > 2: 
        return f"Varios ({len(lista_sel)})"
    return " + ".join(lista_sel)

def motor_fuzzy_match(seleccion_usuario, valor_columna_excel):
    """
    MOTOR DE CRUCE INDUSTRIAL:
    Busca coincidencias entre Eficiencias e Improductivas aunque los textos varíen.
    Indispensable para recuperar las HH de estaciones que tienen nombres distintos.
    """
    if pd.isna(valor_columna_excel) or pd.isna(seleccion_usuario): 
        return False
    
    # 1. Normalización profunda de strings
    txt1 = str(seleccion_usuario).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    txt2 = str(valor_columna_excel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # 2. Eliminación de caracteres no alfanuméricos
    limpio1 = re.sub(r'[^A-Z0-9]', '', txt1)
    limpio2 = re.sub(r'[^A-Z0-9]', '', txt2)
    
    if not limpio1 or not limpio2: 
        return False
        
    # 3. Coincidencia directa de texto
    if limpio1 in limpio2 or limpio2 in limpio1: 
        return True
    
    # 4. Coincidencia por códigos numéricos (códigos de estación de 3+ dígitos)
    n_sel = set(re.findall(r'\d{3,}', txt1))
    n_exc = set(re.findall(r'\d{3,}', txt2))
    if n_sel and n_exc and n_sel.intersection(n_exc): 
        return True
        
    # 5. Búsqueda por palabras clave raíz
    pal_sel = set(re.findall(r'[A-Z]{4,}', txt1))
    pal_exc = set(re.findall(r'[A-Z]{4,}', txt2))
    
    exclusiones = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1_final = pal_sel - exclusiones
    v2_final = pal_exc - exclusiones
    
    for p in v1_final:
        if any(p in x for x in v2_final): 
            return True
                
    return False

# =========================================================================
# 5. HEADER Y LOGOUT
# =========================================================================
col_logo_h, col_tit_h, col_out_h = st.columns([1, 3, 1])

with col_logo_h:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with col_tit_h:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Control de Gestión Productiva</p>", unsafe_allow_html=True)

with col_out_h:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA DE DATOS (LECTURA DE EXCEL)
# =========================================================================
try:
    # Lectura de bases
    df_ef_maestra = pd.read_excel("eficiencias.xlsx")
    df_im_maestra = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de encabezados
    df_ef_maestra.columns = df_ef_maestra.columns.str.strip()
    df_im_maestra.columns = [str(c).strip().upper() for c in df_im_maestra.columns]
    
    # Auto-corrector inteligente de columnas para Improductivas
    if 'TIPO_PARADA' not in df_im_maestra.columns:
        c_alt_m = next((c for c in df_im_maestra.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if c_alt_m: df_im_maestra.rename(columns={c_alt_m: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im_maestra.columns:
        c_alt_h = next((c for c in df_im_maestra.columns if 'HH' in c and 'IMP' in c), None)
        if c_alt_h: df_im_maestra.rename(columns={c_alt_h: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im_maestra.columns:
        c_alt_f = next((c for c in df_im_maestra.columns if 'FECHA' in c), None)
        if c_alt_f: df_im_maestra.rename(columns={c_alt_f: 'FECHA'}, inplace=True)
    
    # Estandarización de Fechas (Todo al primer día del mes)
    df_ef_maestra['Fecha'] = pd.to_datetime(df_ef_maestra['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im_maestra['FECHA'] = pd.to_datetime(df_im_maestra['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación de puestos de salida
    df_ef_maestra['Es_Ultimo_Puesto'] = df_ef_maestra['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Etiquetas para filtros temporales
    df_ef_maestra['Etiqueta_Mes'] = df_ef_maestra['Fecha'].dt.strftime('%b-%Y')
    df_im_maestra['Etiqueta_Mes'] = df_im_maestra['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e_error_carga:
    st.error(f"Error crítico en la carga de archivos: {e_error_carga}")
    st.stop()

# =========================================================================
# 7. PANEL DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    f1, f2, f3, f4 = st.columns(4)
    
    with f1: 
        list_plantas = list(df_ef_maestra['Planta'].dropna().unique())
        sel_planta = st.multiselect("🏭 Planta", list_plantas, placeholder="Todas")
        
    df_temp_l = df_ef_maestra[df_ef_maestra['Planta'].isin(sel_planta)] if sel_planta else df_ef_maestra
    
    with f2: 
        list_lineas = list(df_temp_l['Linea'].dropna().unique())
        sel_linea = st.multiselect("⚙️ Línea", list_lineas, placeholder="Todas")
        
    df_temp_p = df_temp_l[df_temp_l['Linea'].isin(sel_linea)] if sel_linea else df_temp_l
    
    with f3: 
        list_puestos = list(df_temp_p['Puesto_Trabajo'].dropna().unique())
        sel_puesto = st.multiselect("🛠️ Puesto de Trabajo", list_puestos, placeholder="Todos")
        
    with f4: 
        list_meses = ["🎯 Acumulado YTD"] + list(df_ef_maestra['Etiqueta_Mes'].unique())
        sel_mes = st.multiselect("📅 Mes", list_meses, placeholder="Todos")

# =========================================================================
# 8. PROCESAMIENTO FINAL DE FILTROS (CONSTRUCCIÓN DE VARIABLES)
# =========================================================================
df_ef_f = df_ef_maestra.copy()
df_im_f = df_im_maestra.copy()

# 8.1 Aplicar filtros a Eficiencias
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
if sel_linea: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]
if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_ef_f = df_ef_f[df_ef_f['Etiqueta_Mes'].isin(sel_mes)]

# 8.2 Aplicar filtros a Improductivas (Usando el Motor Fuzzy Inteligente)
if sel_planta:
    mask_pl = df_im_f.iloc[:,0].apply(lambda x: any(motor_fuzzy_match(p, x) for p in sel_planta))
    df_im_f = df_im_f[mask_pl]

if sel_linea:
    c_l_idx = next((c for c in df_im_f.columns if 'LINEA' in c), df_im_f.columns[1])
    mask_li = df_im_f[c_l_idx].apply(lambda x: any(motor_fuzzy_match(l, x) for l in sel_linea))
    df_im_f = df_im_f[mask_li]

if sel_puesto:
    c_p_idx = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2])
    mask_ps = df_im_f[c_p_idx].apply(lambda x: any(motor_fuzzy_match(ps, x) for ps in sel_puesto))
    df_im_f = df_im_f[mask_ps]

if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_im_f = df_im_f[df_im_f['Etiqueta_Mes'].isin(sel_mes)]

# Texto para encabezados informativos en gráficos
txt_h_info = f"Filtros: {formatear_labels_filtros(sel_planta, 'Todas')} | {formatear_labels_filtros(sel_linea, 'Todas')} | {formatear_labels_filtros(sel_puesto, 'Todos')}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (PRODUCTIVIDAD)
# =========================================================================
col_izq_1, col_der_1 = st.columns(2)

with col_izq_1:
    st.header("1. EFICIENCIA REAL")
    # ESPECIFICACIÓN DE FÓRMULA SOLICITADA POR EL USUARIO
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    # Lógica de dibujo: Si no hay puesto seleccionado, mostramos solo salida (SI)
    df_m1_build = df_ef_f.copy() if sel_puesto else df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_build.empty:
        # Agrupación temporal
        agrup_1 = df_m1_build.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo del KPI
        agrup_1['Ef_Real_Pct'] = (agrup_1['HH_STD_TOTAL'] / agrup_1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Preparación de gráfico con dos ejes
        fig1, ax1_bars = plt.subplots(figsize=(14, 10))
        ax1_line = ax1_bars.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices_1 = np.arange(len(agrup_1))
        ancho_1 = 0.35
        
        # Dibujo de Barras de Volumen Horario
        bar_s_1 = ax1_bars.bar(x_indices_1 - ancho_1/2, agrup_1['HH_STD_TOTAL'], ancho_1, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_d_1 = ax1_bars.bar(x_indices_1 + ancho_1/2, agrup_1['HH_Disponibles'], ancho_1, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_escala_superior(ax1_bars, agrup_1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas de valores sobre las barras
        ax1_bars.bar_label(bar_s_1, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax1_bars.bar_label(bar_d_1, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        dibujar_guias_meses(ax1_bars, len(x_indices_1))

        # Texto vertical de Unidades Producidas (Dentro de la barra)
        for i, bar in enumerate(bar_s_1):
            val_und = int(agrup_1['Cant._Prod._A1'].iloc[i])
            if val_und > 0: 
                ax1_bars.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{val_und} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=c_negro, zorder=4)

        # Línea de Eficiencia %
        ax1_line.plot(x_indices_1, agrup_1['Ef_Real_Pct'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=c_blanco, label='% Eficiencia Real', zorder=5)
        
        # Línea de Meta 85%
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_indices_1[0], 86, 'META = 85%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_line.set_ylim(0, max(120, agrup_1['Ef_Real_Pct'].max()*1.8))
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación de Porcentaje sobre la línea
        for i, val in enumerate(agrup_1['Ef_Real_Pct']):
            ax1_line.annotate(f"{val:.1f}%", (x_indices_1[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        # Configuración de Ejes
        ax1_bars.set_xticks(x_indices_1)
        ax1_bars.set_xticklabels(agrup_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        
        ax1_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin datos para graficar Eficiencia Real con los filtros aplicados.")

with col_der_1:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_build.empty:
        # Agrupación temporal
        agrup_2 = df_m1_build.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        agrup_2['Ef_Prod_Pct'] = (agrup_2['HH_STD_TOTAL'] / agrup_2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Construcción del gráfico 2
        fig2, ax2_bars = plt.subplots(figsize=(14, 10))
        ax2_line = ax2_bars.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices_2 = np.arange(len(agrup_2))
        
        # Barras comparativas
        b_s_2 = ax2_bars.bar(x_indices_2 - 0.35/2, agrup_2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_p_2 = ax2_bars.bar(x_indices_2 + 0.35/2, agrup_2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_superior(ax2_bars, max(agrup_2['HH_STD_TOTAL'].max(), agrup_2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax2_bars.bar_label(b_s_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax2_bars.bar_label(b_p_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        dibujar_guias_meses(ax2_bars, len(x_indices_2))

        # Línea de Eficiencia Productiva (%)
        ax2_line.plot(x_indices_2, agrup_2['Ef_Prod_Pct'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=c_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta Corporativa 100%
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_indices_2[0], 101, 'META = 100%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_line.set_ylim(0, max(150, agrup_2['Ef_Prod_Pct'].max()*1.8))
        ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación en línea
        for i, val in enumerate(agrup_2['Ef_Prod_Pct']):
            ax2_line.annotate(f"{val:.1f}%", (x_indices_2[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax2_bars.set_xticks(x_indices_2)
        ax2_bars.set_xticklabels(agrup_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ Sin datos para análisis de Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (ANÁLISIS DE BRECHA Y COSTOS)
# =========================================================================
col_izq_2, col_der_2 = st.columns(2)

with col_izq_2:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Horas Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        c_prod_fix = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        
        agrup_3 = df_ef_f.groupby('Fecha').agg({
            c_prod_fix: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        # Suma de lo que se declaró (Prod + Impro) para comparar con Disponibles
        agrup_3['Total_Declaradas'] = agrup_3[c_prod_fix] + agrup_3['HH_Improductivas']
        
        # Inicio Gráfico 3
        fig3, ax3_base = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_ticks_3 = np.arange(len(agrup_3))
        
        # Barras Apiladas (Horas Productivas en base, Improductivas arriba)
        ax3_base.bar(x_ticks_3, agrup_3[c_prod_fix], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3_base.bar(x_ticks_3, agrup_3['HH_Improductivas'], bottom=agrup_3[c_prod_fix], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad Máxima
        ax3_base.plot(x_ticks_3, agrup_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=c_blanco, label='HH DISPONIBLES', zorder=5)
        
        set_escala_superior(ax3_base, agrup_3['HH_Disponibles'].max(), 2.6)
        dibujar_guias_meses(ax3_base, len(x_ticks_3))

        # Visualización técnica del GAP (Brecha no declarada)
        for i in range(len(x_ticks_3)):
            h_dis = agrup_3['HH_Disponibles'].iloc[i]
            h_dec = agrup_3['Total_Declaradas'].iloc[i]
            v_brecha = h_dis - h_dec
            
            # Flecha/Línea vertical del GAP
            ax3_base.plot([i, i], [h_dec, h_dis], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Cartelera del GAP
            ax3_base.annotate(f"GAP:\n{int(v_brecha)}", (i, h_dec + 5), color='firebrick', bbox=bbox_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Etiqueta de Disponibilidad Final
            ax3_base.annotate(f"{int(h_dis)}", (i, h_dis + (ax3_base.get_ylim()[1]*0.08)), color='black', bbox=bbox_blanco, ha='center', fontweight='bold', zorder=10)

        ax3_base.set_xticks(x_ticks_3)
        ax3_base.set_xticklabels(agrup_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3_base.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ Sin datos disponibles para el análisis de GAP.")

with col_der_2:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización económica de la ineficiencia (HH Improductivas)</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        agrup_4 = df_ef_f.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        # Inicio Gráfico 4
        fig4, ax4_left = plt.subplots(figsize=(14, 10))
        ax4_right = ax4_left.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_ticks_4 = np.arange(len(agrup_4))
        
        # Barras de horas perdidas
        bar_lost_m4 = ax4_left.bar(x_ticks_4, agrup_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_left.bar_label(bar_lost_m4, padding=4, color='black', fontweight='bold', path_effects=c_blanco, zorder=4)
        
        set_escala_superior(ax4_left, agrup_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Costo ($) en el eje derecho
        ax4_right.plot(x_ticks_4, agrup_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=c_blanco, label='COSTO ARS', zorder=5)
        
        ax4_right.set_ylim(0, max(1000, agrup_4['Costo_Improd._$'].max() * 1.8))
        ax4_right.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_right.get_yticks()], fontweight='bold')

        # Cartel de Resumen Gerencial de Costo
        tot_pesos = agrup_4['Costo_Improd._$'].sum()
        ax4_left.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${tot_pesos:,.0f}", transform=ax4_left.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=bbox_oro, weight='bold', zorder=10)

        for i, val in enumerate(agrup_4['Costo_Improd._$']):
            ax4_right.annotate(f"${val:,.0f}", (x_ticks_4[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax4_left.set_xticks(x_ticks_4)
        ax4_left.set_xticklabels(agrup_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ No hay datos económicos para visualizar.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (ANÁLISIS CAUSA RAÍZ)
# =========================================================================
col_izq_3, col_der_3 = st.columns(2)

with col_izq_3:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        # Procesamiento de Pareto
        agrup_5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de meses para promedio
        n_meses_u = df_im_f['FECHA'].nunique()
        divisor_p = n_meses_u if n_meses_u > 0 else 1
        
        agrup_5['Promedio_M'] = agrup_5['HH_IMPRODUCTIVAS'] / divisor_p
        agrup_5 = agrup_5.sort_values(by='Promedio_M', ascending=False)
        agrup_5['Pct_Acumulado'] = (agrup_5['Promedio_M'].cumsum() / agrup_5['Promedio_M'].sum()) * 100

        # Inicio Gráfico 5
        fig5, ax5_left = plt.subplots(figsize=(14, 10))
        ax5_right = ax5_left.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_pos_5 = np.arange(len(agrup_5))
        
        # Barras Pareto
        bar_p5 = ax5_left.bar(x_pos_5, agrup_5['Promedio_M'], color='maroon', edgecolor='white', zorder=2)
        set_escala_superior(ax5_left, agrup_5['Promedio_M'].max(), 2.8)
        ax5_left.bar_label(bar_p5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva de Lorenz
        ax5_right.plot(x_pos_5, agrup_5['Pct_Acumulado'], color='red', marker='D', markersize=10, linewidth=4, path_effects=c_blanco, zorder=5)
        ax5_right.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax5_right.set_ylim(0, 200)
        ax5_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Wrap de texto para causas largas
        lbls_5 = [textwrap.fill(str(txt), 12) for txt in agrup_5['TIPO_PARADA']]
        ax5_left.set_xticks(x_pos_5)
        ax5_left.set_xticklabels(lbls_5, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val in enumerate(agrup_5['Pct_Acumulado']):
            ax5_right.annotate(f"{val:.1f}%", (x_pos_5[i], val + 4), color='white', bbox=bbox_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        total_m_sum = agrup_5['Promedio_M'].sum()
        ax5_left.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{total_m_sum:.1f} HH", transform=ax5_left.transAxes, bbox=bbox_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl_fin = agrup_5.copy()
        total_p_final_hs = df_tbl_fin['HH_IMPRODUCTIVAS'].sum()
        df_tbl_fin['% sobre Selección'] = (df_tbl_fin['HH_IMPRODUCTIVAS'] / total_p_final_hs) * 100
        
        # INYECCIÓN FILA DE TOTAL DEFINITIVA
        f_tot_final = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [total_p_final_hs], 
            'Promedio_M': [df_tbl_fin['Promedio_M'].sum()],
            'Pct_Acumulado': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tbl_fin = pd.concat([df_tbl_fin, f_tot_final], ignore_index=True)
        
        st.dataframe(
            df_tbl_fin.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Botón de Descarga
        csv_master = df_tbl_fin.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=csv_master, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col_der_3:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        # Cruce de datos por Mes para Stacked Chart
        df_ef_f['Mes_Cruce'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp_6 = df_ef_f.groupby('Mes_Cruce', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['Mes_Cruce'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            # Tabla dinámica para abrir causas por mes
            pivote_m6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='Mes_Cruce', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_m6_final = pd.merge(ag_disp_6, pivote_m6, on='Mes_Cruce', how='left').fillna(0)
            list_causas_m6 = [c for c in df_m6_final.columns if c not in ['HH_Disponibles', 'Mes_Cruce']]
        else:
            df_m6_final = ag_disp_6.copy()
            list_causas_m6 = []
            
        # Indicadores finales de incidencia
        df_m6_final['Suma_I_Hs'] = df_m6_final[list_causas_m6].sum(axis=1) if list_causas_m6 else 0
        df_m6_final['Incidencia_%'] = (df_m6_final['Suma_I_Hs'] / df_m6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden Cronológico CORREGIDO (Sin error de Syntax)
        df_m6_final['Orden_Date'] = pd.to_datetime(df_m6_final['Mes_Cruce'] + '-01')
        df_m6_final = df_m6_final.sort_values(by='Orden_Date')

        # Inicio Gráfico 6
        fig6, ax6_base = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_6 = np.arange(len(df_m6_final))
        
        # Barras Apiladas por Causa
        if list_causas_m6:
            piso_6 = np.zeros(len(df_m6_final))
            paleta_6 = plt.cm.tab20.colors
            for idx_c, nom_c in enumerate(list_causas_m6):
                vals_c = df_m6_final[nom_c].values
                ax6_base.bar(pos_6, vals_c, bottom=piso_6, label=nom_c, color=paleta_6[idx_c % 20], edgecolor='white', zorder=2)
                piso_6 += vals_c
        else:
            ax6_base.bar(pos_6, np.zeros(len(df_m6_final)), color='white')

        set_escala_superior(ax6_base, df_m6_final['Suma_I_Hs'].max(), 2.2)
        
        # Anotación volumen vs disponibilidad
        for i in range(len(pos_6)):
            v_i = df_m6_final['Suma_I_Hs'].iloc[i]
            v_d = df_m6_final['HH_Disponibles'].iloc[i]
            if v_i > 0: 
                ax6_base.annotate(f"Imp: {int(v_i)}\nDisp: {int(v_d)}", (i, v_i + (ax6_base.get_ylim()[1]*0.05)), ha='center', bbox=bbox_oro, fontweight='bold', zorder=10)

        # Línea de Porcentaje de Incidencia
        ax6_twin = ax6_base.twinx()
        ax6_twin.plot(pos_6, df_m6_final['Incidencia_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=c_blanco, label='% Incidencia', zorder=5)
        
        # Meta 15%
        ax6_twin.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_twin.text(pos_6[0], 16, 'META = 15%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(df_m6_final['Incidencia_%']):
            ax6_twin.annotate(f"{val:.1f}%", (pos_6[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=c_blanco, zorder=10)

        ax6_base.set_xticks(pos_6)
        ax6_base.set_xticklabels(df_m6_final['Mes_Cruce'], fontsize=14, fontweight='bold')
        ax6_twin.set_ylim(0, max(30, df_m6_final['Incidencia_%'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia para este sector.")
