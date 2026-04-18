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
# 2. SISTEMA DE ACCESO Y SEGURIDAD (LLAVE MAESTRA)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    """Dibuja la puerta de acceso institucional con el diseño solicitado."""
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_izq, col_login, col_der = st.columns([1, 1.8, 1])
    
    with col_login:
        # Franja Azul Ombú superior
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGOTIPO PEQUEÑO (160px)
        s_col1, s_col2, s_col3 = st.columns([1, 1, 1])
        with s_col2:
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
        
        # Formulario de acceso
        with st.form("form_acceso_restringido"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            u_digitado = st.text_input("Usuario Corporativo")
            p_digitado = st.text_input("Contraseña", type="password")
            
            btn_entrar = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if btn_entrar:
                if u_digitado in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_digitado] == p_digitado:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos.")

# Bloqueo de seguridad preventivo
if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES AVANZADOS (CSS Y PANEL FIJO)
# =========================================================================
st.markdown("""
<style>
    /* Ocultar elementos nativos para mayor limpieza visual */
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
        padding-top: 15px !important;
        padding-bottom: 15px !important;
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

# Efectos de contorno para etiquetas de alta visibilidad
c_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
c_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cuadros decorativos (Bboxes)
box_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
box_oro = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
box_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
box_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. FUNCIONES DE LÓGICA Y MOTOR DE CRUCE INTELIGENTE
# =========================================================================
def set_escala_superior(ax_plot, valor_maximo, factor=2.6):
    """Genera espacio arriba para que las etiquetas no se pisen con el título."""
    if valor_maximo > 0: 
        ax_plot.set_ylim(0, valor_maximo * factor)
    else: 
        ax_plot.set_ylim(0, 100)

def dibujar_guia_meses(ax_plot, cantidad):
    """Líneas verticales grises para separar visualmente los periodos."""
    for i in range(cantidad):
        ax_plot.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def motor_fuzzy_cruce(sel_actual, val_excel_impro):
    """
    MOTOR DE CRUCE INDUSTRIAL:
    Busca coincidencias entre bases aunque los textos no sean idénticos.
    Ejemplo: '475-CARROZADO' con 'SECTOR CARROZADO'.
    """
    if pd.isna(val_excel_impro) or pd.isna(sel_actual): 
        return False
    
    # 1. Normalización profunda
    s1 = str(sel_actual).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(val_excel_impro).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # 2. Solo alfanumérico
    l1 = re.sub(r'[^A-Z0-9]', '', s1)
    l2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not l1 or not l2: 
        return False
        
    # 3. Coincidencia directa
    if l1 in l2 or l2 in l1: 
        return True
    
    # 4. Coincidencia por código de estación (3+ dígitos)
    numeros1 = set(re.findall(r'\d{3,}', s1))
    numeros2 = set(re.findall(r'\d{3,}', s2))
    if numeros1 and numeros2 and numeros1.intersection(numeros2): 
        return True
        
    # 5. Coincidencia por raíces de palabras largas
    palabras1 = set(re.findall(r'[A-Z]{4,}', s1))
    palabras2 = set(re.findall(r'[A-Z]{4,}', s2))
    
    exclusiones = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1_limpio = palabras1 - exclusiones
    v2_limpio = palabras2 - exclusiones
    
    for word in v1_limpio:
        if any(word in x for x in v2_limpio): 
            return True
                
    return False

# =========================================================================
# 5. HEADER PRINCIPAL Y SALIDA
# =========================================================================
col_logo_m, col_tit_m, col_out_m = st.columns([1, 3, 1])

with col_logo_m:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with col_tit_m:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Control de Gestión Productiva</p>", unsafe_allow_html=True)

with col_out_m:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA DE DATOS (BASE EXCEL)
# =========================================================================
try:
    # Carga de archivos
    df_ef_base = pd.read_excel("eficiencias.xlsx")
    df_im_base = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de encabezados
    df_ef_base.columns = df_ef_base.columns.str.strip()
    df_im_base.columns = [str(c).strip().upper() for c in df_im_base.columns]
    
    # Auto-corrector de columnas para la base de improductivas
    if 'TIPO_PARADA' not in df_im_base.columns:
        c_alt_mot = next((c for c in df_im_base.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if c_alt_mot: df_im_base.rename(columns={c_alt_mot: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im_base.columns:
        c_alt_h = next((c for c in df_im_base.columns if 'HH' in c and 'IMP' in c), None)
        if c_alt_h: df_im_base.rename(columns={c_alt_h: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im_base.columns:
        c_alt_f = next((c for c in df_im_base.columns if 'FECHA' in c), None)
        if c_alt_f: df_im_base.rename(columns={c_alt_f: 'FECHA'}, inplace=True)
    
    # Estandarización de Fechas (Todo al día 1 del mes)
    df_ef_base['Fecha'] = pd.to_datetime(df_ef_base['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im_base['FECHA'] = pd.to_datetime(df_im_base['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Identificación técnica de puestos
    df_ef_base['Es_Ultimo_Puesto'] = df_ef_base['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Etiquetas legibles para filtros
    df_ef_base['Mes_Label'] = df_ef_base['Fecha'].dt.strftime('%b-%Y')
    df_im_base['Mes_Label'] = df_im_base['FECHA'].dt.strftime('%b-%Y')
    
except Exception as err_general:
    st.error(f"Error crítico cargando bases de datos: {err_general}")
    st.stop()

# =========================================================================
# 7. PANEL DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    fl_1, fl_2, fl_3, fl_4 = st.columns(4)
    
    with fl_1: 
        op_plantas = list(df_ef_base['Planta'].dropna().unique())
        sel_planta = st.multiselect("🏭 Planta", op_plantas, placeholder="Todas")
        
    df_tmp_lineas = df_ef_base[df_ef_base['Planta'].isin(sel_planta)] if sel_planta else df_ef_base
    
    with fl_2: 
        op_lineas = list(df_tmp_lineas['Linea'].dropna().unique())
        sel_linea = st.multiselect("⚙️ Línea", op_lineas, placeholder="Todas")
        
    df_tmp_puestos = df_tmp_lineas[df_tmp_lineas['Linea'].isin(sel_linea)] if sel_linea else df_tmp_lineas
    
    with fl_3: 
        op_puestos = list(df_tmp_puestos['Puesto_Trabajo'].dropna().unique())
        sel_puesto = st.multiselect("🛠️ Puesto de Trabajo", op_puestos, placeholder="Todos")
        
    with fl_4: 
        op_meses = ["🎯 Acumulado YTD"] + list(df_ef_base['Mes_Label'].unique())
        sel_mes = st.multiselect("📅 Mes", op_meses, placeholder="Todos")

# =========================================================================
# 8. PROCESAMIENTO DE DATOS FILTRADOS (BLINDAJE DE VARIABLES)
# =========================================================================
# Creación de dataframes de trabajo filtrados
df_ef_f = df_ef_base.copy()
df_im_f = df_im_base.copy()

# 8.1 Filtrar Eficiencias
if sel_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(sel_planta)]
if sel_linea: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(sel_linea)]
if sel_puesto: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(sel_puesto)]
if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_ef_f = df_ef_f[df_ef_f['Mes_Label'].isin(sel_mes)]

# 8.2 Filtrar Improductivas (Usando el Motor Fuzzy Inteligente)
if sel_planta:
    mask_planta_im = df_im_f.iloc[:,0].apply(lambda x: any(motor_fuzzy_cruce(p, x) for p in sel_planta))
    df_im_f = df_im_f[mask_planta_im]

if sel_linea:
    c_linea_idx = next((c for c in df_im_f.columns if 'LINEA' in c), df_im_f.columns[1])
    mask_linea_im = df_im_f[c_linea_idx].apply(lambda x: any(motor_fuzzy_cruce(l, x) for l in sel_linea))
    df_im_f = df_im_f[mask_linea_im]

if sel_puesto:
    c_puesto_idx = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2])
    mask_puesto_im = df_im_f[c_puesto_idx].apply(lambda x: any(motor_fuzzy_cruce(ps, x) for ps in sel_puesto))
    df_im_f = df_im_f[mask_puesto_im]

if sel_mes and "🎯 Acumulado YTD" not in sel_mes:
    df_im_f = df_im_f[df_im_f['Mes_Label'].isin(sel_mes)]

# Texto para encabezados dinámicos
txt_h_info = f"Escenario: {sel_planta if sel_planta else 'Todas'} > {sel_linea if sel_linea else 'Todas'} > {sel_puesto if sel_puesto else 'Todos'}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (EFICIENCIAS)
# =========================================================================
col1, col2 = st.columns(2)

with col1:
    st.header("1. EFICIENCIA REAL")
    # ESPECIFICACIÓN DE FÓRMULA SOLICITADA
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    # Lógica de dibujo: Si no hay puesto elegido, mostrar solo 'SI' (estaciones terminales)
    df_m1_build = df_ef_f.copy() if sel_puesto else df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_build.empty:
        # Agrupación temporal
        ag_1 = df_m1_build.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo de indicador
        ag_1['Ef_Real'] = (ag_1['HH_STD_TOTAL'] / ag_1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Configuración de gráfico
        fig1, ax1_bars = plt.subplots(figsize=(14, 10))
        ax1_line = ax1_bars.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_ticks_1 = np.arange(len(ag_1))
        
        # Dibujo de Barras de Volumen
        bar_s_1 = ax1_bars.bar(x_ticks_1 - 0.17, ag_1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_d_1 = ax1_bars.bar(x_ticks_1 + 0.17, ag_1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_escala_superior(ax1_bars, ag_1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas de valores en las barras
        ax1_bars.bar_label(bar_s_1, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax1_bars.bar_label(bar_d_1, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        dibujar_guia_meses(ax1_bars, len(x_ticks_1))

        # Texto vertical de Unidades (Dentro de la barra)
        for i, bar in enumerate(bar_s_1):
            val_und = int(ag_1['Cant._Prod._A1'].iloc[i])
            if val_und > 0: 
                ax1_bars.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{val_und} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=c_negro, zorder=4)

        # Línea de Eficiencia %
        ax1_line.plot(x_ticks_1, ag_1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=c_blanco, label='% Eficiencia Real', zorder=5)
        
        # Meta 85%
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_ticks_1[0], 86, 'META = 85%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_line.set_ylim(0, max(120, ag_1['Ef_Real'].max()*1.8))
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación de Porcentaje
        for i, val in enumerate(ag_1['Ef_Real']):
            ax1_line.annotate(f"{val:.1f}%", (x_ticks_1[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax1_bars.set_xticks(x_ticks_1)
        ax1_bars.set_xticklabels(ag_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin datos para graficar Eficiencia Real.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_build.empty:
        # Agrupación temporal
        ag_2 = df_m1_build.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        ag_2['Ef_Prod'] = (ag_2['HH_STD_TOTAL'] / ag_2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Configuración de gráfico
        fig2, ax2_bars = plt.subplots(figsize=(14, 10))
        ax2_line = ax2_bars.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_ticks_2 = np.arange(len(ag_2))
        
        # Barras comparativas
        b_s_2 = ax2_bars.bar(x_ticks_2 - 0.17, ag_2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_p_2 = ax2_bars.bar(x_ticks_2 + 0.17, ag_2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_superior(ax2_bars, max(ag_2['HH_STD_TOTAL'].max(), ag_2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax2_bars.bar_label(b_s_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        ax2_bars.bar_label(b_p_2, padding=4, color='black', fontweight='bold', path_effects=c_blanco, fmt='%.0f', zorder=3)
        
        dibujar_guia_meses(ax2_bars, len(x_ticks_2))

        # Línea de Eficiencia Productiva
        ax2_line.plot(x_ticks_2, ag_2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=c_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta 100%
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_ticks_2[0], 101, 'META = 100%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_line.set_ylim(0, max(150, ag_2['Ef_Prod'].max()*1.8))
        ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación en línea
        for i, val in enumerate(ag_2['Ef_Prod']):
            ax2_line.annotate(f"{val:.1f}%", (x_ticks_2[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax2_bars.set_xticks(x_ticks_2)
        ax2_bars.set_xticklabels(ag_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ Sin datos para Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (ANÁLISIS DE BRECHA Y COSTOS)
# =========================================================================
col3, col4 = st.columns(2)

with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Horas Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        # CONTROL DE VARIABLE CRÍTICA: Aseguramos nombre exacto de columna
        c_prod_maestra = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        
        ag_3 = df_ef_f.groupby('Fecha').agg({
            c_prod_maestra: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        ag_3['Suma_Declaradas'] = ag_3[c_prod_maestra] + ag_3['HH_Improductivas']
        
        # Gráfico GAP
        fig3, ax3_base = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_ticks_3 = np.arange(len(ag_3))
        
        # Barras Apiladas (Prod + Impro)
        ax3_base.bar(x_ticks_3, ag_3[c_prod_maestra], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3_base.bar(x_ticks_3, ag_3['HH_Improductivas'], bottom=ag_3[c_prod_maestra], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad
        ax3_base.plot(x_ticks_3, ag_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=c_blanco, label='HH DISPONIBLES', zorder=5)
        
        set_escala_superior(ax3_base, ag_3['HH_Disponibles'].max(), 2.6)
        dibujar_guia_meses(ax3_base, len(x_ticks_3))

        # Visualización del GAP
        for i in range(len(x_ticks_3)):
            val_disp = ag_3['HH_Disponibles'].iloc[i]
            val_decl = ag_3['Suma_Declaradas'].iloc[i]
            val_gap = val_disp - val_decl
            
            # Flecha vertical
            ax3_base.plot([i, i], [val_decl, val_disp], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            # Etiqueta GAP
            ax3_base.annotate(f"GAP:\n{int(val_gap)}", (i, val_decl + 5), color='firebrick', bbox=box_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            # Valor Disponibilidad Arriba
            ax3_base.annotate(f"{int(val_disp)}", (i, val_disp + (ax3_base.get_ylim()[1]*0.08)), color='black', bbox=box_blanco, ha='center', fontweight='bold', zorder=10)

        ax3_base.set_xticks(x_ticks_3)
        ax3_base.set_xticklabels(ag_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3_base.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ Sin datos para análisis de GAP.")

with col4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización del impacto de las improductividades</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag_4 = df_ef_f.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        # Gráfico Costos
        fig4, ax4_left = plt.subplots(figsize=(14, 10))
        ax4_right = ax4_left.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_ticks_4 = np.arange(len(ag_4))
        
        # Barras Horas Perdidas
        bar_imp_4 = ax4_left.bar(x_ticks_4, ag_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_left.bar_label(bar_imp_4, padding=4, color='black', fontweight='bold', path_effects=c_blanco, zorder=4)
        
        set_escala_superior(ax4_left, ag_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Pesos (Eje secundario)
        ax4_right.plot(x_ticks_4, ag_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=c_blanco, label='COSTO ARS', zorder=5)
        
        ax4_right.set_ylim(0, max(1000, ag_4['Costo_Improd._$'].max() * 1.8))
        ax4_right.set_yticklabels([f'${int(val/1000000)}M' for val in ax4_right.get_yticks()], fontweight='bold')

        # Cartelera de Costo Acumulado
        total_p_pesos = ag_4['Costo_Improd._$'].sum()
        total_p_hs = ag_4['HH_Improductivas'].sum()
        ax4_left.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${total_p_pesos:,.0f}", transform=ax4_left.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=box_oro, weight='bold', zorder=10)

        for i, val in enumerate(ag_4['Costo_Improd._$']):
            ax4_right.annotate(f"${val:,.0f}", (x_ticks_4[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax4_left.set_xticks(x_ticks_4)
        ax4_left.set_xticklabels(ag_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ Sin datos de costos económicos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (PARETO E INCIDENCIA)
# =========================================================================
col5, col6 = st.columns(2)

with col5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de causas de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        # Lógica Pareto
        ag_5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de meses para promedios
        n_m_unicos = df_im_f['FECHA'].nunique()
        divisor_p = n_m_unicos if n_m_unicos > 0 else 1
        
        ag_5['Promedio_Mensual'] = ag_5['HH_IMPRODUCTIVAS'] / divisor_p
        ag_5 = ag_5.sort_values(by='Promedio_Mensual', ascending=False)
        ag_5['%_Acumulado'] = (ag_5['Promedio_Mensual'].cumsum() / ag_5['Promedio_Mensual'].sum()) * 100

        # Gráfico Pareto
        fig5, ax5_left = plt.subplots(figsize=(14, 10))
        ax5_right = ax5_left.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_x_5 = np.arange(len(ag_5))
        
        # Barras de causas
        bar_p_5 = ax5_left.bar(pos_x_5, ag_5['Promedio_Mensual'], color='maroon', edgecolor='white', zorder=2)
        set_escala_superior(ax5_left, ag_5['Promedio_Mensual'].max(), 2.8)
        ax5_left.bar_label(bar_p_5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Línea Pareto ( Lorenz )
        ax5_right.plot(pos_x_5, ag_5['%_Acumulado'], color='red', marker='D', markersize=10, linewidth=4, path_effects=c_blanco, zorder=5)
        ax5_right.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax5_right.set_ylim(0, 200)
        ax5_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Etiquetas con wrap
        lbl_5 = [textwrap.fill(str(t), 12) for t in ag_5['TIPO_PARADA']]
        ax5_left.set_xticks(pos_x_5)
        ax5_left.set_xticklabels(lbl_5, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val_p in enumerate(ag_5['%_Acumulado']):
            ax5_right.annotate(f"{val_p:.1f}%", (pos_x_5[i], val_p + 4), color='white', bbox=box_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        sum_v_mensual = ag_5['Promedio_Mensual'].sum()
        ax5_left.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{sum_v_mensual:.1f} HH", transform=ax5_left.transAxes, bbox=box_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo")
        df_tbl_5 = ag_5.copy()
        sum_hs_total = df_tbl_5['HH_IMPRODUCTIVAS'].sum()
        df_tbl_5['% sobre Selección'] = (df_tbl_5['HH_IMPRODUCTIVAS'] / sum_hs_total) * 100
        
        # FILA TOTAL MAESTRA
        f_total_5 = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [sum_hs_total], 
            'Promedio_Mensual': [df_tbl_5['Promedio_Mensual'].sum()],
            '%_Acumulado': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tbl_5 = pd.concat([df_tbl_5, f_total_5], ignore_index=True)
        
        st.dataframe(
            df_tbl_5.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Causa'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Exportación CSV
        csv_p = df_tbl_5.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=csv_p, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ Sin horas improductivas en la selección actual.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de Horas Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        # Cruce temporal
        df_ef_f['K_Cruce'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp_6 = df_ef_f.groupby('K_Cruce', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['K_Cruce'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            pivote_6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K_Cruce', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_6_final = pd.merge(ag_disp_6, pivote_6, on='K_Cruce', how='left').fillna(0)
            list_cols_6 = [c for c in df_6_final.columns if c not in ['HH_Disponibles', 'K_Cruce']]
        else:
            df_6_final = ag_disp_6.copy(); list_cols_6 = []
            
        # Indicadores finales de incidencia
        df_6_final['Suma_I_6'] = df_6_final[list_cols_6].sum(axis=1) if list_cols_6 else 0
        df_6_final['Inc_Pct_6'] = (df_6_final['Suma_I_6'] / df_6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden cronológico
        df_6_final['F_Date_6'] = pd.to_datetime(df_6_final['K_Cruce'] + '-01')
        df_6_final = df_6_final.sort_values(by='F_Date_6')

        # Gráfico Incidencia
        fig6, ax6_base = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(txt_h_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_x_6 = np.arange(len(df_6_final))
        
        # Barras apiladas
        if list_cols_6:
            stack_base = np.zeros(len(df_6_final))
            paleta_inc = plt.cm.tab20.colors
            for idx_c, col_name in enumerate(list_cols_6):
                vals_c = df_6_final[col_name].values
                ax6_base.bar(pos_x_6, vals_c, bottom=stack_base, label=col_name, color=paleta_inc[idx_c % 20], edgecolor='white', zorder=2)
                stack_base += vals_c
        else:
            ax6_base.bar(pos_x_6, np.zeros(len(df_6_final)), color='white')

        set_escala_superior(ax6_base, df_6_final['Suma_I_6'].max(), 2.2)
        
        # Etiquetas volumen
        for i in range(len(pos_x_6)):
            val_i = df_6_final['Suma_I_6'].iloc[i]
            val_d = df_6_final['HH_Disponibles'].iloc[i]
            if val_i > 0: 
                ax6_base.annotate(f"Imp: {int(val_i)}\nDisp: {int(val_d)}", (i, val_i + (ax6_base.get_ylim()[1]*0.05)), ha='center', bbox=box_oro, fontweight='bold', zorder=10)

        # Línea de Porcentaje
        ax6_twin = ax6_base.twinx()
        ax6_twin.plot(pos_x_6, df_6_final['Inc_Pct_6'], color='red', marker='o', markersize=12, linewidth=6, path_effects=c_blanco, label='% Incidencia', zorder=5)
        
        # Meta 15%
        ax6_twin.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_twin.text(pos_x_6[0], 16, 'META = 15%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val_6 in enumerate(df_6_final['Inc_Pct_6']):
            ax6_twin.annotate(f"{val_6:.1f}%", (pos_x_6[i], val_6 + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=c_blanco, zorder=10)

        ax6_base.set_xticks(pos_x_6)
        ax6_base.set_xticklabels(df_6_final['K_Cruce'], fontsize=14, fontweight='bold')
        ax6_twin.set_ylim(0, max(30, df_6_final['Inc_Pct_6'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia para este sector.")
