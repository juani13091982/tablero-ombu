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
# 1. CONFIGURACIÓN DE LA PÁGINA (VISIÓN GERENCIAL)
# =========================================================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================================
# 2. SISTEMA DE SEGURIDAD Y ACCESO RESTRINGIDO
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    """
    Dibuja la pantalla de bienvenida institucional.
    Logo pequeño + Nombre oficial solicitado.
    """
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_vacia_1, col_login, col_vacia_2 = st.columns([1, 1.8, 1])
    
    with col_login:
        # Franja azul corporativa
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGO PEQUEÑO CENTRADO (160px)
        inner_l, inner_c, inner_r = st.columns([1, 1, 1])
        with inner_c:
            try:
                st.image("LOGO OMBÚ.jpg", width=160)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        # TEXTO INSTITUCIONAL EXPLICITO
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Formulario de acceso
        with st.form("formulario_seguridad"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            user_input = st.text_input("Usuario Corporativo")
            pass_input = st.text_input("Contraseña", type="password")
            
            boton_ingreso = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if boton_ingreso:
                if user_input in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[user_input] == pass_input:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos e intente de nuevo.")

# Ejecución del bloqueo
if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES AVANZADOS (CSS Y FUENTES)
# =========================================================================
css_estilos_interfaz = """
<style>
    /* Ocultar elementos nativos de Streamlit */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* FIJACIÓN DEL PANEL DE FILTROS EN LA PARTE SUPERIOR (STICKY) */
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
"""
st.markdown(css_estilos_interfaz, unsafe_allow_html=True)

# Configuración Maestra de Gráficos (Matplotlib)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Efectos de contorno de texto para máxima legibilidad
efecto_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
efecto_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Diseños de cajas de texto (Bboxes)
box_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
box_amarillo = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
box_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
box_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. FUNCIONES DE LÓGICA MATEMÁTICA Y CRUCE (MOTOR FUZZY)
# =========================================================================
def set_escala_superior(eje_plot, val_max, factor_margen=2.6):
    """Deja espacio arriba para que las etiquetas no se choquen."""
    if val_max > 0: 
        eje_plot.set_ylim(0, val_max * factor_margen)
    else: 
        eje_plot.set_ylim(0, 100)

def dibujar_lineas_verticales_mes(eje_plot, n_fechas):
    """Separa visualmente los meses en el eje X."""
    for i in range(n_fechas):
        eje_plot.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_labels_filtros(lista_seleccionada, texto_default):
    """Escribe los filtros activos en el título del gráfico."""
    if not lista_seleccionada: 
        return texto_default
    if len(lista_seleccionada) > 2: 
        return f"Varios ({len(lista_seleccionada)})"
    return " + ".join(lista_seleccionada)

def motor_fuzzy_de_cruce(nombre_seleccionado, nombre_excel_impro):
    """
    MOTOR INTELIGENTE DE CRUCE:
    Permite que '475-CARROZADO' coincida con 'SECTOR CARROZADO'.
    Indispensable para no perder las HH de improductivas.
    """
    if pd.isna(nombre_excel_impro) or pd.isna(nombre_seleccionado): 
        return False
    
    # 1. Normalización total
    t1 = str(nombre_seleccionado).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    t2 = str(nombre_excel_impro).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # 2. Limpieza de símbolos
    limpio1 = re.sub(r'[^A-Z0-9]', '', t1)
    limpio2 = re.sub(r'[^A-Z0-9]', '', t2)
    
    if not limpio1 or not limpio2: 
        return False
        
    # 3. Coincidencia por texto contenido
    if limpio1 in limpio2 or limpio2 in limpio1: 
        return True
    
    # 4. Coincidencia por código de estación (3+ dígitos)
    cod_sel = set(re.findall(r'\d{3,}', t1))
    cod_exc = set(re.findall(r'\d{3,}', t2))
    if cod_sel and cod_exc and cod_sel.intersection(cod_exc): 
        return True
        
    # 5. Coincidencia por palabras raíz (ej. 'CARROZADO')
    p_sel = set(re.findall(r'[A-Z]{4,}', t1))
    p_exc = set(re.findall(r'[A-Z]{4,}', t2))
    
    ignorar = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1 = p_sel - ignorar
    v2 = p_exc - ignorar
    
    for p in v1:
        if any(p in x for x in v2): 
            return True
                
    return False

# =========================================================================
# 5. CABECERA PRINCIPAL Y LOGOUT
# =========================================================================
col_l, col_t, col_s = st.columns([1, 3, 1])

with col_l:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with col_t:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Control de Gestión Productiva</p>", unsafe_allow_html=True)

with col_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA DE DATOS (EXCEL)
# =========================================================================
try:
    # Carga de archivos maestros
    df_ef_base = pd.read_excel("eficiencias.xlsx")
    df_im_base = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de encabezados
    df_ef_base.columns = df_ef_base.columns.str.strip()
    df_im_base.columns = [str(c).strip().upper() for c in df_im_base.columns]
    
    # Auto-corrector de columnas en improductivas (Blindaje contra cambios de nombre)
    if 'TIPO_PARADA' not in df_im_base.columns:
        c_alt_motivo = next((c for c in df_im_base.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if c_alt_motivo: df_im_base.rename(columns={c_alt_motivo: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im_base.columns:
        c_alt_hs = next((c for c in df_im_base.columns if 'HH' in c and 'IMP' in c), None)
        if c_alt_hs: df_im_base.rename(columns={c_alt_hs: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im_base.columns:
        c_alt_fec = next((c for c in df_im_base.columns if 'FECHA' in c), None)
        if c_alt_fec: df_im_base.rename(columns={c_alt_fec: 'FECHA'}, inplace=True)
    
    # Estandarización de Fechas
    df_ef_base['Fecha'] = pd.to_datetime(df_ef_base['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im_base['FECHA'] = pd.to_datetime(df_im_base['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación de estaciones finales (Es_Ultimo_Puesto)
    df_ef_base['Es_Ultimo_Puesto'] = df_ef_base['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Etiquetas de filtro temporal
    df_ef_base['Filtro_Mes'] = df_ef_base['Fecha'].dt.strftime('%b-%Y')
    df_im_base['Filtro_Mes'] = df_im_base['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e_carga:
    st.error(f"Error cargando los archivos Excel: {e_carga}")
    st.stop()

# =========================================================================
# 7. INTERFAZ DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    fl1, fl2, fl3, fl4 = st.columns(4)
    
    with fl1: 
        p_list = list(df_ef_base['Planta'].dropna().unique())
        s_planta = st.multiselect("🏭 Planta", p_list, placeholder="Todas")
        
    df_temp_l = df_ef_base[df_ef_base['Planta'].isin(s_planta)] if s_planta else df_ef_base
    
    with fl2: 
        l_list = list(df_temp_l['Linea'].dropna().unique())
        s_linea = st.multiselect("⚙️ Línea", l_list, placeholder="Todas")
        
    df_temp_p = df_temp_l[df_temp_l['Linea'].isin(s_linea)] if s_linea else df_temp_l
    
    with fl3: 
        ps_list = list(df_temp_p['Puesto_Trabajo'].dropna().unique())
        s_puesto = st.multiselect("🛠️ Puesto de Trabajo", ps_list, placeholder="Todos")
        
    with fl4: 
        m_list = ["🎯 Acumulado YTD"] + list(df_ef_base['Filtro_Mes'].unique())
        s_mes = st.multiselect("📅 Mes", m_list, placeholder="Todos")

# =========================================================================
# 8. PROCESAMIENTO DE DATOS FILTRADOS (CONSTRUCCIÓN DE VARIABLES)
# =========================================================================
# Copias locales para trabajar el filtrado
df_ef_f = df_ef_base.copy()
df_im_f = df_im_base.copy()

# APLICACIÓN FILTROS EFICIENCIA
if s_planta: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].isin(s_planta)]
if s_linea: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].isin(s_linea)]
if s_puesto: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(s_puesto)]
if s_mes and "🎯 Acumulado YTD" not in s_mes:
    df_ef_f = df_ef_f[df_ef_f['Filtro_Mes'].isin(s_mes)]

# APLICACIÓN FILTROS IMPRODUCTIVAS (MOTOR INTELIGENTE)
if s_planta:
    m_pl = df_im_f.iloc[:,0].apply(lambda x: any(motor_fuzzy_cruce(p, x) for p in s_planta))
    df_im_f = df_im_f[m_pl]

if s_linea:
    c_l_idx = next((c for c in df_im_f.columns if 'LINEA' in c), df_im_f.columns[1])
    m_li = df_im_f[c_l_idx].apply(lambda x: any(motor_fuzzy_cruce(l, x) for l in s_linea))
    df_im_f = df_im_f[m_li]

if s_puesto:
    c_p_idx = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2])
    m_ps = df_im_f[c_p_idx].apply(lambda x: any(motor_fuzzy_cruce(ps, x) for ps in s_puesto))
    df_im_f = df_im_f[m_ps]

if s_mes and "🎯 Acumulado YTD" not in s_mes:
    df_im_f = df_im_f[df_im_f['Filtro_Mes'].isin(s_mes)]

# Texto para el encabezado de los gráficos
txt_header_info = f"Planta: {formatear_labels_filtros(s_planta, 'Todas')} | Línea: {formatear_labels_filtros(s_linea, 'Todas')} | Puesto: {formatear_labels_filtros(s_puesto, 'Todos')}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (EFICIENCIAS)
# =========================================================================
col_izq_1, col_der_1 = st.columns(2)

with col_izq_1:
    st.header("1. EFICIENCIA REAL")
    # Lógica de salida: Si no hay puesto, mostramos solo 'SI' en salida.
    df_m1_plot = df_ef_f.copy() if s_puesto else df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_plot.empty:
        # Agrupación por mes
        agrup_1 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo de Eficiencia
        agrup_1['Efic_Calculada'] = (agrup_1['HH_STD_TOTAL'] / agrup_1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Inicio Gráfico 1
        fig1, ax1_bars = plt.subplots(figsize=(14, 10))
        ax1_line = ax1_bars.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_vals_1 = np.arange(len(agrup_1))
        
        # Dibujo de Barras
        bar_s_1 = ax1_bars.bar(x_vals_1 - 0.17, agrup_1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_d_1 = ax1_bars.bar(x_vals_1 + 0.17, agrup_1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_escala_superior(ax1_bars, agrup_1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas de volumen
        ax1_bars.bar_label(bar_s_1, padding=4, color='black', fontweight='bold', path_effects=efecto_blanco, fmt='%.0f', zorder=3)
        ax1_bars.bar_label(bar_d_1, padding=4, color='black', fontweight='bold', path_effects=efecto_blanco, fmt='%.0f', zorder=3)
        
        dibujar_lineas_verticales_mes(ax1_bars, len(x_vals_1))

        # Texto de Unidades Producidas (Vertical)
        for i, bar in enumerate(bar_s_1):
            val_u = int(agrup_1['Cant._Prod._A1'].iloc[i])
            if val_u > 0: 
                ax1_bars.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{val_u} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_negro, zorder=4)

        # Línea de Eficiencia Real (%)
        ax1_line.plot(x_vals_1, agrup_1['Efic_Calculada'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_blanco, label='% Eficiencia Real', zorder=5)
        
        # Línea de Meta Corporativa (85%)
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_vals_1[0], 86, 'META = 85%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_line.set_ylim(0, max(120, agrup_1['Efic_Calculada'].max()*1.8))
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación de Porcentaje en línea
        for i, val in enumerate(agrup_1['Efic_Calculada']):
            ax1_line.annotate(f"{val:.1f}%", (x_vals_1[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        # Configuración final ejes X
        ax1_bars.set_xticks(x_vals_1)
        ax1_bars.set_xticklabels(agrup_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        
        ax1_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin registros para graficar Eficiencia Real.")

with col_der_1:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_plot.empty:
        # Agrupación temporal
        agrup_2 = df_m1_plot.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        agrup_2['Efic_P'] = (agrup_2['HH_STD_TOTAL'] / agrup_2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Inicio Gráfico 2
        fig2, ax2_bars = plt.subplots(figsize=(14, 10))
        ax2_line = ax2_bars.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_vals_2 = np.arange(len(agrup_2))
        
        # Barras de Horas Comparativas
        bar_s_2 = ax2_bars.bar(x_vals_2 - 0.17, agrup_2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_p_2 = ax2_bars.bar(x_vals_2 + 0.17, agrup_2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_superior(ax2_bars, max(agrup_2['HH_STD_TOTAL'].max(), agrup_2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax2_bars.bar_label(bar_s_2, padding=4, color='black', fontweight='bold', path_effects=efecto_blanco, fmt='%.0f', zorder=3)
        ax2_bars.bar_label(bar_p_2, padding=4, color='black', fontweight='bold', path_effects=efecto_blanco, fmt='%.0f', zorder=3)
        
        dibujar_lineas_verticales_mes(ax2_bars, len(x_vals_2))

        # Línea de Eficiencia Productiva
        ax2_line.plot(x_vals_2, agrup_2['Efic_P'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta Corporativa (100%)
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_vals_2[0], 101, 'META = 100%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_line.set_ylim(0, max(150, agrup_2['Efic_P'].max()*1.8))
        ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación en línea
        for i, val in enumerate(agrup_2['Efic_P']):
            ax2_line.annotate(f"{val:.1f}%", (x_vals_2[i], val + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax2_bars.set_xticks(x_vals_2)
        ax2_bars.set_xticklabels(agrup_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_bars.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ Datos insuficientes para Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (ANÁLISIS DE GAP Y COSTOS)
# =========================================================================
col_izq_2, col_der_2 = st.columns(2)

with col_izq_2:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        # CORRECCIÓN DE SINTAXIS: Variable unificada sin espacios
        c_p_pura = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        
        ag_3 = df_ef_f.groupby('Fecha').agg({
            c_p_pura: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        ag_3['Suma_Declarada'] = ag_3[c_p_pura] + ag_3['HH_Improductivas']
        
        # Inicio Gráfico 3
        fig3, ax3_base = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_vals_3 = np.arange(len(ag_3))
        
        # Barras Apiladas (Prod + Impro)
        ax3_base.bar(x_vals_3, ag_3[c_p_pura], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3_base.bar(x_vals_3, ag_3['HH_Improductivas'], bottom=ag_3[c_p_pura], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad
        ax3_base.plot(x_vals_3, ag_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_blanco, label='HH DISPONIBLES', zorder=5)
        
        set_escala_superior(ax3_base, ag_3['HH_Disponibles'].max(), 2.6)
        dibujar_lineas_verticales_mes(ax3_base, len(x_vals_3))

        # Visualización técnica del GAP
        for i in range(len(x_vals_3)):
            h_dis = ag_3['HH_Disponibles'].iloc[i]
            h_dec = ag_3['Suma_Declarada'].iloc[i]
            val_gap = h_dis - h_dec
            
            # Línea vertical del GAP
            ax3_base.plot([i, i], [h_dec, h_dis], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Etiqueta GAP
            ax3_base.annotate(f"GAP:\n{int(val_gap)}", (i, h_dec + 5), color='firebrick', bbox=box_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Valor Disponibilidad Arriba
            ax3_base.annotate(f"{int(h_dis)}", (i, h_dis + (ax3_base.get_ylim()[1]*0.08)), color='black', bbox=box_blanco, ha='center', fontweight='bold', zorder=10)

        ax3_base.set_xticks(x_vals_3)
        ax3_base.set_xticklabels(ag_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3_base.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ No hay datos para análisis de GAP.")

with col_der_2:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Impacto económico de las horas perdidas (Improductivas)</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag_4 = df_ef_f.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        # Inicio Gráfico 4
        fig4, ax4_left = plt.subplots(figsize=(14, 10))
        ax4_right = ax4_left.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_vals_4 = np.arange(len(ag_4))
        
        # Barras de horas perdidas
        bar_lost_4 = ax4_left.bar(x_vals_4, ag_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_left.bar_label(bar_lost_4, padding=4, color='black', fontweight='bold', path_effects=efecto_blanco, zorder=4)
        
        set_escala_superior(ax4_left, ag_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Costo ($)
        ax4_right.plot(x_vals_4, ag_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_blanco, label='COSTO ARS', zorder=5)
        
        ax4_right.set_ylim(0, max(1000, ag_4['Costo_Improd._$'].max() * 1.8))
        ax4_right.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_right.get_yticks()], fontweight='bold')

        # Cartelera de Resumen de Costo
        val_t_pesos = ag_4['Costo_Improd._$'].sum()
        val_t_hs = ag_4['HH_Improductivas'].sum()
        ax4_left.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${val_t_pesos:,.0f}", transform=ax4_left.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=box_amarillo, weight='bold', zorder=10)

        for i, v in enumerate(ag_4['Costo_Improd._$']):
            ax4_right.annotate(f"${v:,.0f}", (x_vals_4[i], v + 5), color='white', bbox=box_gris, ha='center', fontweight='bold', zorder=10)

        ax4_left.set_xticks(x_vals_4)
        ax4_left.set_xticklabels(ag_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ Sin datos de costos económicos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (ANÁLISIS DE CAUSA RAÍZ)
# =========================================================================
col_izq_3, col_der_3 = st.columns(2)

with col_izq_3:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        # Agrupación por causas
        df_p_logic = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de meses
        meses_u = df_im_f['FECHA'].nunique()
        div_m_u = meses_u if meses_u > 0 else 1
        
        df_p_logic['Promedio_Mensual'] = df_p_logic['HH_IMPRODUCTIVAS'] / div_m_u
        df_p_logic = df_p_logic.sort_values(by='Promedio_Mensual', ascending=False)
        df_p_logic['%_Acumulado'] = (df_p_logic['Promedio_Mensual'].cumsum() / df_p_logic['Promedio_Mensual'].sum()) * 100

        # Inicio Gráfico 5
        fig5, ax5_left = plt.subplots(figsize=(14, 10))
        ax5_right = ax5_left.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_x_5 = np.arange(len(df_p_logic))
        
        # Barras Pareto
        b_p_m5 = ax5_left.bar(pos_x_5, df_p_logic['Promedio_Mensual'], color='maroon', edgecolor='white', zorder=2)
        set_escala_superior(ax5_left, df_p_logic['Promedio_Mensual'].max(), 2.8)
        ax5_left.bar_label(b_p_m5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva Lorenz
        ax5_right.plot(pos_x_5, df_p_logic['%_Acumulado'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_blanco, zorder=5)
        ax5_right.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax5_right.set_ylim(0, 200)
        ax5_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Nombres de causas envueltos
        lbls_5 = [textwrap.fill(str(txt), 12) for txt in df_p_logic['TIPO_PARADA']]
        ax5_left.set_xticks(pos_x_5)
        ax5_left.set_xticklabels(lbls_5, rotation=90, fontsize=12, fontweight='bold')
        
        for i, v in enumerate(df_p_logic['%_Acumulado']):
            ax5_right.annotate(f"{v:.1f}%", (pos_x_5[i], v + 4), color='white', bbox=box_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        v_mensual_sum = df_p_logic['Promedio_Mensual'].sum()
        ax5_left.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{v_mensual_sum:.1f} HH", transform=ax5_left.transAxes, bbox=box_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl_final = df_p_logic.copy()
        total_hs_final = df_tbl_final['HH_IMPRODUCTIVAS'].sum()
        df_tbl_final['% sobre Selección'] = (df_tbl_final['HH_IMPRODUCTIVAS'] / total_hs_final) * 100
        
        # INYECCIÓN FILA DE TOTAL
        f_total_row = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [total_hs_final], 
            'Promedio_Mensual': [df_tbl_final['Promedio_Mensual'].sum()],
            '%_Acumulado': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tbl_final = pd.concat([df_tbl_final, f_total_row], ignore_index=True)
        
        st.dataframe(
            df_tbl_final.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Botón de Descarga
        out_csv = df_tbl_final.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=out_csv, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Felicitaciones! Sin horas improductivas en la selección actual.")

with col_der_3:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de Horas Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        # Cruce por mes para el gráfico stacked
        df_ef_f['C_Key'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp_m6 = df_ef_f.groupby('C_Key', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['C_Key'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv_m6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='C_Key', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            m6_final = pd.merge(ag_disp_m6, piv_m6, on='C_Key', how='left').fillna(0)
            list_m6 = [c for c in m6_final.columns if c not in ['HH_Disponibles', 'C_Key']]
        else:
            m6_final = ag_disp_m6.copy()
            list_m6 = []
            
        # Indicadores Incidencia
        m6_final['Suma_I'] = m6_final[list_m6].sum(axis=1) if list_m6 else 0
        m6_final['Incid_%'] = (m6_final['Suma_I'] / m6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden de Fechas
        m6_final['F_Sort'] = pd.to_datetime(m6_final['C_Key'] + '-01')
        m6_final = m6_final.sort_values(by='F_Sort')

        # Inicio Gráfico 6
        fig6, ax6_base = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(txt_header_info, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_6 = np.arange(len(m6_final))
        
        # Barras Apiladas por Tipo de Parada
        if list_m6:
            piso_6 = np.zeros(len(m6_final))
            paleta_6 = plt.cm.tab20.colors
            for i, c_nom in enumerate(list_m6):
                val_c = m6_final[c_nom].values
                ax6_base.bar(pos_6, val_c, bottom=piso_6, label=c_nom, color=paleta_6[i % 20], edgecolor='white', zorder=2)
                piso_6 += val_c
        else:
            ax6_base.bar(pos_6, np.zeros(len(m6_final)), color='white')

        set_escala_superior(ax6_base, m6_final['Suma_I'].max(), 2.2)
        
        # Anotación volumen vs incidencia
        for i in range(len(pos_6)):
            iv = m6_final['Suma_I'].iloc[i]
            dv = m6_final['HH_Disponibles'].iloc[i]
            if iv > 0: 
                ax6_base.annotate(f"Imp: {int(iv)}\nDisp: {int(dv)}", (i, iv + (ax6_base.get_ylim()[1]*0.05)), ha='center', bbox=box_amarillo, fontweight='bold', zorder=10)

        # Línea Roja de Incidencia %
        ax6_twin = ax6_base.twinx()
        ax6_twin.plot(pos_6, m6_final['Incid_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_blanco, label='% Incidencia', zorder=5)
        
        # Meta 15%
        ax6_twin.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_twin.text(pos_6[0], 16, 'META = 15%', color='white', bbox=box_verde, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(m6_final['Incid_%']):
            ax6_twin.annotate(f"{val:.1f}%", (pos_6[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_blanco, zorder=10)

        ax6_base.set_xticks(pos_6)
        ax6_base.set_xticklabels(m6_final['C_Key'], fontsize=14, fontweight='bold')
        ax6_twin.set_ylim(0, max(30, m6_final['Incid_%'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia para este sector.")
