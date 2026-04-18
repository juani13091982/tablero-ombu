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
# 1. CONFIGURACIÓN DE LA ESTRUCTURA DE LA PÁGINA
# =========================================================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================================
# 2. SISTEMA DE ACCESO RESTRINGIDO (SEGURIDAD CORPORATIVA)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    """
    Dibuja la pantalla de inicio de sesión con el diseño institucional.
    Logo pequeño + Nombre completo de la compañía.
    """
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Columnas para centrar el formulario de acceso
    col_vacia_izq, col_login, col_vacia_der = st.columns([1, 1.8, 1])
    
    with col_login:
        # Franja decorativa superior
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # LOGOTIPO PEQUEÑO Y CENTRADO
        sub_col1, sub_col2, sub_col3 = st.columns([1, 1, 1])
        with sub_col2:
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
        
        # Formulario de credenciales
        with st.form("form_seguridad_corporativa"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            user_input = st.text_input("Usuario")
            pass_input = st.text_input("Contraseña", type="password")
            
            boton_acceso = st.form_submit_button("Ingresar al Tablero", use_container_width=True)

            if boton_acceso:
                if user_input in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[user_input] == pass_input:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos e intente nuevamente.")

# Ejecutar el bloqueo si no está logueado
if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES AVANZADOS Y CSS PERSONALIZADO
# =========================================================================
estilos_css = """
<style>
    /* Limpieza de la interfaz estándar de Streamlit */
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
st.markdown(estilos_css, unsafe_allow_html=True)

# Configuración Maestra de Matplotlib (Fuentes en Negrita y Grandes)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Efectos de contorno para que las etiquetas se vean claras sobre las barras
contorno_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
contorno_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cuadros de texto (Bboxes) para anotaciones
bbox_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_amarillo = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
bbox_blanco = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. FUNCIONES AUXILIARES DE LÓGICA Y CRUCE (MOTOR FUZZY)
# =========================================================================
def set_margen_superior(eje, max_val, multiplicador=2.6):
    """Deja espacio en el gráfico arriba para que las etiquetas no se corten."""
    if max_val > 0: 
        eje.set_ylim(0, max_val * multiplicador)
    else: 
        eje.set_ylim(0, 100)

def dibujar_separadores_mes(eje, num_periodos):
    """Dibuja las líneas punteadas para separar los meses en el eje X."""
    for i in range(num_periodos):
        eje.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_titulo_filtros(lista, predeterminado):
    """Prepara el texto que indica qué filtros están activos."""
    if not lista: 
        return predeterminado
    if len(lista) > 2: 
        return f"Varios ({len(lista)})"
    return " + ".join(lista)

def motor_cruce_robusto(seleccion_usuario, valor_excel):
    """
    CRUCE INTELIGENTE DE TEXTO:
    Resuelve el problema de que en el Excel de improductivas los nombres 
    de los puestos no siempre son idénticos a la base de eficiencias.
    """
    if pd.isna(valor_excel) or pd.isna(seleccion_usuario): 
        return False
    
    # Limpieza de acentos, mayúsculas y espacios
    s1 = str(seleccion_usuario).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(valor_excel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # Comparación solo con caracteres alfanuméricos
    limpio1 = re.sub(r'[^A-Z0-9]', '', s1)
    limpio2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not limpio1 or not limpio2: 
        return False
        
    # Prueba 1: Coincidencia directa
    if limpio1 in limpio2 or limpio2 in limpio1: 
        return True
    
    # Prueba 2: Coincidencia por código numérico (ej. '475')
    num_sel = set(re.findall(r'\d{3,}', s1))
    num_exc = set(re.findall(r'\d{3,}', s2))
    if num_sel and num_exc and num_sel.intersection(num_exc): 
        return True
        
    # Prueba 3: Coincidencia por palabras raíz (ej. 'CARROZADO')
    pal_sel = set(re.findall(r'[A-Z]{4,}', s1))
    pal_exc = set(re.findall(r'[A-Z]{4,}', s2))
    
    exclusiones = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    v1 = pal_sel - exclusiones
    v2 = pal_exc - exclusiones
    
    for p in v1:
        if any(p in x for x in v2): 
            return True
                
    return False

# =========================================================================
# 5. CABECERA PRINCIPAL DEL REPORTE
# =========================================================================
col_logo_main, col_tit_main, col_out_main = st.columns([1, 3, 1])

with col_logo_main:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with col_tit_main:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.subheader("Control de Gestión Productiva")

with col_out_main:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA Y PREPARACIÓN DE DATOS (AUTO-CORRECTOR)
# =========================================================================
try:
    # Lectura de archivos
    df_efi_base = pd.read_excel("eficiencias.xlsx")
    df_imp_base = pd.read_excel("improductivas.xlsx")
    
    # Limpieza de columnas
    df_efi_base.columns = df_efi_base.columns.str.strip()
    df_imp_base.columns = [str(c).strip().upper() for c in df_imp_base.columns]
    
    # Motor de auto-mapeo para la base de improductivas
    if 'TIPO_PARADA' not in df_imp_base.columns:
        c_mot = next((c for c in df_imp_base.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if c_mot: df_imp_base.rename(columns={c_mot: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_imp_base.columns:
        c_hhs = next((c for c in df_imp_base.columns if 'HH' in c and 'IMP' in c), None)
        if c_hhs: df_imp_base.rename(columns={c_hhs: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_imp_base.columns:
        c_fec = next((c for c in df_imp_base.columns if 'FECHA' in c), None)
        if c_fec: df_imp_base.rename(columns={c_fec: 'FECHA'}, inplace=True)
    
    # Conversión de Fechas a formato estándar mensual
    df_efi_base['Fecha'] = pd.to_datetime(df_efi_base['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_imp_base['FECHA'] = pd.to_datetime(df_imp_base['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación de puestos terminales
    df_efi_base['Es_Ultimo_Puesto'] = df_efi_base['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Etiquetas de filtro de meses
    df_efi_base['Etiqueta_Mes'] = df_efi_base['Fecha'].dt.strftime('%b-%Y')
    df_imp_base['Etiqueta_Mes'] = df_imp_base['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e_carga:
    st.error(f"Error fatal cargando bases de datos: {e_carga}")
    st.stop()

# =========================================================================
# 7. INTERFAZ DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    fl1, fl2, fl3, fl4 = st.columns(4)
    
    with fl1: 
        plantas_disponibles = list(df_efi_base['Planta'].dropna().unique())
        filtro_planta = st.multiselect("🏭 Planta", plantas_disponibles, placeholder="Todas")
        
    df_tmp_lineas = df_efi_base[df_efi_base['Planta'].isin(filtro_planta)] if filtro_planta else df_efi_base
    
    with fl2: 
        lineas_disponibles = list(df_tmp_lineas['Linea'].dropna().unique())
        filtro_linea = st.multiselect("⚙️ Línea", lineas_disponibles, placeholder="Todas")
        
    df_tmp_puestos = df_tmp_lineas[df_tmp_lineas['Linea'].isin(filtro_linea)] if filtro_linea else df_tmp_lineas
    
    with fl3: 
        puestos_disponibles = list(df_tmp_puestos['Puesto_Trabajo'].dropna().unique())
        filtro_puesto = st.multiselect("🛠️ Puesto de Trabajo", puestos_disponibles, placeholder="Todos")
        
    with fl4: 
        meses_disponibles = ["🎯 Acumulado YTD"] + list(df_efi_base['Etiqueta_Mes'].unique())
        filtro_mes = st.multiselect("📅 Mes", meses_disponibles, placeholder="Todos")

# =========================================================================
# 8. LÓGICA DE FILTRADO FINAL DE DATAFRAMES
# =========================================================================
df_efi_final = df_efi_base.copy()
df_imp_final = df_imp_base.copy()

# Filtros para Eficiencias
if filtro_planta: 
    df_efi_final = df_efi_final[df_efi_final['Planta'].isin(filtro_planta)]
if filtro_linea: 
    df_efi_final = df_efi_final[df_efi_final['Linea'].isin(filtro_linea)]
if filtro_puesto: 
    df_efi_final = df_efi_final[df_efi_final['Puesto_Trabajo'].isin(filtro_puesto)]
if filtro_mes and "🎯 Acumulado YTD" not in filtro_mes:
    df_efi_final = df_efi_final[df_efi_final['Etiqueta_Mes'].isin(filtro_mes)]

# Filtros para Improductivas (Utilizando el Motor Robusto)
if filtro_planta:
    mask_pl = df_imp_final.iloc[:,0].apply(lambda x: any(motor_cruce_robusto(p, x) for p in filtro_planta))
    df_imp_final = df_imp_final[mask_pl]

if filtro_linea:
    c_lin_idx = next((c for c in df_imp_final.columns if 'LINEA' in c), df_imp_final.columns[1])
    mask_li = df_imp_final[c_lin_idx].apply(lambda x: any(motor_cruce_robusto(l, x) for l in filtro_linea))
    df_imp_final = df_imp_final[mask_li]

if filtro_puesto:
    c_pue_idx = next((c for c in df_imp_final.columns if 'PUESTO' in c), df_imp_final.columns[2])
    mask_ps = df_imp_final[c_pue_idx].apply(lambda x: any(motor_cruce_robusto(ps, x) for ps in filtro_puesto))
    df_imp_final = df_imp_final[mask_ps]

if filtro_mes and "🎯 Acumulado YTD" not in filtro_mes:
    df_imp_final = df_imp_final[df_imp_final['Etiqueta_Mes'].isin(filtro_mes)]

# Encabezado dinámico para títulos de gráficos
txt_info_escenario = f"FILTRO: {formatear_titulo_filtros(filtro_planta, 'Todas')} > {formatear_titulo_filtros(filtro_linea, 'Todas')} > {formatear_titulo_filtros(filtro_puesto, 'Todos')}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2 (PRODUCTIVIDAD)
# =========================================================================
col_m1, col_m2 = st.columns(2)

with col_m1:
    st.header("1. EFICIENCIA REAL")
    # Regla: Si no hay puesto específico, mostrar solo datos de salida ('SI')
    df_m1_grafico = df_efi_final.copy() if filtro_puesto else df_efi_final[df_efi_final['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_m1_grafico.empty:
        agrup_m1 = df_m1_grafico.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        agrup_m1['Ef_Real'] = (agrup_m1['HH_STD_TOTAL'] / agrup_m1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Construcción detallada del gráfico
        fig1, ax1_left = plt.subplots(figsize=(14, 10))
        ax1_right = ax1_left.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices = np.arange(len(agrup_m1))
        ancho_b = 0.35
        
        # Barras de Horas
        b_std = ax1_left.bar(x_indices - ancho_b/2, agrup_m1['HH_STD_TOTAL'], ancho_b, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_disp = ax1_left.bar(x_indices + ancho_b/2, agrup_m1['HH_Disponibles'], ancho_b, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_margen_superior(ax1_left, agrup_m1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas numéricas en barras
        ax1_left.bar_label(b_std, padding=4, color='black', fontweight='bold', path_effects=contorno_blanco, fmt='%.0f', zorder=3)
        ax1_left.bar_label(b_disp, padding=4, color='black', fontweight='bold', path_effects=contorno_blanco, fmt='%.0f', zorder=3)
        
        dibujar_separadores_mes(ax1_left, len(x_indices))

        # Texto vertical de Unidades Producidas
        for i, bar in enumerate(b_std):
            num_u = int(agrup_m1['Cant._Prod._A1'].iloc[i])
            if num_u > 0: 
                ax1_left.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{num_u} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=contorno_negro, zorder=4)

        # Línea de Eficiencia Real (%)
        ax1_right.plot(x_indices, agrup_m1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=contorno_blanco, label='% Efic. Real', zorder=5)
        
        # Meta 85%
        ax1_right.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_right.text(x_indices[0], 86, 'META = 85%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_right.set_ylim(0, max(120, agrup_m1['Ef_Real'].max()*1.8))
        ax1_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación de % en línea
        for i, val in enumerate(agrup_m1['Ef_Real']):
            ax1_right.annotate(f"{val:.1f}%", (x_indices[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax1_left.set_xticks(x_indices)
        ax1_left.set_xticklabels(agrup_m1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        
        ax1_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax1_right.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
    else: 
        st.warning("⚠️ Sin datos para graficar con el filtro actual.")

with col_m2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1_grafico.empty:
        agrup_m2 = df_m1_grafico.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        agrup_m2['Ef_Prod'] = (agrup_m2['HH_STD_TOTAL'] / agrup_m2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2_left = plt.subplots(figsize=(14, 10))
        ax2_right = ax2_left.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_indices2 = np.arange(len(agrup_m2))
        
        # Barras Comparativas
        b_std2 = ax2_left.bar(x_indices2 - ancho_b/2, agrup_m2['HH_STD_TOTAL'], ancho_b, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_prod2 = ax2_left.bar(x_indices2 + ancho_b/2, agrup_m2['HH_Productivas_C/GAP'], ancho_b, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_margen_superior(ax2_left, max(agrup_m2['HH_STD_TOTAL'].max(), agrup_m2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax2_left.bar_label(b_std2, padding=4, color='black', fontweight='bold', path_effects=contorno_blanco, fmt='%.0f', zorder=3)
        ax2_left.bar_label(b_prod2, padding=4, color='black', fontweight='bold', path_effects=contorno_blanco, fmt='%.0f', zorder=3)
        
        dibujar_separadores_mes(ax2_left, len(x_indices2))

        # Línea de Eficiencia Productiva
        ax2_right.plot(x_indices2, agrup_m2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=contorno_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta 100%
        ax2_right.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_right.text(x_indices2[0], 101, 'META = 100%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_right.set_ylim(0, max(150, agrup_m2['Ef_Prod'].max()*1.8))
        ax2_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación en línea
        for i, val in enumerate(agrup_m2['Ef_Prod']):
            ax2_right.annotate(f"{val:.1f}%", (x_indices2[i], val + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax2_left.set_xticks(x_indices2)
        ax2_left.set_xticklabels(agrup_m2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ No existen registros para la comparativa productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (GAP Y COSTOS)
# =========================================================================
col_m3, col_m4 = st.columns(2)

with col_m3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Horas Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_efi_final.empty:
        # CORRECCIÓN DE SINTAXIS DEFINITIVA: Variable sin espacios
        c_prod_maestra = 'HH_Productivas' if 'HH_Productivas' in df_efi_final.columns else 'HH Productivas'
        
        ag_m3 = df_efi_final.groupby('Fecha').agg({
            c_prod_maestra: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        ag_m3['Total_Declaradas'] = ag_m3[c_prod_maestra] + ag_m3['HH_Improductivas']
        
        fig3, ax3_gap = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_m3 = np.arange(len(ag_m3))
        
        # Dibujo de Barras Apiladas
        ax3_gap.bar(x_m3, ag_m3[c_prod_maestra], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        ax3_gap.bar(x_m3, ag_m3['HH_Improductivas'], bottom=ag_m3[c_prod_maestra], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad
        ax3_gap.plot(x_m3, ag_m3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=contorno_blanco, label='HH DISPONIBLES', zorder=5)
        
        set_margen_superior(ax3_gap, ag_m3['HH_Disponibles'].max(), 2.6)
        dibujar_separadores_mes(ax3_gap, len(x_m3))

        # Cálculo y dibujo visual del GAP
        for i in range(len(x_m3)):
            h_disp = ag_m3['HH_Disponibles'].iloc[i]
            h_decl = ag_m3['Total_Declaradas'].iloc[i]
            val_gap = h_disp - h_decl
            
            # Línea vertical del GAP
            ax3_gap.plot([i, i], [h_decl, h_disp], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Cartel del GAP
            ax3_gap.annotate(f"GAP:\n{int(val_gap)}", (i, h_decl + 5), color='firebrick', bbox=bbox_blanco, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Valor Disponibilidad arriba
            ax3_gap.annotate(f"{int(h_disp)}", (i, h_disp + (ax3_gap.get_ylim()[1]*0.08)), color='black', bbox=bbox_blanco, ha='center', fontweight='bold', zorder=10)

        ax3_gap.set_xticks(x_m3)
        ax3_gap.set_xticklabels(ag_m3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3_gap.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ Sin datos para análisis de GAP.")

with col_m4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización económica de las horas perdidas</i></div>", unsafe_allow_html=True)
    
    if not df_efi_final.empty:
        ag_m4 = df_efi_final.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        fig4, ax4_left = plt.subplots(figsize=(14, 10))
        ax4_right = ax4_left.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_m4 = np.arange(len(ag_m4))
        
        # Barras de horas perdidas
        b_perdida = ax4_left.bar(x_m4, ag_m4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4_left.bar_label(b_perdida, padding=4, color='black', fontweight='bold', path_effects=contorno_blanco, zorder=4)
        
        set_margen_superior(ax4_left, ag_m4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Pesos (Costo)
        ax4_right.plot(x_m4, ag_m4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=contorno_blanco, label='COSTO ARS', zorder=5)
        
        ax4_right.set_ylim(0, max(1000, ag_m4['Costo_Improd._$'].max() * 1.8))
        ax4_right.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_right.get_yticks()], fontweight='bold')

        # Cartel de Costo Total
        p_total = ag_m4['Costo_Improd._$'].sum()
        ax4_left.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${p_total:,.0f}", transform=ax4_left.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=bbox_amarillo, weight='bold', zorder=10)

        for i, v in enumerate(ag_m4['Costo_Improd._$']):
            ax4_right.annotate(f"${v:,.0f}", (x_m4[i], v + 5), color='white', bbox=bbox_gris, ha='center', fontweight='bold', zorder=10)

        ax4_left.set_xticks(x_m4)
        ax4_left.set_xticklabels(ag_m4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4_left.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ Sin datos de costos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (PARETO E INCIDENCIA)
# =========================================================================
col_m5, col_m6 = st.columns(2)

with col_m5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        # Procesamiento Pareto
        df_par = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de meses para promedios
        n_m = df_im_f['FECHA'].nunique()
        divisor_m = n_m if n_m > 0 else 1
        
        df_par['Promedio'] = df_par['HH_IMPRODUCTIVAS'] / divisor_m
        df_par = df_par.sort_values(by='Promedio', ascending=False)
        df_par['%_Acu'] = (df_par['Promedio'].cumsum() / df_par['Promedio'].sum()) * 100

        fig5, ax5_left = plt.subplots(figsize=(14, 10))
        ax5_right = ax5_left.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_x = np.arange(len(df_par))
        
        # Barras
        bar_p = ax5_left.bar(pos_x, df_par['Promedio'], color='maroon', edgecolor='white', zorder=2)
        set_margen_superior(ax5_left, df_par['Promedio'].max(), 2.8)
        ax5_left.bar_label(bar_p, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Línea Pareto
        ax5_right.plot(pos_x, df_par['%_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=contorno_blanco, zorder=5)
        ax5_right.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax5_right.set_ylim(0, 200)
        ax5_right.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Wrap nombres causas
        labels_c = [textwrap.fill(str(txt), 12) for txt in df_par['TIPO_PARADA']]
        ax5_left.set_xticks(pos_x)
        ax5_left.set_xticklabels(labels_c, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val in enumerate(df_par['%_Acu']):
            ax5_right.annotate(f"{val:.1f}%", (pos_x[i], val + 4), color='white', bbox=bbox_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        s_prom = df_par['Promedio'].sum()
        ax5_left.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{s_prom:.1f} HH", transform=ax5_left.transAxes, bbox=bbox_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO E IMPACTO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl_final = df_par.copy()
        total_hs_par = df_tbl_final['HH_IMPRODUCTIVAS'].sum()
        df_tbl_final['% sobre Selección'] = (df_tbl_final['HH_IMPRODUCTIVAS'] / total_hs_par) * 100
        
        # FILA TOTAL
        f_total = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [total_hs_par], 
            'Promedio': [df_tbl_final['Promedio'].sum()],
            '%_Acu': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tbl_final = pd.concat([df_tbl_final, f_total], ignore_index=True)
        
        st.dataframe(
            df_tbl_final.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Categoría'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Descarga Excel
        c_csv = df_tbl_final.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=c_csv, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Felicidades! 0 Horas improductivas en la selección.")

with col_m6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de Horas Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_efi_final.empty:
        # Cruce de datos por mes
        df_efi_final['Mes_Cruce'] = df_efi_final['Fecha'].dt.strftime('%Y-%m')
        ag_disp_m6 = df_efi_final.groupby('Mes_Cruce', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['Mes_Cruce'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv_m6 = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='Mes_Cruce', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_m6_final = pd.merge(ag_disp_m6, piv_m6, on='Mes_Cruce', how='left').fillna(0)
            lista_motivos = [c for c in df_m6_final.columns if c not in ['HH_Disponibles', 'Mes_Cruce']]
        else:
            df_m6_final = ag_disp_m6.copy()
            lista_motivos = []
            
        # Indicadores finales
        df_m6_final['Total_Hs_Imp'] = df_m6_final[lista_motivos].sum(axis=1) if lista_motivos else 0
        df_m6_final['Incidencia_%'] = (df_m6_final['Total_Hs_Imp'] / df_m6_final['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Orden de fechas
        df_m6_final['Orden'] = pd.to_datetime(df_m6_final['Mes_Cruce'] + '-01')
        df_m6_final = df_m6_final.sort_values(by='Orden')

        fig6, ax6_base = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(txt_info_escenario, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        pos_m6 = np.arange(len(df_m6_final))
        
        # Dibujo de barras apiladas
        if lista_motivos:
            piso = np.zeros(len(df_m6_final))
            paleta_c = plt.cm.tab20.colors
            for idx_c, nom_c in enumerate(lista_motivos):
                vals_c = df_m6_final[nom_c].values
                ax6_base.bar(pos_m6, vals_c, bottom=piso, label=nom_c, color=paleta_c[idx_c % 20], edgecolor='white', zorder=2)
                piso += vals_c
        else:
            ax6_base.bar(pos_m6, np.zeros(len(df_m6_final)), color='white')

        set_margen_superior(ax6_base, df_m6_final['Total_Hs_Imp'].max(), 2.2)
        
        # Anotación volumen
        for i in range(len(pos_m6)):
            iv = df_m6_final['Total_Hs_Imp'].iloc[i]
            dv = df_m6_final['HH_Disponibles'].iloc[i]
            if iv > 0: 
                ax6_base.annotate(f"Imp: {int(iv)}\nDisp: {int(dv)}", (i, iv + (ax6_base.get_ylim()[1]*0.05)), ha='center', bbox=bbox_amarillo, fontweight='bold', zorder=10)

        # Línea de Porcentaje
        ax6_inc = ax6_base.twinx()
        ax6_inc.plot(pos_m6, df_m6_final['Incidencia_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=contorno_blanco, label='% Incidencia', zorder=5)
        
        ax6_inc.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_inc.text(pos_m6[0], 16, 'META = 15%', color='white', bbox=bbox_verde, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(df_m6_final['Incidencia_%']):
            ax6_inc.annotate(f"{val:.1f}%", (pos_m6[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=contorno_blanco, zorder=10)

        ax6_base.set_xticks(pos_m6)
        ax6_base.set_xticklabels(df_m6_final['Mes_Cruce'], fontsize=14, fontweight='bold')
        ax6_inc.set_ylim(0, max(30, df_m6_final['Incidencia_%'].max() * 1.8))
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ Sin datos históricos de eficiencia.")
