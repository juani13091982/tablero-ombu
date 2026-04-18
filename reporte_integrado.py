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
# 2. SISTEMA DE SEGURIDAD CORPORATIVA (LLAVE MAESTRA)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_pantalla_login():
    """
    Despliega la interfaz inicial de seguridad para proteger los datos de la compañía.
    """
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    columna_vacia_izq, columna_login, columna_vacia_der = st.columns([1, 1.8, 1])
    
    with columna_login:
        # Franja decorativa institucional
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:5px; border-radius:10px 10px 0px 0px; text-align:center;'>
            </div>
        """, unsafe_allow_html=True)
        
        # Logotipo centrado
        sub_col1, sub_col2, sub_col3 = st.columns([1, 1, 1])
        with sub_col2:
            try:
                st.image("LOGO OMBÚ.jpg", width=160)
            except Exception:
                st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        # Textos oficiales
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold; letter-spacing: 1px;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2>
                <p style='margin:0; color:#666; font-size:16px; font-weight:bold;'>Acceso Restringido al Tablero de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Formulario de credenciales
        with st.form("formulario_de_acceso"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            input_usuario = st.text_input("Usuario Corporativo")
            input_clave = st.text_input("Contraseña", type="password")
            
            boton_ingresar = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if boton_ingresar:
                if input_usuario in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[input_usuario] == input_clave:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Por favor, verifique los datos.")

# Validación de estado
if not st.session_state['autenticado']:
    mostrar_pantalla_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES (CSS) Y CONFIGURACIÓN DE GRÁFICOS
# =========================================================================
codigo_css = """
<style>
    /* Ocultar menú nativo de Streamlit */
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}

    /* FIJACIÓN DEL PANEL DE FILTROS EN LA PARTE SUPERIOR */
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
st.markdown(codigo_css, unsafe_allow_html=True)

# Parámetros estrictos de Matplotlib para legibilidad gerencial
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Contornos para evitar superposición visual de textos
efecto_contorno_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
efecto_contorno_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cajas de texto
caja_gris = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_amarilla = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
caja_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
caja_blanca = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. FUNCIONES AUXILIARES (MOTOR INTELIGENTE)
# =========================================================================
def aplicar_margen_superior_grafico(eje_matplotlib, valor_maximo, multiplicador=2.6):
    """Evita que las barras altas oculten el título o las etiquetas superiores."""
    if valor_maximo > 0: 
        eje_matplotlib.set_ylim(0, valor_maximo * multiplicador)
    else: 
        eje_matplotlib.set_ylim(0, 100)

def trazar_lineas_divisorias_meses(eje_matplotlib, cantidad_periodos):
    """Dibuja guías verticales para separar los meses en el eje X."""
    for indice in range(cantidad_periodos):
        eje_matplotlib.axvline(x=indice, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_texto_filtros(lista_seleccion, texto_por_defecto):
    """Resume los filtros para los títulos de los gráficos."""
    if not lista_seleccion: 
        return texto_por_defecto
    if len(lista_seleccion) > 2: 
        return f"Varios ({len(lista_seleccion)})"
    return " + ".join(lista_seleccion)

def motor_analisis_robusto(texto_seleccionado, texto_excel_improductivas):
    """
    Función crítica: Cruza estaciones de Eficiencia con Improductivas 
    buscando raíces de palabras y números, ignorando diferencias de tipeo.
    """
    if pd.isna(texto_excel_improductivas) or pd.isna(texto_seleccionado): 
        return False
    
    str_1 = str(texto_seleccionado).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    str_2 = str(texto_excel_improductivas).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    alfa_1 = re.sub(r'[^A-Z0-9]', '', str_1)
    alfa_2 = re.sub(r'[^A-Z0-9]', '', str_2)
    
    if not alfa_1 or not alfa_2: 
        return False
        
    if alfa_1 in alfa_2 or alfa_2 in alfa_1: 
        return True
    
    numeros_1 = set(re.findall(r'\d{3,}', str_1))
    numeros_2 = set(re.findall(r'\d{3,}', str_2))
    if numeros_1 and numeros_2 and numeros_1.intersection(numeros_2): 
        return True
        
    palabras_1 = set(re.findall(r'[A-Z]{4,}', str_1))
    palabras_2 = set(re.findall(r'[A-Z]{4,}', str_2))
    
    palabras_ignoradas = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    conjunto_1 = palabras_1 - palabras_ignoradas
    conjunto_2 = palabras_2 - palabras_ignoradas
    
    for palabra in conjunto_1:
        if any(palabra in x for x in conjunto_2): 
            return True
                
    return False

# =========================================================================
# 5. ENCABEZADO DEL TABLERO
# =========================================================================
columna_logo, columna_titulo, columna_salida = st.columns([1, 3, 1])

with columna_logo:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with columna_titulo:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Gerencia de Control de Gestión</p>", unsafe_allow_html=True)

with columna_salida:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA Y DEPURACIÓN DE DATOS (EXCEL)
# =========================================================================
try:
    df_eficiencias_raw = pd.read_excel("eficiencias.xlsx")
    df_improductivas_raw = pd.read_excel("improductivas.xlsx")
    
    df_eficiencias_raw.columns = df_eficiencias_raw.columns.str.strip()
    df_improductivas_raw.columns = [str(c).strip().upper() for c in df_improductivas_raw.columns]
    
    # Auto-identificador de columnas de Improductivas
    if 'TIPO_PARADA' not in df_improductivas_raw.columns:
        col_tipo = next((c for c in df_improductivas_raw.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if col_tipo: df_improductivas_raw.rename(columns={col_tipo: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_improductivas_raw.columns:
        col_hs_imp = next((c for c in df_improductivas_raw.columns if 'HH' in c and 'IMP' in c), None)
        if col_hs_imp: df_improductivas_raw.rename(columns={col_hs_imp: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_improductivas_raw.columns:
        col_fecha_imp = next((c for c in df_improductivas_raw.columns if 'FECHA' in c), None)
        if col_fecha_imp: df_improductivas_raw.rename(columns={col_fecha_imp: 'FECHA'}, inplace=True)
    
    # Procesamiento de fechas
    df_eficiencias_raw['Fecha'] = pd.to_datetime(df_eficiencias_raw['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_improductivas_raw['FECHA'] = pd.to_datetime(df_improductivas_raw['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    df_eficiencias_raw['Es_Ultimo_Puesto'] = df_eficiencias_raw['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_eficiencias_raw['Mes_String'] = df_eficiencias_raw['Fecha'].dt.strftime('%b-%Y')
    df_improductivas_raw['Mes_String'] = df_improductivas_raw['FECHA'].dt.strftime('%b-%Y')
    
except Exception as error_lectura:
    st.error(f"Error crítico al leer los archivos: {error_lectura}")
    st.stop()

# =========================================================================
# 7. BARRA DE FILTROS (CASCADA)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns(4)
    
    with filtro_col1: 
        lista_plantas = list(df_eficiencias_raw['Planta'].dropna().unique())
        seleccion_planta = st.multiselect("🏭 Planta", lista_plantas, placeholder="Todas")
        
    df_transitorio_lineas = df_eficiencias_raw[df_eficiencias_raw['Planta'].isin(seleccion_planta)] if seleccion_planta else df_eficiencias_raw
    
    with filtro_col2: 
        lista_lineas = list(df_transitorio_lineas['Linea'].dropna().unique())
        seleccion_linea = st.multiselect("⚙️ Línea", lista_lineas, placeholder="Todas")
        
    df_transitorio_puestos = df_transitorio_lineas[df_transitorio_lineas['Linea'].isin(seleccion_linea)] if seleccion_linea else df_transitorio_lineas
    
    with filtro_col3: 
        lista_puestos = list(df_transitorio_puestos['Puesto_Trabajo'].dropna().unique())
        seleccion_puesto = st.multiselect("🛠️ Puesto de Trabajo", lista_puestos, placeholder="Todos")
        
    with filtro_col4: 
        lista_meses = ["🎯 Acumulado YTD"] + list(df_eficiencias_raw['Mes_String'].unique())
        seleccion_mes = st.multiselect("📅 Mes", lista_meses, placeholder="Todos")

# =========================================================================
# 8. FILTRADO FINAL DE DATAFRAMES (BLINDADO CONTRA INDEX ERROR)
# =========================================================================
df_eficiencias_filtrado = df_eficiencias_raw.copy()
df_improductivas_filtrado = df_improductivas_raw.copy()

# Aplicar filtros a Eficiencias
if seleccion_planta: 
    df_eficiencias_filtrado = df_eficiencias_filtrado[df_eficiencias_filtrado['Planta'].isin(seleccion_planta)]
if seleccion_linea: 
    df_eficiencias_filtrado = df_eficiencias_filtrado[df_eficiencias_filtrado['Linea'].isin(seleccion_linea)]
if seleccion_puesto: 
    df_eficiencias_filtrado = df_eficiencias_filtrado[df_eficiencias_filtrado['Puesto_Trabajo'].isin(seleccion_puesto)]
if seleccion_mes and "🎯 Acumulado YTD" not in seleccion_mes:
    df_eficiencias_filtrado = df_eficiencias_filtrado[df_eficiencias_filtrado['Mes_String'].isin(seleccion_mes)]

# Aplicar filtros a Improductivas (Cruce Robusto y Seguro)
if seleccion_planta and not df_improductivas_filtrado.empty:
    col_planta = next((c for c in df_improductivas_filtrado.columns if 'PLANTA' in str(c).upper()), df_improductivas_filtrado.columns[0] if len(df_improductivas_filtrado.columns) > 0 else None)
    if col_planta:
        mascara_planta = df_improductivas_filtrado[col_planta].apply(lambda valor: any(motor_analisis_robusto(p, valor) for p in seleccion_planta))
        df_improductivas_filtrado = df_improductivas_filtrado[mascara_planta]

if seleccion_linea and not df_improductivas_filtrado.empty:
    col_linea = next((c for c in df_improductivas_filtrado.columns if 'LINEA' in str(c).upper()), df_improductivas_filtrado.columns[1] if len(df_improductivas_filtrado.columns) > 1 else None)
    if col_linea:
        mascara_linea = df_improductivas_filtrado[col_linea].apply(lambda valor: any(motor_analisis_robusto(l, valor) for l in seleccion_linea))
        df_improductivas_filtrado = df_improductivas_filtrado[mascara_linea]

if seleccion_puesto and not df_improductivas_filtrado.empty:
    col_puesto = next((c for c in df_improductivas_filtrado.columns if 'PUESTO' in str(c).upper()), df_improductivas_filtrado.columns[2] if len(df_improductivas_filtrado.columns) > 2 else None)
    if col_puesto:
        mascara_puesto = df_improductivas_filtrado[col_puesto].apply(lambda valor: any(motor_analisis_robusto(ps, valor) for ps in seleccion_puesto))
        df_improductivas_filtrado = df_improductivas_filtrado[mascara_puesto]

if seleccion_mes and "🎯 Acumulado YTD" not in seleccion_mes and not df_improductivas_filtrado.empty:
    df_improductivas_filtrado = df_improductivas_filtrado[df_improductivas_filtrado['Mes_String'].isin(seleccion_mes)]

# Texto dinámico para títulos
texto_parametros_activos = f"Filtros >> Planta: {formatear_texto_filtros(seleccion_planta, 'Todas')} | Línea: {formatear_texto_filtros(seleccion_linea, 'Todas')} | Puesto: {formatear_texto_filtros(seleccion_puesto, 'Todos')}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2
# =========================================================================
columna_metrica_1, columna_metrica_2 = st.columns(2)

with columna_metrica_1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    # Lógica de terminal: si no hay puesto filtrado, tomar solo 'SI'
    df_base_grafico_1 = df_eficiencias_filtrado.copy() if seleccion_puesto else df_eficiencias_filtrado[df_eficiencias_filtrado['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_base_grafico_1.empty:
        agrupacion_1 = df_base_grafico_1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo seguro de KPI
        calculo_kpi_1 = (agrupacion_1['HH_STD_TOTAL'] / agrupacion_1['HH_Disponibles'])
        agrupacion_1['Efic_Real_Porcentaje'] = calculo_kpi_1.replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        figura_1, eje_1_barras = plt.subplots(figsize=(14, 10))
        eje_1_linea = eje_1_barras.twinx()
        
        figura_1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_1.suptitle(texto_parametros_activos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        posiciones_x_1 = np.arange(len(agrupacion_1))
        ancho_barra = 0.35
        
        # Barras
        barra_std_1 = eje_1_barras.bar(posiciones_x_1 - ancho_barra/2, agrupacion_1['HH_STD_TOTAL'], ancho_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        barra_dis_1 = eje_1_barras.bar(posiciones_x_1 + ancho_barra/2, agrupacion_1['HH_Disponibles'], ancho_barra, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        aplicar_margen_superior_grafico(eje_1_barras, agrupacion_1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas
        eje_1_barras.bar_label(barra_std_1, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        eje_1_barras.bar_label(barra_dis_1, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        
        trazar_lineas_divisorias_meses(eje_1_barras, len(posiciones_x_1))

        # Texto vertical interno
        for indice, barra_obj in enumerate(barra_std_1):
            unidades = int(agrupacion_1['Cant._Prod._A1'].iloc[indice])
            if unidades > 0: 
                eje_1_barras.text(barra_obj.get_x() + barra_obj.get_width()/2, barra_obj.get_height()*0.05, f"{unidades} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_contorno_negro, zorder=4)

        # Línea de Porcentaje
        eje_1_linea.plot(posiciones_x_1, agrupacion_1['Efic_Real_Porcentaje'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='% Efic. Real', zorder=5)
        eje_1_linea.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        eje_1_linea.text(posiciones_x_1[0], 86, 'META = 85%', color='white', bbox=caja_verde, fontsize=14, fontweight='bold', zorder=10)
        
        eje_1_linea.set_ylim(0, max(120, agrupacion_1['Efic_Real_Porcentaje'].max()*1.8))
        eje_1_linea.yaxis.set_major_formatter(mtick.PercentFormatter())

        for idx, valor_pct in enumerate(agrupacion_1['Efic_Real_Porcentaje']):
            eje_1_linea.annotate(f"{valor_pct:.1f}%", (posiciones_x_1[idx], valor_pct + 5), color='white', bbox=caja_gris, ha='center', fontweight='bold', zorder=10)

        eje_1_barras.set_xticks(posiciones_x_1)
        eje_1_barras.set_xticklabels(agrupacion_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        
        eje_1_barras.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        eje_1_linea.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(figura_1)
    else: 
        st.warning("⚠️ No hay datos suficientes para graficar Eficiencia Real.")

with columna_metrica_2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_base_grafico_1.empty:
        agrupacion_2 = df_base_grafico_1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        calculo_kpi_2 = (agrupacion_2['HH_STD_TOTAL'] / agrupacion_2['HH_Productivas_C/GAP'])
        agrupacion_2['Efic_Prod_Porcentaje'] = calculo_kpi_2.replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        figura_2, eje_2_barras = plt.subplots(figsize=(14, 10))
        eje_2_linea = eje_2_barras.twinx()
        
        figura_2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_2.suptitle(texto_parametros_activos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        posiciones_x_2 = np.arange(len(agrupacion_2))
        
        # Barras
        barra_std_2 = eje_2_barras.bar(posiciones_x_2 - ancho_barra/2, agrupacion_2['HH_STD_TOTAL'], ancho_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        barra_pro_2 = eje_2_barras.bar(posiciones_x_2 + ancho_barra/2, agrupacion_2['HH_Productivas_C/GAP'], ancho_barra, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        aplicar_margen_superior_grafico(eje_2_barras, max(agrupacion_2['HH_STD_TOTAL'].max(), agrupacion_2['HH_Productivas_C/GAP'].max()), 2.6)
        
        eje_2_barras.bar_label(barra_std_2, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        eje_2_barras.bar_label(barra_pro_2, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        
        trazar_lineas_divisorias_meses(eje_2_barras, len(posiciones_x_2))

        # Línea %
        eje_2_linea.plot(posiciones_x_2, agrupacion_2['Efic_Prod_Porcentaje'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='% Efic. Prod.', zorder=5)
        eje_2_linea.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        eje_2_linea.text(posiciones_x_2[0], 101, 'META = 100%', color='white', bbox=caja_verde, fontsize=14, fontweight='bold', zorder=10)
        
        eje_2_linea.set_ylim(0, max(150, agrupacion_2['Efic_Prod_Porcentaje'].max()*1.8))
        eje_2_linea.yaxis.set_major_formatter(mtick.PercentFormatter())

        for idx, valor_pct in enumerate(agrupacion_2['Efic_Prod_Porcentaje']):
            eje_2_linea.annotate(f"{valor_pct:.1f}%", (posiciones_x_2[idx], valor_pct + 5), color='white', bbox=caja_gris, ha='center', fontweight='bold', zorder=10)

        eje_2_barras.set_xticks(posiciones_x_2)
        eje_2_barras.set_xticklabels(agrupacion_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_2_barras.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(figura_2)
    else: 
        st.warning("⚠️ No hay datos para Eficiencia Productiva.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (GAP Y COSTOS)
# =========================================================================
columna_metrica_3, columna_metrica_4 = st.columns(2)

with columna_metrica_3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Desvío entre Horas Disponibles y Horas Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_eficiencias_filtrado.empty:
        nombre_columna_productiva = 'HH_Productivas' if 'HH_Productivas' in df_eficiencias_filtrado.columns else 'HH Productivas'
        
        agrupacion_3 = df_eficiencias_filtrado.groupby('Fecha').agg({
            nombre_columna_productiva: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        agrupacion_3['Sumatoria_Declarada'] = agrupacion_3[nombre_columna_productiva] + agrupacion_3['HH_Improductivas']
        
        figura_3, eje_3_barras = plt.subplots(figsize=(14, 10))
        figura_3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_3.suptitle(texto_parametros_activos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        posiciones_x_3 = np.arange(len(agrupacion_3))
        
        # Barras Apiladas
        eje_3_barras.bar(posiciones_x_3, agrupacion_3[nombre_columna_productiva], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        eje_3_barras.bar(posiciones_x_3, agrupacion_3['HH_Improductivas'], bottom=agrupacion_3[nombre_columna_productiva], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad
        eje_3_barras.plot(posiciones_x_3, agrupacion_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='HH DISPONIBLES', zorder=5)
        
        aplicar_margen_superior_grafico(eje_3_barras, agrupacion_3['HH_Disponibles'].max(), 2.6)
        trazar_lineas_divisorias_meses(eje_3_barras, len(posiciones_x_3))

        # Visualización de GAP
        for indice_x in range(len(posiciones_x_3)):
            valor_disponible = agrupacion_3['HH_Disponibles'].iloc[indice_x]
            valor_declarado = agrupacion_3['Sumatoria_Declarada'].iloc[indice_x]
            valor_del_gap = valor_disponible - valor_declarado
            
            # Flecha/Línea vertical del GAP
            eje_3_barras.plot([indice_x, indice_x], [valor_declarado, valor_disponible], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Anotación GAP
            eje_3_barras.annotate(f"GAP:\n{int(valor_del_gap)}", (indice_x, valor_declarado + 5), color='firebrick', bbox=caja_blanca, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Anotación Disponibilidad
            eje_3_barras.annotate(f"{int(valor_disponible)}", (indice_x, valor_disponible + (eje_3_barras.get_ylim()[1]*0.08)), color='black', bbox=caja_blanca, ha='center', fontweight='bold', zorder=10)

        eje_3_barras.set_xticks(posiciones_x_3)
        eje_3_barras.set_xticklabels(agrupacion_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_3_barras.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(figura_3)
    else: 
        st.warning("⚠️ Sin datos para el cálculo de GAP.")

with columna_metrica_4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización económica de la ineficiencia productiva</i></div>", unsafe_allow_html=True)
    
    if not df_eficiencias_filtrado.empty:
        agrupacion_4 = df_eficiencias_filtrado.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        figura_4, eje_4_barras = plt.subplots(figsize=(14, 10))
        eje_4_linea = eje_4_barras.twinx()
        
        figura_4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_4.suptitle(texto_parametros_activos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        posiciones_x_4 = np.arange(len(agrupacion_4))
        
        # Barras de volumen
        barra_perdida_4 = eje_4_barras.bar(posiciones_x_4, agrupacion_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        eje_4_barras.bar_label(barra_perdida_4, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, zorder=4)
        
        aplicar_margen_superior_grafico(eje_4_barras, agrupacion_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Pesos (Eje Secundario)
        eje_4_linea.plot(posiciones_x_4, agrupacion_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_contorno_blanco, label='COSTO ARS', zorder=5)
        
        eje_4_linea.set_ylim(0, max(1000, agrupacion_4['Costo_Improd._$'].max() * 1.8))
        eje_4_linea.set_yticklabels([f'${int(val/1000000)}M' for val in eje_4_linea.get_yticks()], fontweight='bold')

        # Cartel de Costo Total
        gran_total_pesos = agrupacion_4['Costo_Improd._$'].sum()
        eje_4_barras.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${gran_total_pesos:,.0f}", transform=eje_4_barras.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=caja_amarilla, weight='bold', zorder=10)

        for indice_v, valor_ars in enumerate(agrupacion_4['Costo_Improd._$']):
            eje_4_linea.annotate(f"${valor_ars:,.0f}", (posiciones_x_4[indice_v], valor_ars + 5), color='white', bbox=caja_gris, ha='center', fontweight='bold', zorder=10)

        eje_4_barras.set_xticks(posiciones_x_4)
        eje_4_barras.set_xticklabels(agrupacion_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_4_barras.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(figura_4)
    else: 
        st.warning("⚠️ Sin información económica disponible.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (ANÁLISIS CAUSA RAÍZ)
# =========================================================================
columna_metrica_5, columna_metrica_6 = st.columns(2)

with columna_metrica_5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_improductivas_filtrado.empty:
        # Agrupación por causas
        agrupacion_5 = df_improductivas_filtrado.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor mensual
        meses_unicos_filtro = df_improductivas_filtrado['FECHA'].nunique()
        divisor_promedio = meses_unicos_filtro if meses_unicos_filtro > 0 else 1
        
        agrupacion_5['Promedio_Mensual_HH'] = agrupacion_5['HH_IMPRODUCTIVAS'] / divisor_promedio
        agrupacion_5 = agrupacion_5.sort_values(by='Promedio_Mensual_HH', ascending=False)
        agrupacion_5['Porcentaje_Acumulado_Par'] = (agrupacion_5['Promedio_Mensual_HH'].cumsum() / agrupacion_5['Promedio_Mensual_HH'].sum()) * 100

        figura_5, eje_5_barras = plt.subplots(figsize=(14, 10))
        eje_5_linea = eje_5_barras.twinx()
        
        figura_5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        figura_5.suptitle(texto_parametros_activos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        posiciones_x_5 = np.arange(len(agrupacion_5))
        
        # Barras de Pareto
        barra_pareto_5 = eje_5_barras.bar(posiciones_x_5, agrupacion_5['Promedio_Mensual_HH'], color='maroon', edgecolor='white', zorder=2)
        aplicar_margen_superior_grafico(eje_5_barras, agrupacion_5['Promedio_Mensual_HH'].max(), 2.8)
        eje_5_barras.bar_label(barra_pareto_5, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva de Lorenz (Acumulado)
        eje_5_linea.plot(posiciones_x_5, agrupacion_5['Porcentaje_Acumulado_Par'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_contorno_blanco, zorder=5)
        eje_5_linea.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        eje_5_linea.set_ylim(0, 200)
        eje_5_linea.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Ajuste de texto para eje X
        etiquetas_x_envueltas = [textwrap.fill(str(texto_parada), 12) for texto_parada in agrupacion_5['TIPO_PARADA']]
        eje_5_barras.set_xticks(posiciones_x_5)
        eje_5_barras.set_xticklabels(etiquetas_x_envueltas, rotation=90, fontsize=12, fontweight='bold')
        
        for idx_p, valor_acum in enumerate(agrupacion_5['Porcentaje_Acumulado_Par']):
            eje_5_linea.annotate(f"{valor_acum:.1f}%", (posiciones_x_5[idx_p], valor_acum + 4), color='white', bbox=caja_gris, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        # Cuadro de suma promedio
        gran_suma_promedio = agrupacion_5['Promedio_Mensual_HH'].sum()
        eje_5_barras.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{gran_suma_promedio:.1f} HH", transform=eje_5_barras.transAxes, bbox=caja_gris, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(figura_5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tabla_final_5 = agrupacion_5.copy()
        horas_improductivas_totales = df_tabla_final_5['HH_IMPRODUCTIVAS'].sum()
        df_tabla_final_5['% sobre Selección'] = (df_tabla_final_5['HH_IMPRODUCTIVAS'] / horas_improductivas_totales) * 100
        
        # INYECCIÓN EXPRESA DE FILA DE TOTAL
        fila_totalizador = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [horas_improductivas_totales], 
            'Promedio_Mensual_HH': [df_tabla_final_5['Promedio_Mensual_HH'].sum()],
            'Porcentaje_Acumulado_Par': [100.0],
            '% sobre Selección': [100.0]
        })
        df_tabla_final_5 = pd.concat([df_tabla_final_5, fila_totalizador], ignore_index=True)
        
        st.dataframe(
            df_tabla_final_5.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Causa Detectada'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Generación de CSV
        data_csv_salida = df_tabla_final_5.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=data_csv_salida, file_name="Plan_Gestion_Ombu.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Objetivo Cumplido! No se detectaron horas improductivas en los filtros seleccionados.")

with columna_metrica_6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_eficiencias_filtrado.empty:
        # Cruce de fechas
        df_eficiencias_filtrado['Llave_Mes'] = df_eficiencias_filtrado['Fecha'].dt.strftime('%Y-%m')
        agrupacion_disp_6 = df_eficiencias_filtrado.groupby('Llave_Mes', as_index=False)['HH_Disponibles'].sum()

        if not df_improductivas_filtrado.empty:
            df_improductivas_filtrado['Llave_Mes'] = df_improductivas_filtrado['FECHA'].dt.strftime('%Y-%m')
            pivot_causas_6 = pd.pivot_table(df_improductivas_filtrado, values='HH_IMPRODUCTIVAS', index='Llave_Mes', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df_grafico_6 = pd.merge(agrupacion_disp_6, pivot_causas_6, on='Llave_Mes', how='left').fillna(0)
            nombres_causas_6 = [c_nombre for c_nombre in df_grafico_6.columns if c_nombre not in ['HH_Disponibles', 'Llave_Mes']]
        else:
            df_grafico_6 = agrupacion_disp_6.copy()
            nombres_causas_6 = []
            
        # Cálculos de incidencia general
        df_grafico_6['Suma_Horas_Perdidas'] = df
