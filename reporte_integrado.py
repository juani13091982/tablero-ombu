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
# 2. SISTEMA DE SEGURIDAD CORPORATIVA (ACCESO RESTRINGIDO)
# =========================================================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_pantalla_de_login():
    """
    Despliega la interfaz inicial de seguridad para proteger los datos.
    Incluye el logotipo y el nombre oficial de la compañía.
    """
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    columna_vacia_izq, columna_login, columna_vacia_der = st.columns([1, 1.8, 1])
    
    with columna_login:
        # Franja decorativa institucional superior
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
        
        # Textos oficiales de la gerencia
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

# Bloqueo estricto si no hay sesión iniciada
if not st.session_state['autenticado']:
    mostrar_pantalla_de_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES (CSS) Y CONFIGURACIÓN DE GRÁFICOS MATPLOTLIB
# =========================================================================
codigo_css_interfaz = """
<style>
    /* Ocultar menú nativo de Streamlit para aspecto de software propio */
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
st.markdown(codigo_css_interfaz, unsafe_allow_html=True)

# Parámetros estrictos de Matplotlib para legibilidad en presentaciones gerenciales
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold',
    'legend.fontsize': 12
})

# Contornos para evitar superposición visual de textos sobre los gráficos
efecto_contorno_blanco = [path_effects.withStroke(linewidth=3, foreground='white')]
efecto_contorno_negro = [path_effects.withStroke(linewidth=3, foreground='black')]

# Estilos de cajas de texto de las etiquetas
caja_gris_oscura = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_amarilla = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
caja_verde = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
caja_blanca = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 4. FUNCIONES AUXILIARES Y MOTOR INTELIGENTE DE CRUCE
# =========================================================================
def aplicar_margen_superior_grafico(eje_matplotlib, valor_maximo, multiplicador=2.6):
    """Evita que las barras altas oculten el título o las etiquetas superiores."""
    if valor_maximo > 0: 
        eje_matplotlib.set_ylim(0, valor_maximo * multiplicador)
    else: 
        eje_matplotlib.set_ylim(0, 100)

def trazar_lineas_divisorias_meses(eje_matplotlib, cantidad_periodos):
    """Dibuja guías verticales para separar visualmente los meses en el eje X."""
    for indice in range(cantidad_periodos):
        eje_matplotlib.axvline(x=indice, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_texto_filtros(lista_seleccion, texto_por_defecto):
    """Resume los filtros seleccionados para insertarlos en los títulos."""
    if not lista_seleccion: 
        return texto_por_defecto
    if len(lista_seleccion) > 2: 
        return f"Varios ({len(lista_seleccion)})"
    return " + ".join(lista_seleccion)

def motor_analisis_robusto(texto_seleccionado, texto_excel_improductivas):
    """
    Motor de cruce Fuzzy: Permite conectar bases de datos aunque los nombres 
    de los puestos tengan pequeñas variaciones (Ej: "475-CARROZADO" vs "SECTOR CARROZADO").
    """
    if pd.isna(texto_excel_improductivas) or pd.isna(texto_seleccionado): 
        return False
    
    # 1. Normalización de caracteres y mayúsculas
    str_1 = str(texto_seleccionado).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    str_2 = str(texto_excel_improductivas).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    # 2. Limpieza total dejando solo letras y números
    alfa_1 = re.sub(r'[^A-Z0-9]', '', str_1)
    alfa_2 = re.sub(r'[^A-Z0-9]', '', str_2)
    
    if not alfa_1 or not alfa_2: 
        return False
        
    # 3. Prueba de inclusión directa
    if alfa_1 in alfa_2 or alfa_2 in alfa_1: 
        return True
    
    # 4. Prueba por código de estación (ej. extraer el '475')
    numeros_1 = set(re.findall(r'\d{3,}', str_1))
    numeros_2 = set(re.findall(r'\d{3,}', str_2))
    if numeros_1 and numeros_2 and numeros_1.intersection(numeros_2): 
        return True
        
    # 5. Prueba por palabras clave raíz
    palabras_1 = set(re.findall(r'[A-Z]{4,}', str_1))
    palabras_2 = set(re.findall(r'[A-Z]{4,}', str_2))
    
    palabras_ignoradas = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    conjunto_1_limpio = palabras_1 - palabras_ignoradas
    conjunto_2_limpio = palabras_2 - palabras_ignoradas
    
    for palabra in conjunto_1_limpio:
        if any(palabra in x for x in conjunto_2_limpio): 
            return True
                
    return False

# =========================================================================
# 5. ENCABEZADO PRINCIPAL DEL TABLERO
# =========================================================================
columna_logo, columna_titulo, columna_salida = st.columns([1, 3, 1])

with columna_logo:
    try: 
        st.image("LOGO OMBÚ.jpg", width=120)
    except: 
        st.markdown("### OMBÚ")

with columna_titulo:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>Gerencia de Control de Gestión Productiva</p>", unsafe_allow_html=True)

with columna_salida:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# =========================================================================
# 6. CARGA Y DEPURACIÓN DE DATOS DESDE EXCEL
# =========================================================================
try:
    # Lectura de los archivos fuente
    dataframe_eficiencias_raw = pd.read_excel("eficiencias.xlsx")
    dataframe_improductivas_raw = pd.read_excel("improductivas.xlsx")
    
    # Estandarización de encabezados
    dataframe_eficiencias_raw.columns = dataframe_eficiencias_raw.columns.str.strip()
    dataframe_improductivas_raw.columns = [str(c).strip().upper() for c in dataframe_improductivas_raw.columns]
    
    # Auto-identificador de columnas de Improductivas (Blindaje contra cambios de formato)
    if 'TIPO_PARADA' not in dataframe_improductivas_raw.columns:
        col_tipo = next((c for c in dataframe_improductivas_raw.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if col_tipo: 
            dataframe_improductivas_raw.rename(columns={col_tipo: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in dataframe_improductivas_raw.columns:
        col_hs_imp = next((c for c in dataframe_improductivas_raw.columns if 'HH' in c and 'IMP' in c), None)
        if col_hs_imp: 
            dataframe_improductivas_raw.rename(columns={col_hs_imp: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in dataframe_improductivas_raw.columns:
        col_fecha_imp = next((c for c in dataframe_improductivas_raw.columns if 'FECHA' in c), None)
        if col_fecha_imp: 
            dataframe_improductivas_raw.rename(columns={col_fecha_imp: 'FECHA'}, inplace=True)
    
    # Procesamiento y unificación de formato de fechas al inicio de mes
    dataframe_eficiencias_raw['Fecha'] = pd.to_datetime(dataframe_eficiencias_raw['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    dataframe_improductivas_raw['FECHA'] = pd.to_datetime(dataframe_improductivas_raw['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    # Clasificación estricta de puestos terminales
    dataframe_eficiencias_raw['Es_Ultimo_Puesto'] = dataframe_eficiencias_raw['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    # Creación de etiquetas de string para los selectores
    dataframe_eficiencias_raw['Mes_String'] = dataframe_eficiencias_raw['Fecha'].dt.strftime('%b-%Y')
    dataframe_improductivas_raw['Mes_String'] = dataframe_improductivas_raw['FECHA'].dt.strftime('%b-%Y')
    
except Exception as error_lectura:
    st.error(f"Error crítico al leer los archivos: {error_lectura}")
    st.stop()

# =========================================================================
# 7. BARRA DE FILTROS SUPERIORES (CASCADA DINÁMICA)
# =========================================================================
with st.container():
    # El ancla CSS para mantener el menú fijo arriba
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns(4)
    
    with filtro_col1: 
        lista_plantas_unicas = list(dataframe_eficiencias_raw['Planta'].dropna().unique())
        seleccion_planta = st.multiselect("🏭 Planta", lista_plantas_unicas, placeholder="Todas")
        
    df_transitorio_lineas = dataframe_eficiencias_raw[dataframe_eficiencias_raw['Planta'].isin(seleccion_planta)] if seleccion_planta else dataframe_eficiencias_raw
    
    with filtro_col2: 
        lista_lineas_unicas = list(df_transitorio_lineas['Linea'].dropna().unique())
        seleccion_linea = st.multiselect("⚙️ Línea", lista_lineas_unicas, placeholder="Todas")
        
    df_transitorio_puestos = df_transitorio_lineas[df_transitorio_lineas['Linea'].isin(seleccion_linea)] if seleccion_linea else df_transitorio_lineas
    
    with filtro_col3: 
        lista_puestos_unicos = list(df_transitorio_puestos['Puesto_Trabajo'].dropna().unique())
        seleccion_puesto = st.multiselect("🛠️ Puesto de Trabajo", lista_puestos_unicos, placeholder="Todos")
        
    with filtro_col4: 
        lista_meses_unicos = ["🎯 Acumulado YTD"] + list(dataframe_eficiencias_raw['Mes_String'].unique())
        seleccion_mes = st.multiselect("📅 Mes", lista_meses_unicos, placeholder="Todos")

# =========================================================================
# 8. APLICACIÓN DE FILTROS A LOS DATAFRAMES
# =========================================================================
dataframe_eficiencias_filtrado = dataframe_eficiencias_raw.copy()
dataframe_improductivas_filtrado = dataframe_improductivas_raw.copy()

# A) Filtrado Directo en Eficiencias
if seleccion_planta: 
    dataframe_eficiencias_filtrado = dataframe_eficiencias_filtrado[dataframe_eficiencias_filtrado['Planta'].isin(seleccion_planta)]
if seleccion_linea: 
    dataframe_eficiencias_filtrado = dataframe_eficiencias_filtrado[dataframe_eficiencias_filtrado['Linea'].isin(seleccion_linea)]
if seleccion_puesto: 
    dataframe_eficiencias_filtrado = dataframe_eficiencias_filtrado[dataframe_eficiencias_filtrado['Puesto_Trabajo'].isin(seleccion_puesto)]
if seleccion_mes and "🎯 Acumulado YTD" not in seleccion_mes:
    dataframe_eficiencias_filtrado = dataframe_eficiencias_filtrado[dataframe_eficiencias_filtrado['Mes_String'].isin(seleccion_mes)]

# B) Filtrado Robusto en Improductivas (BLINDADO CONTRA INDEX ERROR)
if seleccion_planta and not dataframe_improductivas_filtrado.empty:
    columna_planta_buscada = next((col for col in dataframe_improductivas_filtrado.columns if 'PLANTA' in str(col).upper()), None)
    if not columna_planta_buscada and len(dataframe_improductivas_filtrado.columns) > 0:
        columna_planta_buscada = dataframe_improductivas_filtrado.columns[0]
        
    if columna_planta_buscada:
        mascara_planta = dataframe_improductivas_filtrado[columna_planta_buscada].apply(
            lambda valor_celda: any(motor_analisis_robusto(p_sel, valor_celda) for p_sel in seleccion_planta)
        )
        dataframe_improductivas_filtrado = dataframe_improductivas_filtrado[mascara_planta]

if seleccion_linea and not dataframe_improductivas_filtrado.empty:
    columna_linea_buscada = next((col for col in dataframe_improductivas_filtrado.columns if 'LINEA' in str(col).upper()), None)
    if not columna_linea_buscada and len(dataframe_improductivas_filtrado.columns) > 1:
        columna_linea_buscada = dataframe_improductivas_filtrado.columns[1]
        
    if columna_linea_buscada:
        mascara_linea = dataframe_improductivas_filtrado[columna_linea_buscada].apply(
            lambda valor_celda: any(motor_analisis_robusto(l_sel, valor_celda) for l_sel in seleccion_linea)
        )
        dataframe_improductivas_filtrado = dataframe_improductivas_filtrado[mascara_linea]

if seleccion_puesto and not dataframe_improductivas_filtrado.empty:
    columna_puesto_buscada = next((col for col in dataframe_improductivas_filtrado.columns if 'PUESTO' in str(col).upper()), None)
    if not columna_puesto_buscada and len(dataframe_improductivas_filtrado.columns) > 2:
        columna_puesto_buscada = dataframe_improductivas_filtrado.columns[2]
        
    if columna_puesto_buscada:
        mascara_puesto = dataframe_improductivas_filtrado[columna_puesto_buscada].apply(
            lambda valor_celda: any(motor_analisis_robusto(ps_sel, valor_celda) for ps_sel in seleccion_puesto)
        )
        dataframe_improductivas_filtrado = dataframe_improductivas_filtrado[mascara_puesto]

if seleccion_mes and "🎯 Acumulado YTD" not in seleccion_mes and not dataframe_improductivas_filtrado.empty:
    dataframe_improductivas_filtrado = dataframe_improductivas_filtrado[dataframe_improductivas_filtrado['Mes_String'].isin(seleccion_mes)]

# String dinámico para incluir en los títulos de cada gráfico
texto_encabezado_graficos = f"Filtros Activos >> Planta: {formatear_texto_filtros(seleccion_planta, 'Todas')} | Línea: {formatear_texto_filtros(seleccion_linea, 'Todas')} | Puesto: {formatear_texto_filtros(seleccion_puesto, 'Todos')}"

st.markdown("---")

# =========================================================================
# 9. FILA 1: MÉTRICAS 1 Y 2
# =========================================================================
columna_metrica_1, columna_metrica_2 = st.columns(2)

with columna_metrica_1:
    st.header("1. EFICIENCIA REAL")
    # ESPECIFICACIÓN EXPLÍCITA DE LA FÓRMULA SOLICITADA
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    # Lógica de terminal: si no hay puesto filtrado, tomar solo 'SI'
    df_base_grafico_1 = dataframe_eficiencias_filtrado.copy() if seleccion_puesto else dataframe_eficiencias_filtrado[dataframe_eficiencias_filtrado['Es_Ultimo_Puesto'] == 'SI']
    
    if not df_base_grafico_1.empty:
        agrupacion_metrica_1 = df_base_grafico_1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        # Cálculo del indicador
        calculo_eficiencia_real = (agrupacion_metrica_1['HH_STD_TOTAL'] / agrupacion_metrica_1['HH_Disponibles'])
        agrupacion_metrica_1['Efic_Real_Porcentaje'] = calculo_eficiencia_real.replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        # Construcción del Gráfico
        figura_1, eje_izq_1 = plt.subplots(figsize=(14, 10))
        eje_der_1 = eje_izq_1.twinx()
        
        figura_1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_1.suptitle(texto_encabezado_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        array_posiciones_x_1 = np.arange(len(agrupacion_metrica_1))
        ancho_de_barra = 0.35
        
        # Barras
        barra_std_1 = eje_izq_1.bar(array_posiciones_x_1 - ancho_de_barra/2, agrupacion_metrica_1['HH_STD_TOTAL'], ancho_de_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        barra_dis_1 = eje_izq_1.bar(array_posiciones_x_1 + ancho_de_barra/2, agrupacion_metrica_1['HH_Disponibles'], ancho_de_barra, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        aplicar_margen_superior_grafico(eje_izq_1, agrupacion_metrica_1['HH_Disponibles'].max(), 2.6)
        
        # Etiquetas de Barras
        eje_izq_1.bar_label(barra_std_1, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        eje_izq_1.bar_label(barra_dis_1, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        
        trazar_lineas_divisorias_meses(eje_izq_1, len(array_posiciones_x_1))

        # Texto vertical de Unidades
        for indice, barra_objeto in enumerate(barra_std_1):
            unidades_producidas = int(agrupacion_metrica_1['Cant._Prod._A1'].iloc[indice])
            if unidades_producidas > 0: 
                eje_izq_1.text(barra_objeto.get_x() + barra_objeto.get_width()/2, barra_objeto.get_height()*0.05, f"{unidades_producidas} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_contorno_negro, zorder=4)

        # Línea de Porcentaje
        eje_der_1.plot(array_posiciones_x_1, agrupacion_metrica_1['Efic_Real_Porcentaje'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='% Efic. Real', zorder=5)
        
        # Línea de Meta 85%
        eje_der_1.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        eje_der_1.text(array_posiciones_x_1[0], 86, 'META = 85%', color='white', bbox=caja_verde, fontsize=14, fontweight='bold', zorder=10)
        
        eje_der_1.set_ylim(0, max(120, agrupacion_metrica_1['Efic_Real_Porcentaje'].max()*1.8))
        eje_der_1.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Anotación sobre la curva
        for idx, valor_pct in enumerate(agrupacion_metrica_1['Efic_Real_Porcentaje']):
            eje_der_1.annotate(f"{valor_pct:.1f}%", (array_posiciones_x_1[idx], valor_pct + 5), color='white', bbox=caja_gris_oscura, ha='center', fontweight='bold', zorder=10)

        eje_izq_1.set_xticks(array_posiciones_x_1)
        eje_izq_1.set_xticklabels(agrupacion_metrica_1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        
        eje_izq_1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        eje_der_1.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(figura_1)
    else: 
        st.warning("⚠️ No hay datos suficientes para graficar Eficiencia Real.")

with columna_metrica_2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_base_grafico_1.empty:
        agrupacion_metrica_2 = df_base_grafico_1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum'
        }).reset_index()
        
        calculo_eficiencia_prod = (agrupacion_metrica_2['HH_STD_TOTAL'] / agrupacion_metrica_2['HH_Productivas_C/GAP'])
        agrupacion_metrica_2['Efic_Prod_Porcentaje'] = calculo_eficiencia_prod.replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        figura_2, eje_izq_2 = plt.subplots(figsize=(14, 10))
        eje_der_2 = eje_izq_2.twinx()
        
        figura_2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_2.suptitle(texto_encabezado_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        array_posiciones_x_2 = np.arange(len(agrupacion_metrica_2))
        
        # Barras Comparativas
        barra_std_2 = eje_izq_2.bar(array_posiciones_x_2 - ancho_de_barra/2, agrupacion_metrica_2['HH_STD_TOTAL'], ancho_de_barra, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        barra_pro_2 = eje_izq_2.bar(array_posiciones_x_2 + ancho_de_barra/2, agrupacion_metrica_2['HH_Productivas_C/GAP'], ancho_de_barra, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        aplicar_margen_superior_grafico(eje_izq_2, max(agrupacion_metrica_2['HH_STD_TOTAL'].max(), agrupacion_metrica_2['HH_Productivas_C/GAP'].max()), 2.6)
        
        eje_izq_2.bar_label(barra_std_2, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        eje_izq_2.bar_label(barra_pro_2, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, fmt='%.0f', zorder=3)
        
        trazar_lineas_divisorias_meses(eje_izq_2, len(array_posiciones_x_2))

        # Línea de Porcentaje Productivo
        eje_der_2.plot(array_posiciones_x_2, agrupacion_metrica_2['Efic_Prod_Porcentaje'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='% Efic. Prod.', zorder=5)
        
        # Meta 100%
        eje_der_2.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        eje_der_2.text(array_posiciones_x_2[0], 101, 'META = 100%', color='white', bbox=caja_verde, fontsize=14, fontweight='bold', zorder=10)
        
        eje_der_2.set_ylim(0, max(150, agrupacion_metrica_2['Efic_Prod_Porcentaje'].max()*1.8))
        eje_der_2.yaxis.set_major_formatter(mtick.PercentFormatter())

        for idx, valor_pct in enumerate(agrupacion_metrica_2['Efic_Prod_Porcentaje']):
            eje_der_2.annotate(f"{valor_pct:.1f}%", (array_posiciones_x_2[idx], valor_pct + 5), color='white', bbox=caja_gris_oscura, ha='center', fontweight='bold', zorder=10)

        eje_izq_2.set_xticks(array_posiciones_x_2)
        eje_izq_2.set_xticklabels(agrupacion_metrica_2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_izq_2.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
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
    
    if not dataframe_eficiencias_filtrado.empty:
        # Prevención de error por nombre de columna
        nombre_col_productiva = 'HH_Productivas' if 'HH_Productivas' in dataframe_eficiencias_filtrado.columns else 'HH Productivas'
        
        agrupacion_metrica_3 = dataframe_eficiencias_filtrado.groupby('Fecha').agg({
            nombre_col_productiva: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        agrupacion_metrica_3['Sumatoria_Declarada_Global'] = agrupacion_metrica_3[nombre_col_productiva] + agrupacion_metrica_3['HH_Improductivas']
        
        figura_3, eje_unico_3 = plt.subplots(figsize=(14, 10))
        figura_3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_3.suptitle(texto_encabezado_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        array_posiciones_x_3 = np.arange(len(agrupacion_metrica_3))
        
        # Barras Apiladas (Productivas + Improductivas)
        eje_unico_3.bar(array_posiciones_x_3, agrupacion_metrica_3[nombre_col_productiva], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        eje_unico_3.bar(array_posiciones_x_3, agrupacion_metrica_3['HH_Improductivas'], bottom=agrupacion_metrica_3[nombre_col_productiva], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        # Línea Diamante de Disponibilidad Real
        eje_unico_3.plot(array_posiciones_x_3, agrupacion_metrica_3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_contorno_blanco, label='HH DISPONIBLES', zorder=5)
        
        aplicar_margen_superior_grafico(eje_unico_3, agrupacion_metrica_3['HH_Disponibles'].max(), 2.6)
        trazar_lineas_divisorias_meses(eje_unico_3, len(array_posiciones_x_3))

        # Visualización matemática del GAP
        for indice_x in range(len(array_posiciones_x_3)):
            valor_disponible_max = agrupacion_metrica_3['HH_Disponibles'].iloc[indice_x]
            valor_declarado_total = agrupacion_metrica_3['Sumatoria_Declarada_Global'].iloc[indice_x]
            brecha_calculada = valor_disponible_max - valor_declarado_total
            
            # Flecha vertical conector del GAP
            eje_unico_3.plot([indice_x, indice_x], [valor_declarado_total, valor_disponible_max], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            # Anotación GAP
            eje_unico_3.annotate(f"GAP:\n{int(brecha_calculada)}", (indice_x, valor_declarado_total + 5), color='firebrick', bbox=caja_blanca, ha='center', va='bottom', fontweight='bold', zorder=10)
            
            # Anotación Disponibilidad Top
            eje_unico_3.annotate(f"{int(valor_disponible_max)}", (indice_x, valor_disponible_max + (eje_unico_3.get_ylim()[1]*0.08)), color='black', bbox=caja_blanca, ha='center', fontweight='bold', zorder=10)

        eje_unico_3.set_xticks(array_posiciones_x_3)
        eje_unico_3.set_xticklabels(agrupacion_metrica_3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_unico_3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(figura_3)
    else: 
        st.warning("⚠️ Sin datos para el cálculo de GAP.")

with columna_metrica_4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Valorización económica de la ineficiencia productiva</i></div>", unsafe_allow_html=True)
    
    if not dataframe_eficiencias_filtrado.empty:
        agrupacion_metrica_4 = dataframe_eficiencias_filtrado.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        figura_4, eje_izq_4 = plt.subplots(figsize=(14, 10))
        eje_der_4 = eje_izq_4.twinx()
        
        figura_4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        figura_4.suptitle(texto_encabezado_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        array_posiciones_x_4 = np.arange(len(agrupacion_metrica_4))
        
        # Barras de volumen de pérdida
        barra_perdida_4 = eje_izq_4.bar(array_posiciones_x_4, agrupacion_metrica_4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        eje_izq_4.bar_label(barra_perdida_4, padding=4, color='black', fontweight='bold', path_effects=efecto_contorno_blanco, zorder=4)
        
        aplicar_margen_superior_grafico(eje_izq_4, agrupacion_metrica_4['HH_Improductivas'].max(), 2.6)
        
        # Línea de Pesos (Costo Monetario)
        eje_der_4.plot(array_posiciones_x_4, agrupacion_metrica_4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_contorno_blanco, label='COSTO ARS', zorder=5)
        
        eje_der_4.set_ylim(0, max(1000, agrupacion_metrica_4['Costo_Improd._$'].max() * 1.8))
        eje_der_4.set_yticklabels([f'${int(val/1000000)}M' for val in eje_der_4.get_yticks()], fontweight='bold')

        # Cartel de Resumen Financiero Total
        suma_total_pesos = agrupacion_metrica_4['Costo_Improd._$'].sum()
        eje_izq_4.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${suma_total_pesos:,.0f}", transform=eje_izq_4.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=caja_amarilla, weight='bold', zorder=10)

        for indice_v, valor_ars in enumerate(agrupacion_metrica_4['Costo_Improd._$']):
            eje_der_4.annotate(f"${valor_ars:,.0f}", (array_posiciones_x_4[indice_v], valor_ars + 5), color='white', bbox=caja_gris_oscura, ha='center', fontweight='bold', zorder=10)

        eje_izq_4.set_xticks(array_posiciones_x_4)
        eje_izq_4.set_xticklabels(agrupacion_metrica_4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        eje_izq_4.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        
        st.pyplot(figura_4)
    else: 
        st.warning("⚠️ Sin información económica disponible en la base.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (ANÁLISIS CAUSA RAÍZ)
# =========================================================================
columna_metrica_5, columna_metrica_6 = st.columns(2)

with columna_metrica_5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not dataframe_improductivas_filtrado.empty:
        # Agrupación por causas para el Pareto
        agrupacion_metrica_5 = dataframe_improductivas_filtrado.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        # Divisor de normalización mensual
        cantidad_meses_unicos = dataframe_improductivas_filtrado['FECHA'].nunique()
        divisor_promedio_mensual = cantidad_meses_unicos if cantidad_meses_unicos > 0 else 1
        
        agrupacion_metrica_5['Promedio_Mensual_Kpi'] = agrupacion_metrica_5['HH_IMPRODUCTIVAS'] / divisor_promedio_mensual
        agrupacion_metrica_5 = agrupacion_metrica_5.sort_values(by='Promedio_Mensual_Kpi', ascending=False)
        agrupacion_metrica_5['Porcentaje_Acumulado_Curva'] = (agrupacion_metrica_5['Promedio_Mensual_Kpi'].cumsum() / agrupacion_metrica_5['Promedio_Mensual_Kpi'].sum()) * 100

        figura_5, eje_izq_5 = plt.subplots(figsize=(14, 10))
        eje_der_5 = eje_izq_5.twinx()
        
        figura_5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        figura_5.suptitle(texto_encabezado_graficos, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        array_posiciones_x_5 = np.arange(len(agrupacion_metrica_5))
        
        # Barras de Pareto (Orden Descendente)
        barra_pareto_principal = eje_izq_5.bar(array_posiciones_x_5, agrupacion_metrica_5['Promedio_Mensual_Kpi'], color='maroon', edgecolor='white', zorder=2)
        aplicar_margen_superior_grafico(eje_izq_5, agrupacion_metrica_5['Promedio_Mensual_Kpi'].max(), 2.8)
        eje_izq_5.bar_label(barra_pareto_principal, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        # Curva de Lorenz (Impacto Acumulado)
        eje_der_5.plot(array_posiciones_x_5, agrupacion_metrica_5['Porcentaje_Acumulado_Curva'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_contorno_blanco, zorder=5)
        eje_der_5.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        eje_der_5.set_ylim(0, 200)
        eje_der_5.yaxis.set_major_formatter(mtick.PercentFormatter())

        # Ajuste de texto para las descripciones en el eje X
        etiquetas_eje_x_envueltas = [textwrap.fill(str(texto_motivo), 12) for texto_motivo in agrupacion_metrica_5['TIPO_PARADA']]
        eje_izq_5.set_xticks(array_posiciones_x_5)
        eje_izq_5.set_xticklabels(etiquetas_eje_x_envueltas, rotation=90, fontsize=12, fontweight='bold')
        
        for idx_pareto, valor_acumulado in enumerate(agrupacion_metrica_5['Porcentaje_Acumulado_Curva']):
            eje_der_5.annotate(f"{valor_acumulado:.1f}%", (array_posiciones_x_5[idx_pareto], valor_acumulado + 4), color='white', bbox=caja_gris_oscura, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        # Cuadro de suma promedio mensual
        gran_suma_promedio_hs = agrupacion_metrica_5['Promedio_Mensual_Kpi'].sum()
        eje_izq_5.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{gran_suma_promedio_hs:.1f} HH", transform=eje_izq_5.transAxes, bbox=caja_gris_oscura, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        st.pyplot(figura_5)
        
        # ==========================================
        # TABLA DE MESA DE TRABAJO Y DESCARGA
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        dataframe_tabla_impacto = agrupacion_metrica_5.copy()
        totalizador_horas_improductivas = dataframe_tabla_impacto['HH_IMPRODUCTIVAS'].sum()
        dataframe_tabla_impacto['% sobre Selección'] = (dataframe_tabla_impacto['HH_IMPRODUCTIVAS'] / totalizador_horas_improductivas) * 100
        
        # INYECCIÓN EXPRESA DE FILA DE TOTAL ABSOLUTO
        fila_total_matriz = pd.DataFrame({
            'TIPO_PARADA': ['✅ TOTAL'], 
            'HH_IMPRODUCTIVAS': [totalizador_horas_improductivas], 
            'Promedio_Mensual_Kpi': [dataframe_tabla_impacto['Promedio_Mensual_Kpi'].sum()],
            'Porcentaje_Acumulado_Curva': [100.0],
            '% sobre Selección': [100.0]
        })
        dataframe_tabla_impacto = pd.concat([dataframe_tabla_impacto, fila_total_matriz], ignore_index=True)
        
        st.dataframe(
            dataframe_tabla_impacto.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Causa Raíz Detectada'}), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"),
                "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Generación de archivo CSV para plan de acción
        cadena_csv_exportable = dataframe_tabla_impacto.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Plan de Acción (CSV)", data=cadena_csv_exportable, file_name="Plan_Gestion_Industrial.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.success("✅ ¡Objetivo Cumplido! No se detectaron horas improductivas en la combinación de filtros seleccionada.")

with columna_metrica_6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 14px; color: #aaa;'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not dataframe_eficiencias_filtrado.empty:
        # Cruce de fechas entre ambos dataframes
        dataframe_eficiencias_filtrado['Llave_Periodo'] = dataframe_eficiencias_filtrado['Fecha'].dt.strftime('%Y-%m')
        agrupacion_disponibilidad_6 = dataframe_eficiencias_filtrado.groupby('Llave_Periodo', as_index=False)['HH_Disponibles'].sum()

        if not dataframe_improductivas_filtrado.empty:
            dataframe_improductivas_filtrado['Llave_Periodo'] = dataframe_improductivas_filtrado['FECHA'].dt.strftime('%Y-%m')
            
            # Matriz pivot para abrir las causas como apiladas
            matriz_pivot_causas = pd.pivot_table(dataframe_improductivas_filtrado, values='HH_IMPRODUCTIVAS', index='Llave_Periodo', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            dataframe_grafico_6 = pd.merge(agrupacion_disponibilidad_6, matriz_pivot_causas, on='Llave_Periodo', how='left').fillna(0)
            nombres_columnas_causas = [nombre_col for nombre_col in dataframe_grafico_6.columns if nombre_col not in ['HH_Disponibles', 'Llave_Periodo']]
        else:
            dataframe_grafico_6 = agrupacion_disponibilidad_6.copy()
            nombres_columnas_causas = []
            
        # Cálculos de la métrica de incidencia general
        dataframe_grafico_6['Suma_Hs_Perdidas'] = dataframe_grafico_6[nombres_columnas_causas].sum(axis=1) if nombres_columnas_causas else 0
        dataframe_grafico_6['Incidencia_Calculada_Pct'] = (dataframe_grafico_6['Suma_Hs_Perdidas'] / dataframe_grafico_6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        # Ordenamiento cronológico forzado
        dataframe_grafico_6['Fecha_Sort'] = pd.to_datetime(dataframe_grafico_6['Ll
