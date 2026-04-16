import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as path_effects
import textwrap
import datetime
import os

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS GLOBALES
# ==========================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide")

# ESCUDO DE INVISIBILIDAD Y PANEL INMÓVIL (STICKY) UNIFICADO
# CSS Avanzado para romper el bloqueo de scroll de Streamlit y alinear perfectamente.
css_styles = """
<style>
/* 1. Forzar que todos los contenedores padres permitan elementos sticky (Inmóviles) */
.stApp, .main, .block-container, [data-testid="stAppViewBlockContainer"], [data-testid="stMain"] {
    overflow: visible !important;
}

/* 2. Seleccionar el contenedor de los 4 filtros (es el segundo bloque horizontal) y fijarlo */
div[data-testid="stHorizontalBlock"]:nth-of-type(2) {
    position: -webkit-sticky !important;
    position: sticky !important;
    top: 0px !important;
    background-color: #0E1117 !important; /* Fondo oscuro para tapar los gráficos que suben */
    z-index: 99999 !important;
    padding: 10px 5px 15px 5px !important;
    border-bottom: 3px solid #1E3A8A !important; /* Línea azul corporativa */
    box-shadow: 0px 10px 15px -3px rgba(0,0,0,0.8) !important; /* Sombra elegante */
    border-radius: 0px 0px 10px 10px !important;
}
"""

# 3. Lógica Inteligente de Administrador
if "admin" not in st.query_params:
    # Modo Público: Ocultar todo el header y menús
    css_styles += """
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    """
else:
    # Modo Administrador (Tú): Bajar un poco el panel sticky para que no tape la barra superior de Streamlit
    css_styles += """
    div[data-testid="stHorizontalBlock"]:nth-of-type(2) {
        top: 55px !important; 
    }
    """

css_styles += "</style>"

# Inyectamos todo el CSS de una sola vez
st.markdown(css_styles, unsafe_allow_html=True)

# Regla Innegociable: Tamaños de fuente grandes y en negrita
plt.rcParams.update({
    'font.size': 12,
    'font.weight': 'bold',
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold',
    'figure.titlesize': 16,
    'figure.titleweight': 'bold'
})

# Regla Innegociable: Efectos de contorno (Anti-Overlap de texto) y bboxes
outline_white = [path_effects.withStroke(linewidth=3, foreground='white')]
outline_black = [path_effects.withStroke(linewidth=3, foreground='black')]
bbox_gray = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_yellow = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_red = dict(boxstyle="round,pad=0.3", fc="firebrick", ec="white", lw=1.5)
bbox_white = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# ==========================================
# FUNCIONES AUXILIARES ESTRICTAS (ANTI-SOLAPAMIENTO EXTREMO)
# ==========================================
def aplicar_anti_overlap(ax, max_val, multiplier=2.0):
    """Regla Crítica: Empuja las barras hacia abajo (multiplier 2.0 por defecto) 
       para liberar la mitad superior del gráfico y que NADA se solape con las líneas/carteles."""
    if max_val > 0:
        ax.set_ylim(0, max_val * multiplier)
    else:
        ax.set_ylim(0, 100) # Fallback

def dibujar_meses(ax, fechas):
    """Líneas divisorias verticales mensuales tenues"""
    for x in range(len(fechas)):
        ax.axvline(x=x, color='lightgray', linestyle='--', linewidth=1, zorder=0)

# ==========================================
# HEADER: IDENTIDAD CORPORATIVA
# ==========================================
col_logo, col_title = st.columns([1, 4])
with col_logo:
    # Intenta cargar la imagen real del logo subida a GitHub
    try:
        st.image("LOGO OMBÚ.jpg", use_container_width=True)
    except Exception:
        # Fallback
        st.markdown("""
            <div style='background-color:#1E3A8A; color:white; padding:15px; border-radius:10px; text-align:center; height: 100%; display:flex; flex-direction:column; justify-content:center;'>
                <h1 style='margin:0; font-size:32px; font-weight:bold; letter-spacing: 2px;'>OMBÚ</h1>
            </div>
        """, unsafe_allow_html=True)
with col_title:
    st.title("TABLERO MÉTRICAS - C.G.P. REPORTE Integrado")
    st.subheader("Control de Gestión Productiva")

# ==========================================
# MODO AUTOMÁTICO VS MANUAL (CARGA DE DATOS)
# ==========================================
ruta_ef = "eficiencias.xlsx"
ruta_imp = "improductivas.xlsx"

if os.path.exists(ruta_ef) and os.path.exists(ruta_imp):
    try:
        df_ef = pd.read_excel(ruta_ef)
        df_imp = pd.read_excel(ruta_imp)
    except Exception as e:
        st.error(f"Error leyendo los Excel del servidor: {e}")
        st.stop()
else:
    st.sidebar.header("📁 Carga de Datos")
    st.sidebar.info("Modo Manual activo. Faltan archivos en servidor.")
    archivo_eficiencias = st.sidebar.file_uploader("Base Eficiencias (CSV/Excel)", type=['csv', 'xlsx'])
    archivo_improductivas = st.sidebar.file_uploader("Base Hrs Improductivas (CSV/Excel)", type=['csv', 'xlsx'])
    
    if archivo_eficiencias is None or archivo_improductivas is None:
        st.info("👋 Sube los archivos en el panel izquierdo para comenzar.")
        st.stop()
    
    try:
        df_ef = pd.read_csv(archivo_eficiencias) if archivo_eficiencias.name.endswith('.csv') else pd.read_excel(archivo_eficiencias)
        df_imp = pd.read_csv(archivo_improductivas) if archivo_improductivas.name.endswith('.csv') else pd.read_excel(archivo_improductivas)
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

# ==========================================
# LIMPIEZA DE DATOS Y GENERACIÓN DE FILTROS
# ==========================================
try:
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_imp['FECHA'] = pd.to_datetime(df_imp['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    df_ef['Mes_Filtro'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_imp['Mes_Filtro'] = df_imp['FECHA'].dt.strftime('%b-%Y')
except Exception as e:
    st.error(f"Error al estandarizar fechas: {e}")
    st.stop()

st.markdown("### 🔍 Configuración del Escenario")

# ==========================================
# FILTROS EN LA PARTE SUPERIOR (RIBBON CASCADA INMÓVIL MULTIPLE)
# ==========================================
col_f1, col_f2, col_f3, col_f4 = st.columns(4)

# 1. Filtro Planta
with col_f1:
    plantas = list(df_ef['Planta'].dropna().unique())
    planta_sel = st.multiselect("🏭 Planta", ["Todas"] + plantas, default=["Todas"])

# Lógica Cascada 1
plantas_filtrar = plantas if "Todas" in planta_sel or not planta_sel else planta_sel
df_temp_linea = df_ef[df_ef['Planta'].isin(plantas_filtrar)]

# 2. Filtro Línea
with col_f2:
    lineas = list(df_temp_linea['Linea'].dropna().unique())
    linea_sel = st.multiselect("⚙️ Línea", ["Todas"] + lineas, default=["Todas"])

# Lógica Cascada 2
lineas_filtrar = lineas if "Todas" in linea_sel or not linea_sel else linea_sel
df_temp_puesto = df_temp_linea[df_temp_linea['Linea'].isin(lineas_filtrar)]

# 3. Filtro Puesto
with col_f3:
    puestos = list(df_temp_puesto['Puesto_Trabajo'].dropna().unique())
    puesto_sel = st.multiselect("🛠️ Puesto de Trabajo", ["Todos"] + puestos, default=["Todos"])

# Lógica Cascada 3
puestos_filtrar = puestos if "Todos" in puesto_sel or not puesto_sel else puesto_sel
df_temp_mes = df_temp_puesto[df_temp_puesto['Puesto_Trabajo'].isin(puestos_filtrar)]

# 4. Filtro Mes
with col_f4:
    meses_disponibles = list(df_temp_mes['Mes_Filtro'].dropna().unique())
    mes_sel = st.multiselect("📅 Mes", ["Todos"] + meses_disponibles, default=["Todos"])

# ==========================================
# APLICACIÓN DE FILTROS MATEMÁTICOS A LAS BASES
# ==========================================
df_ef_filtrado = df_ef.copy()
df_imp_filtrado = df_imp.copy()

# Eficiencias
if "Todas" not in planta_sel and planta_sel: 
    df_ef_filtrado = df_ef_filtrado[df_ef_filtrado['Planta'].isin(planta_sel)]
if "Todas" not in linea_sel and linea_sel: 
    df_ef_filtrado = df_ef_filtrado[df_ef_filtrado['Linea'].isin(linea_sel)]
if "Todos" not in puesto_sel and puesto_sel: 
    df_ef_filtrado = df_ef_filtrado[df_ef_filtrado['Puesto_Trabajo'].isin(puesto_sel)]
if "Todos" not in mes_sel and mes_sel: 
    df_ef_filtrado = df_ef_filtrado[df_ef_filtrado['Mes_Filtro'].isin(mes_sel)]

# Improductivas
if "Todas" not in planta_sel and planta_sel and 'PLANTA' in df_imp_filtrado.columns: 
    df_imp_filtrado = df_imp_filtrado[df_imp_filtrado['PLANTA'].isin(planta_sel)]
if "Todas" not in linea_sel and linea_sel and 'LÍNEA' in df_imp_filtrado.columns: 
    df_imp_filtrado = df_imp_filtrado[df_imp_filtrado['LÍNEA'].isin(linea_sel)]
if "Todos" not in puesto_sel and puesto_sel and 'PUESTOS' in df_imp_filtrado.columns: 
    df_imp_filtrado = df_imp_filtrado[df_imp_filtrado['PUESTOS'].isin(puesto_sel)]
if "Todos" not in mes_sel and mes_sel and 'Mes_Filtro' in df_imp_filtrado.columns: 
    df_imp_filtrado = df_imp_filtrado[df_imp_filtrado['Mes_Filtro'].isin(mes_sel)]

st.markdown("---")

# =========================================================================
# MÉTRICA 1: % EFICIENCIA REAL
# =========================================================================
st.header("1. % EFICIENCIA REAL")

puestos_especificos = (len(puesto_sel) > 0 and "Todos" not in puesto_sel)

if puestos_especificos:
    st.markdown("*Evaluando métricas exactas de los puestos seleccionados (Anula regla de última estación de línea).*")
    df_m1 = df_ef_filtrado.copy()
else:
    st.markdown("*Lógica Estricta de Línea: Sumatoria(HH STD) / Sumatoria(HH Disponibles) SOLO en última estación.*")
    df_m1 = df_ef_filtrado[df_ef_filtrado['Es_Ultimo_Puesto'] == 'SI'].copy()

if not df_m1.empty:
    agrup_m1 = df_m1.groupby('Fecha').agg({
        'HH_STD_TOTAL': 'sum',
        'HH_Disponibles': 'sum',
        'Cant._Prod._A1': 'sum'
    }).reset_index()
    
    agrup_m1['Eficiencia_Real'] = (agrup_m1['HH_STD_TOTAL'] / agrup_m1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
    agrup_m1['Fecha_str'] = agrup_m1['Fecha'].dt.strftime('%b-%y')

    fig1, ax1 = plt.subplots(figsize=(14, 7))
    ax2 = ax1.twinx()
    x_indexes = np.arange(len(agrup_m1))
    width = 0.35
    
    bars_std = ax1.bar(x_indexes - width/2, agrup_m1['HH_STD_TOTAL'], width, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
    bars_disp = ax1.bar(x_indexes + width/2, agrup_m1['HH_Disponibles'], width, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
    
    # Empujar barras a la mitad inferior (Multiplier 2.0)
    aplicar_anti_overlap(ax1, agrup_m1['HH_Disponibles'].max(), multiplier=2.0)
    
    ax1.bar_label(bars_std, padding=4, color='black', fontweight='bold', fontsize=12, path_effects=outline_white, fmt='%.0f', zorder=3)
    ax1.bar_label(bars_disp, padding=4, color='black', fontweight='bold', fontsize=12, path_effects=outline_white, fmt='%.0f', zorder=3)

    dibujar_meses(ax1, x_indexes)

    for i, bar in enumerate(bars_std):
        cant = int(agrup_m1['Cant._Prod._A1'].iloc[i])
        if cant > 0:
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() * 0.05, f"{cant} UND", 
                     rotation=90, color='white', ha='center', va='bottom', fontsize=16, fontweight='bold', path_effects=outline_black, zorder=4)

    ax2.plot(x_indexes, agrup_m1['Eficiencia_Real'], color='dimgray', marker='o', markersize=10, linewidth=4, path_effects=outline_white, label='% Efic. Real', zorder=5)
    ax2.axhline(y=85, color='forestgreen', linestyle='--', linewidth=3, label='Meta (85%)', zorder=1)
    
    # Elevar límite secundario para que la línea respire (Multiplier 1.5)
    max_ef_real = agrup_m1['Eficiencia_Real'].max()
    ax2.set_ylim(0, max(120, max_ef_real * 1.5))
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

    if len(x_indexes) > 1:
        z = np.polyfit(x_indexes, agrup_m1['Eficiencia_Real'], 1)
        p = np.poly1d(z)
        ax2.plot(x_indexes, p(x_indexes), color='dimgray', linestyle=':', alpha=0.8, linewidth=2, zorder=1)

    offset_y2 = ax2.get_ylim()[1] * 0.06
    for i, val in enumerate(agrup_m1['Eficiencia_Real']):
        ax2.annotate(f"{val:.1f}%", (x_indexes[i], val + offset_y2), color='white', bbox=bbox_gray, ha='center', fontsize=12, fontweight='bold', zorder=10)

    ax1.set_xticks(x_indexes)
    ax1.set_xticklabels(agrup_m1['Fecha_str'], fontsize=12, fontweight='bold')
    ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.15), ncol=2)
    ax2.legend(loc='upper right', bbox_to_anchor=(1, 1.15))
    st.pyplot(fig1)
else:
    st.warning("⚠️ No hay datos evaluables para la selección actual. Revisa los filtros aplicados.")

st.markdown("---")

# =========================================================================
# MÉTRICA 2: % EFICIENCIA PRODUCTIVA
# =========================================================================
st.header("2. % EFICIENCIA PRODUCTIVA")

if not df_m1.empty:
    agrup_m2 = df_m1.groupby('Fecha').agg({
        'HH_STD_TOTAL': 'sum',
        'HH_Productivas_C/GAP': 'sum',
        'Cant._Prod._A1': 'sum'
    }).reset_index()
    
    agrup_m2['Eficiencia_Prod'] = (agrup_m2['HH_STD_TOTAL'] / agrup_m2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
    agrup_m2['Fecha_str'] = agrup_m2['Fecha'].dt.strftime('%b-%y')

    fig2, ax1 = plt.subplots(figsize=(14, 7))
    ax2 = ax1.twinx()
    x_indexes = np.arange(len(agrup_m2))
    
    bars_std2 = ax1.bar(x_indexes - width/2, agrup_m2['HH_STD_TOTAL'], width, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
    bars_prod2 = ax1.bar(x_indexes + width/2, agrup_m2['HH_Productivas_C/GAP'], width, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
    
    # Empujar barras a la mitad inferior
    max_val2 = max(agrup_m2['HH_STD_TOTAL'].max(), agrup_m2['HH_Productivas_C/GAP'].max())
    aplicar_anti_overlap(ax1, max_val2, multiplier=2.0)
    
    ax1.bar_label(bars_std2, padding=4, color='black', fontweight='bold', fontsize=12, path_effects=outline_white, fmt='%.0f', zorder=3)
    ax1.bar_label(bars_prod2, padding=4, color='black', fontweight='bold', fontsize=12, path_effects=outline_white, fmt='%.0f', zorder=3)

    dibujar_meses(ax1, x_indexes)

    for i, bar in enumerate(bars_std2):
        cant = int(agrup_m2['Cant._Prod._A1'].iloc[i])
        if cant > 0:
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() * 0.05, f"{cant} UND", 
                     rotation=90, color='white', ha='center', va='bottom', fontsize=16, fontweight='bold', path_effects=outline_black, zorder=4)

    ax2.plot(x_indexes, agrup_m2['Eficiencia_Prod'], color='dimgray', marker='o', markersize=10, linewidth=4, path_effects=outline_white, label='% Efic. Prod.', zorder=5)
    ax2.axhline(y=100, color='forestgreen', linestyle='--', linewidth=3, label='Meta (100%)', zorder=1)
    
    # Elevar límite secundario
    max_ef_prod = agrup_m2['Eficiencia_Prod'].max()
    ax2.set_ylim(0, max(150, max_ef_prod * 1.5))
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

    if len(x_indexes) > 1:
        z2 = np.polyfit(x_indexes, agrup_m2['Eficiencia_Prod'], 1)
        p2 = np.poly1d(z2)
        ax2.plot(x_indexes, p2(x_indexes), color='dimgray', linestyle=':', alpha=0.8, linewidth=2, zorder=1)

    offset_y2_m2 = ax2.get_ylim()[1] * 0.06
    for i, val in enumerate(agrup_m2['Eficiencia_Prod']):
        ax2.annotate(f"{val:.1f}%", (x_indexes[i], val + offset_y2_m2), color='white', bbox=bbox_gray, ha='center', fontsize=12, fontweight='bold', zorder=10)

    ax1.set_xticks(x_indexes)
    ax1.set_xticklabels(agrup_m2['Fecha_str'], fontsize=12, fontweight='bold')
    ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.15), ncol=2)
    ax2.legend(loc='upper right', bbox_to_anchor=(1, 1.15))
    st.pyplot(fig2)

st.markdown("---")

# =========================================================================
# MÉTRICA 3: GAP DE HH (EVALUACIÓN GLOBAL) - *Usando HH Productivas Puras*
# =========================================================================
st.header("3. GAP DE HH (EVALUACIÓN GLOBAL)")

if not df_ef_filtrado.empty:
    col_prod_pura = 'HH_Productivas' if 'HH_Productivas' in df_ef_filtrado.columns else 'HH Productivas'
    
    agrup_m3 = df_ef_filtrado.groupby('Fecha').agg({
        col_prod_pura: 'sum',
        'HH_Improductivas': 'sum',
        'HH_Disponibles': 'sum'
    }).reset_index()
    
    agrup_m3['Total_Declaradas'] = agrup_m3[col_prod_pura] + agrup_m3['HH_Improductivas']
    agrup_m3['Fecha_str'] = agrup_m3['Fecha'].dt.strftime('%b-%y')

    fig3, ax1 = plt.subplots(figsize=(14, 7))
    x_indexes = np.arange(len(agrup_m3))
    
    bar_prod = ax1.bar(x_indexes, agrup_m3[col_prod_pura], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
    bar_imp = ax1.bar(x_indexes, agrup_m3['HH_Improductivas'], bottom=agrup_m3[col_prod_pura], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
    
    labels_prod = [f'{int(val)}' if val / tot > 0.05 else '' for val, tot in zip(agrup_m3[col_prod_pura], agrup_m3['Total_Declaradas'])]
    labels_imp = [f'{int(val)}' if val / tot > 0.05 else '' for val, tot in zip(agrup_m3['HH_Improductivas'], agrup_m3['Total_Declaradas'])]
    
    ax1.bar_label(bar_prod, labels=labels_prod, label_type='center', color='white', fontweight='bold', fontsize=14, path_effects=outline_black, zorder=4)
    ax1.bar_label(bar_imp, labels=labels_imp, label_type='center', color='white', fontweight='bold', fontsize=14, path_effects=outline_black, zorder=4)

    ax1.plot(x_indexes, agrup_m3['HH_Disponibles'], color='black', marker='D', markersize=10, linewidth=4, path_effects=outline_white, label='HH DISPONIBLES', zorder=5)
    
    aplicar_anti_overlap(ax1, agrup_m3['HH_Disponibles'].max(), multiplier=1.60)
    dibujar_meses(ax1, x_indexes)

    for i in range(len(x_indexes)):
        disp = agrup_m3['HH_Disponibles'].iloc[i]
        decl = agrup_m3['Total_Declaradas'].iloc[i]
        gap = disp - decl
        
        ax1.plot([i, i], [decl, disp], color='dimgray', linewidth=5, alpha=0.6, linestyle='-', zorder=3)
        
        # Desfasaje matemático seguro
        offset_y_gap = decl + (gap / 2) if gap > 0 else decl + (ax1.get_ylim()[1] * 0.05)
        ax1.annotate(f"GAP HH Ocultas:\n{int(gap)}", (i, offset_y_gap), color='firebrick', bbox=bbox_white, ha='center', va='center', fontsize=12, fontweight='bold', zorder=10)
        
        offset_y_disp = disp + (ax1.get_ylim()[1] * 0.08)
        ax1.annotate(f"{int(disp)}", (i, offset_y_disp), color='black', bbox=bbox_white, ha='center', fontsize=12, fontweight='bold', zorder=10)

    ax1.set_xticks(x_indexes)
    ax1.set_xticklabels(agrup_m3['Fecha_str'], fontsize=12, fontweight='bold')
    ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.15), ncol=3)
    st.pyplot(fig3)

st.markdown("---")

# =========================================================================
# MÉTRICA 4: COSTOS H.H. IMPRODUCTIVAS DECLARADAS
# =========================================================================
st.header("4. COSTOS H.H. IMPRODUCTIVAS DECLARADAS")

if not df_ef_filtrado.empty:
    agrup_m4 = df_ef_filtrado.groupby('Fecha').agg({
        'HH_Improductivas': 'sum',
        'Costo_Improd._$': 'sum'
    }).reset_index()
    agrup_m4['Fecha_str'] = agrup_m4['Fecha'].dt.strftime('%b-%y')

    fig4, ax1 = plt.subplots(figsize=(14, 7))
    ax2 = ax1.twinx()
    
    bars_imp = ax1.bar(x_indexes, agrup_m4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
    ax1.bar_label(bars_imp, padding=4, color='black', fontweight='bold', fontsize=12, path_effects=outline_white, zorder=4)
    
    # Barras en la mitad inferior
    aplicar_anti_overlap(ax1, agrup_m4['HH_Improductivas'].max(), multiplier=2.0)
    
    ax2.plot(x_indexes, agrup_m4['Costo_Improd._$'], color='maroon', marker='s', markersize=10, linewidth=5, path_effects=outline_white, label='COSTO ARS', zorder=5)
    
    # Eje secundario elevado
    max_costo = agrup_m4['Costo_Improd._$'].max()
    ax2.set_ylim(0, max(max_costo * 1.50, 1000))
    
    ticks_y = ax2.get_yticks()
    ax2.set_yticklabels([f'${int(x/1000000)}M' for x in ticks_y], fontweight='bold')

    costo_total = agrup_m4['Costo_Improd._$'].sum()
    ax1.text(len(x_indexes)/2 - 0.5, ax1.get_ylim()[1]*0.90, f"COSTO TOTAL ACUMULADO ARS\n${costo_total:,.0f}", 
             ha='center', va='center', fontsize=18, color='black', bbox=bbox_yellow, weight='bold', zorder=10)

    offset_y2_m4 = ax2.get_ylim()[1] * 0.08
    for i, val in enumerate(agrup_m4['Costo_Improd._$']):
        ax2.annotate(f"${val:,.0f}", (x_indexes[i], val + offset_y2_m4), color='white', bbox=bbox_gray, ha='center', fontsize=12, fontweight='bold', zorder=10)

    ax1.set_xticks(x_indexes)
    ax1.set_xticklabels(agrup_m4['Fecha_str'], fontsize=12, fontweight='bold')
    ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.15))
    ax2.legend(loc='upper right', bbox_to_anchor=(1, 1.15))
    st.pyplot(fig4)

st.markdown("---")

# =========================================================================
# MÉTRICA 5: DIAGRAMA DE PARETO GLOBAL (ÚLTIMO TRIMESTRE)
# =========================================================================
st.header("5. DIAGRAMA DE PARETO GLOBAL (ÚLTIMO TRIMESTRE DISPONIBLE)")

if not df_imp_filtrado.empty:
    max_date = df_imp_filtrado['FECHA'].max()
    if pd.notnull(max_date):
        trimestre = df_imp_filtrado[df_imp_filtrado['FECHA'] >= (max_date - pd.DateOffset(months=3))]
        
        pareto_df = trimestre.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        meses_unicos = trimestre['FECHA'].nunique()
        divisor_promedio = meses_unicos if meses_unicos > 0 else 1
        
        pareto_df['Promedio_Mensual'] = pareto_df['HH_IMPRODUCTIVAS'] / divisor_promedio
        pareto_df = pareto_df.sort_values(by='Promedio_Mensual', ascending=False)
        pareto_df['%_Acumulado'] = (pareto_df['Promedio_Mensual'].cumsum() / pareto_df['Promedio_Mensual'].sum()) * 100

        fig5, ax1 = plt.subplots(figsize=(16, 8))
        ax2 = ax1.twinx()
        
        x_pos = np.arange(len(pareto_df))
        bars_pareto = ax1.bar(x_pos, pareto_df['Promedio_Mensual'], color='maroon', edgecolor='white', zorder=2)
        
        # Incrementar límite para alejar las barras del techo
        aplicar_anti_overlap(ax1, pareto_df['Promedio_Mensual'].max(), multiplier=1.60)
        ax1.bar_label(bars_pareto, padding=4, color='black', fontweight='bold', fontsize=11, fmt='%.1f', zorder=4)
        
        ax2.plot(x_pos, pareto_df['%_Acumulado'], color='red', marker='D', markersize=8, linewidth=4, path_effects=outline_white, zorder=5)
        ax2.axhline(y=80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        # Limite elevado para que la curva roja viva sin tocar el techo (130%)
        ax2.set_ylim(0, 130)
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

        labels_wrapped = [textwrap.fill(str(l), 15) for l in pareto_df['TIPO_PARADA']]
        ax1.set_xticks(x_pos)
        ax1.set_xticklabels(labels_wrapped, rotation=90, fontsize=11, fontweight='bold')
        
        # Etiquetas de porcentajes: Rotadas a 45 grados y con tamaño reducido para NO chocar entre sí en la zona plana
        offset_y2_m5 = ax2.get_ylim()[1] * 0.04
        for i, val in enumerate(pareto_df['%_Acumulado']):
            ax2.annotate(f"{val:.1f}%", (x_pos[i], val + offset_y2_m5), color='white', bbox=bbox_gray, 
                         ha='center', va='bottom', fontsize=9, fontweight='bold', rotation=45, zorder=10)

        # MOVIDOS: Carteles a la zona inferior derecha del Pareto (Siempre vacía)
        suma_promedio = pareto_df['Promedio_Mensual'].sum()
        ax1.text(len(x_pos)*0.75, ax1.get_ylim()[1]*0.35, f"SUMA PROMEDIO MENSUAL\n{suma_promedio:.1f} HH", 
                 bbox=bbox_gray, color='white', fontsize=14, fontweight='bold', ha='center', zorder=10)
        
        top5 = pareto_df.head(5)['TIPO_PARADA'].tolist()
        top5_str = "TOP 5 Causas:\n" + "\n".join([f"- {c}" for c in top5])
        ax1.text(len(x_pos)*0.75, ax1.get_ylim()[1]*0.10, top5_str, 
                 bbox=bbox_yellow, color='black', fontsize=12, fontweight='bold', ha='center', va='bottom', zorder=10)

        st.pyplot(fig5)
    else:
        st.warning("No hay fechas válidas en la base de horas improductivas.")

st.markdown("---")

# =========================================================================
# MÉTRICA 6: EVOLUCIÓN % H.H. IMPRODUCTIVAS VS DISPONIBLES
# =========================================================================
st.header("6. EVOLUCIÓN % H.H. IMPRODUCTIVAS VS DISPONIBLES")

if not df_imp_filtrado.empty:
    pivot_imp = pd.pivot_table(df_imp_filtrado, values='HH_IMPRODUCTIVAS', index='FECHA', columns='TIPO_PARADA', aggfunc='sum').fillna(0)
    disp_por_mes = df_ef_filtrado.groupby('Fecha')['HH_Disponibles'].sum()
    
    df_m6 = pivot_imp.join(disp_por_mes).fillna(0)
    df_m6['Total_Imp'] = pivot_imp.sum(axis=1)
    df_m6['Incidencia_%'] = (df_m6['Total_Imp'] / df_m6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    df_m6 = df_m6.sort_index()

    fechas_str = [d.strftime('%b-%y') for d in df_m6.index]
    x_m6 = np.arange(len(df_m6))
    
    fig6, ax1 = plt.subplots(figsize=(16, 8))
    ax2 = ax1.twinx()
    
    bottoms = np.zeros(len(df_m6))
    colors = plt.cm.tab20.colors
    
    for idx, col in enumerate(pivot_imp.columns):
        values = df_m6[col].values
        container = ax1.bar(x_m6, values, bottom=bottoms, label=col, color=colors[idx % len(colors)], edgecolor='white', zorder=2)
        
        labels_seg = [f'{int(val)}' if tot > 0 and (val/tot) > 0.05 else '' for val, tot in zip(values, df_m6['Total_Imp'])]
        ax1.bar_label(container, labels=labels_seg, label_type='center', color='black', fontweight='bold', path_effects=outline_white, fontsize=12, zorder=4)
        bottoms += values

    # Bajar las barras
    aplicar_anti_overlap(ax1, df_m6['Total_Imp'].max(), multiplier=2.0)
    
    for i in range(len(x_m6)):
        imp_val = df_m6['Total_Imp'].iloc[i]
        disp_val = df_m6['HH_Disponibles'].iloc[i]
        if imp_val > 0:
            offset_y_imp = imp_val + (ax1.get_ylim()[1] * 0.05)
            ax1.annotate(f"Imp: {int(imp_val)}\nDisp: {int(disp_val)}", 
                         (i, offset_y_imp), ha='center', bbox=bbox_yellow, fontsize=11, fontweight='bold', zorder=10)

    ax2.plot(x_m6, df_m6['Incidencia_%'], color='red', marker='o', markersize=9, linewidth=4, path_effects=outline_white, label='% Incidencia', zorder=5)
    lim_secundario = df_m6['Incidencia_%'].max() * 1.5 if df_m6['Incidencia_%'].max() > 0 else 100
    ax2.set_ylim(0, lim_secundario)
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
    
    if len(x_m6) > 1:
        z6 = np.polyfit(x_m6, df_m6['Incidencia_%'], 1)
        p6 = np.poly1d(z6)
        ax2.plot(x_m6, p6(x_m6), color='darkred', linestyle='--', linewidth=3, zorder=1)

    offset_y2_m6 = lim_secundario * 0.08
    for i, val in enumerate(df_m6['Incidencia_%']):
        ax2.annotate(f"{val:.1f}%", (x_m6[i], val + offset_y2_m6), color='red', ha='center', fontsize=14, fontweight='bold', path_effects=outline_white, zorder=10)

    ax1.text(len(x_m6)/2 - 0.5, ax1.get_ylim()[1]*0.92, 
             f"PROMEDIO INCIDENCIA: {df_m6['Incidencia_%'].mean():.1f}%\nTotal HH Imp: {df_m6['Total_Imp'].sum():.0f}", 
             bbox=bbox_gray, color='white', ha='center', fontsize=14, fontweight='bold', zorder=10)

    ax1.set_xticks(x_m6)
    ax1.set_xticklabels(fechas_str, fontsize=12, fontweight='bold')
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=5, fontsize=10)
    
    st.pyplot(fig6)

else:
    st.info("👋 Por favor, sube los archivos de 'Datos Finales Eficiencias' y 'Horas Improductivas Limpias' en el panel izquierdo para generar el C.G.P. Reporte Integrado.")
