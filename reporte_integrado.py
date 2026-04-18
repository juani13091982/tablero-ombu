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
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide")

# ==========================================
# 🚨 ACCESO GENERAL ÚNICO 
# ==========================================
USUARIOS_PERMITIDOS = {
    "acceso.ombu": "Gestion2026"
}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        # INTENTA CARGAR EL LOGO. SI NO LO ENCUENTRA, PONE EL CUADRO AZUL DE RESPALDO.
        try:
            st.image("LOGO OMBÚ.jpg", use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
        except Exception:
            st.markdown("""
                <div style='background-color:#1E3A8A; color:white; padding:20px; border-radius:10px 10px 0px 0px; text-align:center;'>
                    <h2 style='margin:0; font-weight:bold; letter-spacing: 2px;'>OMBÚ</h2>
                    <h5 style='margin:0; color:#cbd5e1;'>Acceso al Tablero de Gestión</h5>
                </div>
            """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("<h4 style='text-align: center; color: #333;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submit = st.form_submit_button("Ingresar", use_container_width=True)

            if submit:
                if usuario in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[usuario] == password:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos enviados por Control de Gestión.")

if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# ==========================================
# DISEÑO VISUAL Y ESTILOS (STICKY PANEL)
# ==========================================
css_styles = "<style>\n"

if "admin" not in st.query_params:
    css_styles += """
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    """
else:
    css_styles += """
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        top: 55px !important; 
    }
    """

css_styles += """
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

plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold'
})

outline_white = [path_effects.withStroke(linewidth=3, foreground='white')]
outline_black = [path_effects.withStroke(linewidth=3, foreground='black')]
bbox_gray = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_yellow = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
bbox_green = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
bbox_white = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# ==========================================
# FUNCIONES MATEMÁTICAS Y CRUCE DE DATOS
# ==========================================
def aplicar_anti_overlap(ax, max_val, multiplier=2.6):
    if max_val > 0: 
        ax.set_ylim(0, max_val * multiplier)
    else: 
        ax.set_ylim(0, 100)

def dibujar_meses(ax, fechas):
    for x in range(len(fechas)):
        ax.axvline(x=x, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def formatear_seleccion(lista_sel, default_str):
    if not lista_sel: 
        return default_str
    if len(lista_sel) > 2: 
        return f"Varios ({len(lista_sel)})"
    return " + ".join(lista_sel)

def clean_match(text):
    if pd.isna(text): 
        return ""
    t = str(text).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    return re.sub(r'[^A-Z0-9]', '', t)

def robust_match(val_sel, val_imp):
    """Motor Fuzzy Blindado e Inteligente: Encuentra raíces de palabras sin mezclar puestos."""
    if pd.isna(val_imp) or pd.isna(val_sel): 
        return False
    
    s1 = str(val_sel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(val_imp).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    c1 = re.sub(r'[^A-Z0-9]', '', s1)
    c2 = re.sub(r'[^A-Z0-9]', '', s2)
    
    if not c1 or not c2: 
        return False
        
    if c1 in c2 or c2 in c1: 
        return True
    
    n1 = set(re.findall(r'\d{3,}', s1))
    n2 = set(re.findall(r'\d{3,}', s2))
    
    if n1 and n2 and n1.intersection(n2): 
        return True
        
    w1 = set(re.findall(r'[A-Z]{4,}', s1))
    w2 = set(re.findall(r'[A-Z]{4,}', s2))
    
    exclusion = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'TOLVAS', 'BATEAS', 'REMOLQUES', 'MAQUINA'}
    
    valid_w1 = w1 - exclusion
    valid_w2 = w2 - exclusion
    
    for word1 in valid_w1:
        for word2 in valid_w2:
            if word1 in word2 or word2 in word1: 
                return True
                
    return False

# ==========================================
# HEADER Y LOGOUT
# ==========================================
col_l, col_t, col_o = st.columns([1, 3, 1])

with col_l:
    try: 
        st.image("LOGO OMBÚ.jpg", use_container_width=True)
    except: 
        st.markdown("<div style='background-color:#1E3A8A; color:white; padding:10px; border-radius:5px; text-align:center;'>OMBÚ</div>", unsafe_allow_html=True)

with col_t:
    st.title("REPORTE INTEGRADO C.G.P.")

with col_o:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True):
        st.session_state['autenticado'] = False
        st.rerun()

# ==========================================
# CARGA Y FILTROS
# ==========================================
try:
    df_ef = pd.read_excel("eficiencias.xlsx")
    df_imp = pd.read_excel("improductivas.xlsx")
    
    df_ef.columns = df_ef.columns.str.strip()
    df_imp.columns = [str(c).strip().upper() for c in df_imp.columns]
    
    if 'TIPO_PARADA' not in df_imp.columns:
        col_t = next((c for c in df_imp.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if col_t: 
            df_imp.rename(columns={col_t: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_imp.columns:
        col_h = next((c for c in df_imp.columns if 'HH' in c and 'IMP' in c), None)
        if col_h: 
            df_imp.rename(columns={col_h: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_imp.columns:
        col_f = next((c for c in df_imp.columns if 'FECHA' in c), None)
        if col_f: 
            df_imp.rename(columns={col_f: 'FECHA'}, inplace=True)
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_imp['FECHA'] = pd.to_datetime(df_imp['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    
    df_ef['Mes_Filtro'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_imp['Mes_Filtro'] = df_imp['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e:
    st.error("Error cargando bases. Asegúrese de que eficiencias.xlsx e improductivas.xlsx estén en el servidor.")
    st.stop()

with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    f1, f2, f3, f4 = st.columns(4)
    
    with f1: 
        plantas_disponibles = list(df_ef['Planta'].dropna().unique())
        planta_sel = st.multiselect("🏭 Planta", plantas_disponibles, placeholder="Todas (Dejar vacío)")
        
    plantas_filtrar = planta_sel if planta_sel else plantas_disponibles
    df_temp_linea = df_ef[df_ef['Planta'].isin(plantas_filtrar)]
    
    with f2: 
        lineas_disponibles = list(df_temp_linea['Linea'].dropna().unique())
        linea_sel = st.multiselect("⚙️ Línea", lineas_disponibles, placeholder="Todas (Dejar vacío)")
        
    lineas_filtrar = linea_sel if linea_sel else lineas_disponibles
    df_temp_puesto = df_temp_linea[df_temp_linea['Linea'].isin(lineas_filtrar)]
    
    with f3: 
        puestos_disponibles = list(df_temp_puesto['Puesto_Trabajo'].dropna().unique())
        puesto_sel = st.multiselect("🛠️ Puesto de Trabajo", puestos_disponibles, placeholder="Todos (Dejar vacío)")
        
    puestos_filtrar = puesto_sel if puesto_sel else puestos_disponibles
    df_temp_mes = df_temp_puesto[df_temp_puesto['Puesto_Trabajo'].isin(puestos_filtrar)]
    
    with f4: 
        meses_disponibles = ["🎯 Acumulado YTD"] + list(df_temp_mes['Mes_Filtro'].dropna().unique())
        mes_sel = st.multiselect("📅 Mes", meses_disponibles, placeholder="Todos (Dejar vacío)")

txt_filtro_planta = formatear_seleccion(planta_sel, "Todas")
txt_filtro_linea = formatear_seleccion(linea_sel, "Todas")
txt_filtro_puesto = formatear_seleccion(puesto_sel, "Todos")
texto_filtros_header = f"PLANTA: {txt_filtro_planta} > LÍNEA: {txt_filtro_linea} > PUESTO DE TRABAJO: {txt_filtro_puesto}"

# ==========================================
# APLICACIÓN DE FILTROS ROBUSTOS Y YTD
# ==========================================
df_ef_f = df_ef.copy()
df_imp_f = df_imp.copy()

if planta_sel: 
    df_ef_f = df_ef_f[df_ef_f['Planta'].astype(str).str.strip().str.upper().isin([str(x).strip().upper() for x in planta_sel])]

if linea_sel: 
    df_ef_f = df_ef_f[df_ef_f['Linea'].astype(str).str.strip().str.upper().isin([str(x).strip().upper() for x in linea_sel])]

if puesto_sel: 
    df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].astype(str).str.strip().str.upper().isin([str(x).strip().upper() for x in puesto_sel])]

if mes_sel and "🎯 Acumulado YTD" not in mes_sel:
    df_ef_f = df_ef_f[df_ef_f['Mes_Filtro'].isin(mes_sel)]

col_planta_imp = next((c for c in df_imp_f.columns if str(c).strip().upper() in ['PLANTA', 'PLANTAS', 'ÁREA', 'AREA']), None)
col_linea_imp = next((c for c in df_imp_f.columns if str(c).strip().upper() in ['LÍNEA', 'LINEA', 'LINEAS']), None)
col_puesto_imp = next((c for c in df_imp_f.columns if str(c).strip().upper() in ['PUESTO', 'PUESTOS', 'PUESTO_TRABAJO', 'PUESTO DE TRABAJO']), None)

if planta_sel and col_planta_imp: 
    mask = df_imp_f[col_planta_imp].apply(lambda x: any(robust_match(s, x) for s in planta_sel))
    df_imp_f = df_imp_f[mask]

if linea_sel and col_linea_imp: 
    mask = df_imp_f[col_linea_imp].apply(lambda x: any(robust_match(s, x) for s in linea_sel))
    df_imp_f = df_imp_f[mask]

if puesto_sel and col_puesto_imp: 
    mask = df_imp_f[col_puesto_imp].apply(lambda x: any(robust_match(s, x) for s in puesto_sel))
    df_imp_f = df_imp_f[mask]

if mes_sel and "🎯 Acumulado YTD" not in mes_sel and 'Mes_Filtro' in df_imp_f.columns:
    df_imp_f = df_imp_f[df_imp_f['Mes_Filtro'].isin(mes_sel)]

st.markdown("---")

# =========================================================================
# FILA 1: MÉTRICAS 1 Y 2
# =========================================================================
c1, c2 = st.columns(2)

with c1:
    st.header("1. EFICIENCIA REAL")
    
    puestos_especificos = (len(puesto_sel) > 0)
    mostrar_alerta = False
    
    if puestos_especificos:
        st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Fórmula: (∑ HH STD / ∑ HH Disp.) de puestos seleccionados</i></div>", unsafe_allow_html=True)
        df_m1 = df_ef_f.copy()
    else:
        st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Fórmula: (∑ HH STD / ∑ HH Disp.) SOLO en última estación</i></div>", unsafe_allow_html=True)
        df_m1_si = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI'].copy()
        
        if df_m1_si.empty and not df_ef_f.empty:
            mostrar_alerta = True
            df_m1 = df_ef_f.copy()
        else:
            df_m1 = df_m1_si

    if not df_m1.empty:
        ag1 = df_m1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Disponibles': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        ag1['Ef'] = (ag1['HH_STD_TOTAL'] / ag1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig1, ax1 = plt.subplots(figsize=(14,10))
        ax2 = ax1.twinx()
        
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig1.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')
        
        x_idx = np.arange(len(ag1))
        width = 0.35
        
        b_std = ax1.bar(x_idx - width/2, ag1['HH_STD_TOTAL'], width, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_disp = ax1.bar(x_idx + width/2, ag1['HH_Disponibles'], width, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        aplicar_anti_overlap(ax1, ag1['HH_Disponibles'].max(), 2.6)
        
        ax1.bar_label(b_std, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        ax1.bar_label(b_disp, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        
        dibujar_meses(ax1, x_idx)

        for i, bar in enumerate(b_std):
            cant = int(ag1['Cant._Prod._A1'].iloc[i])
            if cant > 0: 
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{cant} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=outline_black, zorder=4)

        ax2.plot(x_idx, ag1['Ef'], color='dimgray', marker='o', markersize=10, linewidth=4, path_effects=outline_white, label='% Efic. Real', zorder=5)
        
        ax2.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2.text(x_idx[0], 85 + (ax2.get_ylim()[1]*0.01), 'META = 85%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', ha='center', va='bottom', zorder=10)
        
        ax2.set_ylim(0, max(120, ag1['Ef'].max()*1.8))
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

        if len(x_idx) > 1:
            z = np.polyfit(x_idx, ag1['Ef'], 1)
            p = np.poly1d(z)
            ax2.plot(x_idx, p(x_idx), color='dimgray', linestyle=':', alpha=0.8, linewidth=2, zorder=1)

        for i, val in enumerate(ag1['Ef']):
            ax2.annotate(f"{val:.1f}%", (x_idx[i], val + ax2.get_ylim()[1]*0.04), color='white', bbox=bbox_gray, ha='center', zorder=10)

        ax1.set_xticks(x_idx)
        ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'))
        
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax2.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig1)
        
        if mostrar_alerta: 
            st.markdown("<div style='height: 40px; color: #e74c3c; font-weight: bold;'>⚠️ La línea no posee salida ('SI'). Elija un puesto.</div>", unsafe_allow_html=True)
        else: 
            st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
    else: 
        st.warning("⚠️ No hay datos evaluables.")

with c2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_m1.empty:
        ag2 = df_m1.groupby('Fecha').agg({
            'HH_STD_TOTAL': 'sum', 
            'HH_Productivas_C/GAP': 'sum', 
            'Cant._Prod._A1': 'sum'
        }).reset_index()
        
        ag2['Ef'] = (ag2['HH_STD_TOTAL'] / ag2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax1 = plt.subplots(figsize=(14,10))
        ax2 = ax1.twinx()
        
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig2.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')
        
        x_idx = np.arange(len(ag2))
        width = 0.35
        
        b_std2 = ax1.bar(x_idx - width/2, ag2['HH_STD_TOTAL'], width, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_prod2 = ax1.bar(x_idx + width/2, ag2['HH_Productivas_C/GAP'], width, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        aplicar_anti_overlap(ax1, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 2.6)
        
        ax1.bar_label(b_std2, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        ax1.bar_label(b_prod2, padding=4, color='black', fontweight='bold', path_effects=outline_white, fmt='%.0f', zorder=3)
        
        dibujar_meses(ax1, x_idx)

        for i, bar in enumerate(b_std2):
            cant = int(ag2['Cant._Prod._A1'].iloc[i])
            if cant > 0: 
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{cant} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=outline_black, zorder=4)

        ax2.plot(x_idx, ag2['Ef'], color='dimgray', marker='o', markersize=10, linewidth=4, path_effects=outline_white, label='% Efic. Prod.', zorder=5)
        
        ax2.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2.text(x_idx[0], 100 + (ax2.get_ylim()[1]*0.01), 'META = 100%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', ha='center', va='bottom', zorder=10)
        
        ax2.set_ylim(0, max(150, ag2['Ef'].max()*1.8))
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

        if len(x_idx) > 1:
            z2 = np.polyfit(x_idx, ag2['Ef'], 1)
            p2 = np.poly1d(z2)
            ax2.plot(x_idx, p2(x_idx), color='dimgray', linestyle=':', alpha=0.8, linewidth=2, zorder=1)

        for i, val in enumerate(ag2['Ef']):
            ax2.annotate(f"{val:.1f}%", (x_idx[i], val + ax2.get_ylim()[1]*0.04), color='white', bbox=bbox_gray, ha='center', zorder=10)

        ax1.set_xticks(x_idx)
        ax1.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y'))
        
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax2.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig2)
    else: 
        st.warning("⚠️ No hay datos evaluables.")

st.markdown("---")

# =========================================================================
# FILA 2: MÉTRICAS 3 Y 4
# =========================================================================
c3, c4 = st.columns(2)

with c3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Diferencia entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        col_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        
        ag3 = df_ef_f.groupby('Fecha').agg({
            col_prod: 'sum', 
            'HH_Improductivas': 'sum', 
            'HH_Disponibles': 'sum'
        }).reset_index()
        
        ag3['Total_Decl'] = ag3[col_prod] + ag3['HH_Improductivas']
        
        fig3, ax1 = plt.subplots(figsize=(14, 10))
        
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')
        
        x_idx = np.arange(len(ag3))
        
        b_prod = ax1.bar(x_idx, ag3[col_prod], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        b_imp = ax1.bar(x_idx, ag3['HH_Improductivas'], bottom=ag3[col_prod], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        
        l_prod = [f'{int(v)}' if v/t > 0.05 else '' for v,t in zip(ag3[col_prod], ag3['Total_Decl'])]
        l_imp = [f'{int(v)}' if v/t > 0.05 else '' for v,t in zip(ag3['HH_Improductivas'], ag3['Total_Decl'])]
        
        ax1.bar_label(b_prod, labels=l_prod, label_type='center', color='white', fontweight='bold', path_effects=outline_black, zorder=4)
        ax1.bar_label(b_imp, labels=l_imp, label_type='center', color='white', fontweight='bold', path_effects=outline_black, zorder=4)

        ax1.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=10, linewidth=4, path_effects=outline_white, label='HH DISPONIBLES', zorder=5)
        
        aplicar_anti_overlap(ax1, ag3['HH_Disponibles'].max(), 2.6)
        dibujar_meses(ax1, x_idx)

        for i in range(len(x_idx)):
            disp = ag3['HH_Disponibles'].iloc[i]
            decl = ag3['Total_Decl'].iloc[i]
            gap = disp - decl
            
            ax1.plot([i, i], [decl, disp], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            
            offset_y_gap = decl + (gap / 2) if gap > 0 else decl + (ax1.get_ylim()[1]*0.05)
            ax1.annotate(f"GAP:\n{int(gap)}", (i, offset_y_gap), color='firebrick', bbox=bbox_white, ha='center', va='center', zorder=10)
            
            offset_y_disp = disp + (ax1.get_ylim()[1]*0.08)
            ax1.annotate(f"{int(disp)}", (i, offset_y_disp), color='black', bbox=bbox_white, ha='center', zorder=10)

        ax1.set_xticks(x_idx)
        ax1.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y'))
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        
        st.pyplot(fig3)
    else: 
        st.warning("⚠️ No hay datos evaluables.")

with c4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Valorización económica del impacto</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({
            'HH_Improductivas': 'sum', 
            'Costo_Improd._$': 'sum'
        }).reset_index()
        
        fig4, ax1 = plt.subplots(figsize=(14, 10))
        ax2 = ax1.twinx()
        
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')

        x_idx = np.arange(len(ag4))
        
        b_imp = ax1.bar(x_idx, ag4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax1.bar_label(b_imp, padding=4, color='black', fontweight='bold', path_effects=outline_white, zorder=4)
        
        aplicar_anti_overlap(ax1, ag4['HH_Improductivas'].max(), 2.6)
        
        ax2.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=10, linewidth=5, path_effects=outline_white, label='COSTO ARS', zorder=5)
        
        max_costo = ag4['Costo_Improd._$'].max()
        ax2.set_ylim(0, max(1000, max_costo * 1.8))
        
        ticks_y = ax2.get_yticks()
        ax2.set_yticklabels([f'${int(x/1000000)}M' for x in ticks_y], fontweight='bold')

        costo_total = ag4['Costo_Improd._$'].sum()
        hh_imp_total = ag4['HH_Improductivas'].sum()
        
        texto_cartel = f"COSTO TOTAL ACUMULADO ARS\n${costo_total:,.0f}\n(Total: {hh_imp_total:,.0f} HH Imp.)"
        ax1.text(0.5, 0.90, texto_cartel, transform=ax1.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=bbox_yellow, weight='bold', zorder=10)

        for i, val in enumerate(ag4['Costo_Improd._$']):
            ax2.annotate(f"${val:,.0f}", (x_idx[i], val + ax2.get_ylim()[1]*0.05), color='white', bbox=bbox_gray, ha='center', zorder=10)

        ax1.set_xticks(x_idx)
        ax1.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'))
        
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True)
        ax2.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        
        st.pyplot(fig4)
    else: 
        st.warning("⚠️ No hay datos evaluables.")

st.markdown("---")

# =========================================================================
# FILA 3: MÉTRICAS 5 Y 6 - ANÁLISIS DE CAUSA RAÍZ
# =========================================================================
c5, c6 = st.columns(2)

with c5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Distribución 80/20 de los motivos de Improductividad</i></div>", unsafe_allow_html=True)

    if not df_imp_f.empty:
        pareto = df_imp_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        
        meses_unicos = df_imp_f['FECHA'].nunique()
        divisor = meses_unicos if meses_unicos > 0 else 1
        
        pareto['Prom_Mensual'] = pareto['HH_IMPRODUCTIVAS'] / divisor
        pareto = pareto.sort_values(by='Prom_Mensual', ascending=False)
        pareto['%_Acumulado'] = (pareto['Prom_Mensual'].cumsum() / pareto['Prom_Mensual'].sum()) * 100

        fig5, ax1 = plt.subplots(figsize=(14, 10))
        ax2 = ax1.twinx()
        
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        fig5.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')

        x_pos = np.arange(len(pareto))
        
        b_par = ax1.bar(x_pos, pareto['Prom_Mensual'], color='maroon', edgecolor='white', zorder=2)
        aplicar_anti_overlap(ax1, pareto['Prom_Mensual'].max(), 2.8)
        ax1.bar_label(b_par, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        ax2.plot(x_pos, pareto['%_Acumulado'], color='red', marker='D', markersize=8, linewidth=4, path_effects=outline_white, zorder=5)
        ax2.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        
        ax2.set_ylim(0, 200)
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

        l_wrap = [textwrap.fill(str(l), 12) for l in pareto['TIPO_PARADA']]
        ax1.set_xticks(x_pos)
        ax1.set_xticklabels(l_wrap, rotation=90, fontsize=12)
        
        for i, val in enumerate(pareto['%_Acumulado']):
            ax2.annotate(f"{val:.1f}%", (x_pos[i], val + 4), color='white', bbox=bbox_gray, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)

        suma_promedio = pareto['Prom_Mensual'].sum()
        ax1.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{suma_promedio:.1f} HH", transform=ax1.transAxes, bbox=bbox_gray, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        top5_str = "TOP 5 Causas:\n" + "\n".join([f"- {c}" for c in pareto.head(5)['TIPO_PARADA']])
        ax1.text(0.02, 0.82, top5_str, transform=ax1.transAxes, bbox=bbox_yellow, color='black', fontsize=13, ha='left', va='top', zorder=10)
        
        st.pyplot(fig5)
        
        # ==========================================
        # MESA DE TRABAJO INTERACTIVA Y DESCARGA
        # ==========================================
        st.markdown("### 🛠️ Mesa de Trabajo: Análisis de Causa Raíz")
        st.markdown("<div style='font-size: 14px; color: #a0a0a0; margin-top:-10px; margin-bottom:10px;'><i>Selecciona el motivo del Pareto para auditar detalles y estandarizar acciones.</i></div>", unsafe_allow_html=True)
        
        motivos_disp = ["Todos"] + pareto['TIPO_PARADA'].tolist()
        motivo_sel = st.selectbox("🎯 Filtrar Motivo Específico:", motivos_disp)
        
        if motivo_sel:
            col_sub = next((c for c in df_imp_f.columns if 'SUB' in str(c).upper() or 'DETALLE' in str(c).upper()), None)
            
            if col_sub:
                if motivo_sel == "Todos": 
                    df_foco = df_imp_f.copy() 
                else: 
                    df_foco = df_imp_f[df_imp_f['TIPO_PARADA'] == motivo_sel]
                
                if not df_foco.empty:
                    df_sub = df_foco.groupby(col_sub)['HH_IMPRODUCTIVAS'].sum().reset_index()
                    df_top = df_sub.sort_values(by='HH_IMPRODUCTIVAS', ascending=False).copy()
                    
                    def clasificar_sub_motivo(txt):
                        t = str(txt).upper()
                        if any(kw in t for kw in ['RETOQUE', 'PINTURA', 'GOTEO', 'CHORREADURA', 'ADHERENCIA']): 
                            return "Retoque de Pintura"
                        if any(kw in t for kw in ['SOLDADURA', 'REPASO', 'PORO', 'FISURA', 'ESCORIA']): 
                            return "Defecto de Soldadura"
                        if any(kw in t for kw in ['PLEGADA', 'CONJUNTO', 'DIMENSIÓN', 'MEDIDA', 'AJUSTE', 'FUERA DE ESCUADRA']): 
                            return "Desviación Dimensional / Armado"
                        if any(kw in t for kw in ['ESPERANDO', 'FALTA DE MATERIAL', 'ABASTECIMIENTO', 'LOGÍSTICA', 'PUENTE']): 
                            return "Retraso Logístico / Abastecimiento"
                        if any(kw in t for kw in ['ROTURA', 'MANTENIMIENTO', 'ELÉCTRICA', 'MECÁNICA', 'MÁQUINA']): 
                            return "Falla de Equipo / Rotura"
                        return "Otros / Requiere Análisis"

                    df_top["Estandarización Sugerida"] = df_top[col_sub].apply(clasificar_sub_motivo)
                    
                    # TABLA DE RESUMEN CON TOTAL
                    t_hh = df_top['HH_IMPRODUCTIVAS'].sum()
                    
                    df_res = df_top.groupby('Estandarización Sugerida')['HH_IMPRODUCTIVAS'].sum().reset_index()
                    df_res = df_res.sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
                    df_res['%'] = (df_res['HH_IMPRODUCTIVAS'] / t_hh) * 100
                    
                    fila_t = pd.DataFrame({
                        'Estandarización Sugerida': ['✅ TOTAL'], 
                        'HH_IMPRODUCTIVAS': [t_hh], 
                        '%': [100.0]
                    })
                    
                    df_res = pd.concat([df_res, fila_t], ignore_index=True)
                    
                    st.markdown(f"<div style='color: #1E3A8A; font-weight: bold; margin-top: 15px;'>📊 Resumen de Impacto: {motivo_sel}</div>", unsafe_allow_html=True)
                    
                    st.dataframe(
                        df_res.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH'}), 
                        hide_index=True, 
                        use_container_width=True, 
                        column_config={
                            "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), 
                            "%": st.column_config.NumberColumn(format="%.1f %%")
                        }
                    )
                    
                    # AUDITORÍA Y EDICIÓN
                    df_top["Plan de Acción"] = ""
                    df_top["Responsable"] = ""
                    df_top["Estado"] = "Pendiente"
                    
                    st.markdown("<i>Auditoría de textos reales. Doble clic en celdas vacías para editar.</i>", unsafe_allow_html=True)
                    
                    df_editado = st.data_editor(
                        df_top.rename(columns={col_sub: 'Detalle Real Cargado', 'HH_IMPRODUCTIVAS': 'HH Perdidas'}),
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "HH Perdidas": st.column_config.NumberColumn(format="%.1f ⏱️"), 
                            "Estado": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "En Análisis", "Resuelto"], required=True)
                        }
                    )
                    
                    # BOTÓN DE DESCARGA
                    st.markdown("<br>", unsafe_allow_html=True)
                    csv_data = df_editado.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Descargar Plan de Acción a Excel (CSV)", 
                        data=csv_data, 
                        file_name=f"Plan_Accion_{motivo_sel}.csv", 
                        mime="text/csv", 
                        use_container_width=True, 
                        type="primary"
                    )

                else: 
                    st.info("No hay registros de detalles para esta selección.")
            else: 
                st.warning("⚠️ No se encontró la columna de sub-motivos en la base.")
    else:
        st.success("✅ ¡Excelente! No se registraron Horas Improductivas para los filtros aplicados en este período.")

with c6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height: 25px; font-size: 15px; color: #a0a0a0;'><i>Porcentaje histórico de Horas Improductivas sobre las Horas Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        df_ef_f['Cruce'] = pd.to_datetime(df_ef_f['Fecha']).dt.strftime('%Y-%m')
        disp = df_ef_f.groupby('Cruce', as_index=False)['HH_Disponibles'].sum()

        if not df_imp_f.empty:
            df_imp_f['Cruce'] = pd.to_datetime(df_imp_f['FECHA']).dt.strftime('%Y-%m')
            piv = pd.pivot_table(df_imp_f, values='HH_IMPRODUCTIVAS', index='Cruce', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            m6 = pd.merge(disp, piv, on='Cruce', how='left').fillna(0)
            cols_par = [c for c in m6.columns if c not in ['HH_Disponibles', 'Cruce']]
        else:
            m6 = disp.copy()
            cols_par = []
            
        if cols_par:
            m6['Total_Imp'] = m6[cols_par].sum(axis=1)
        else:
            m6['Total_Imp'] = 0
            
        m6['%'] = (m6['Total_Imp'] / m6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        m6['FECHA_R'] = pd.to_datetime(m6['Cruce'] + '-01')
        m6.set_index('FECHA_R', inplace=True)
        m6 = m6.sort_index()

        fig6, ax1 = plt.subplots(figsize=(14, 10))
        ax2 = ax1.twinx()
        
        fig6.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92) 
        fig6.suptitle(texto_filtros_header, x=0.08, y=0.98, ha='left', fontsize=8, fontweight='bold', color='dimgray')

        x_m6 = np.arange(len(m6))
        bottoms = np.zeros(len(m6))
        colors = plt.cm.tab20.colors
        
        if cols_par:
            for idx, col in enumerate(cols_par):
                vals = m6[col].values
                b = ax1.bar(x_m6, vals, bottom=bottoms, label=col, color=colors[idx % len(colors)], edgecolor='white', zorder=2)
                
                l_seg = [f'{int(v)}' if t > 0 and (v/t) > 0.05 else '' for v, t in zip(vals, m6['Total_Imp'])]
                ax1.bar_label(b, labels=l_seg, label_type='center', color='black', fontweight='bold', path_effects=outline_white, zorder=4)
                bottoms += vals
        else:
            ax1.bar(x_m6, np.zeros(len(m6)), color='white')

        aplicar_anti_overlap(ax1, m6['Total_Imp'].max(), 2.2)
        
        for i in range(len(x_m6)):
            i_v = m6['Total_Imp'].iloc[i]
            d_v = m6['HH_Disponibles'].iloc[i]
            
            offset_y_imp = i_v + (ax1.get_ylim()[1]*0.05)
            ax1.annotate(f"Imp: {int(i_v)}\nDisp: {int(d_v)}", (i, offset_y_imp), ha='center', bbox=bbox_yellow, fontsize=13, fontweight='bold', zorder=10)

        ax2.plot(x_m6, m6['%'], color='red', marker='o', markersize=9, linewidth=4, path_effects=outline_white, label='% Incidencia', zorder=5)
        
        ax2.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2.text(x_m6[0], 15 + (ax2.get_ylim()[1]*0.01), 'META = 15%', color='white', bbox=bbox_green, fontsize=14, fontweight='bold', ha='center', va='bottom', zorder=10)
        
        ax2.set_ylim(0, max(40, m6['%'].max() * 1.8))
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        
        if len(x_m6) > 1:
            z6 = np.polyfit(x_m6, m6['%'], 1)
            p6 = np.poly1d(z6)
            ax2.plot(x_m6, p6(x_m6), color='darkred', linestyle='--', linewidth=3, zorder=1)

        for i, val in enumerate(m6['%']):
            ax2.annotate(f"{val:.1f}%", (x_m6[i], val + ax2.get_ylim()[1]*0.05), color='red', ha='center', fontsize=15, fontweight='bold', path_effects=outline_white, zorder=10)

        ax1.text(0.98, 0.95, f"PROMEDIO INCIDENCIA: {m6['%'].mean():.1f}%\nTotal HH Imp: {m6['Total_Imp'].sum():.0f}", transform=ax1.transAxes, bbox=bbox_gray, color='white', ha='right', va='top', fontsize=16, fontweight='bold', zorder=10)
        
        ax1.set_xticks(x_m6)
        ax1.set_xticklabels([d.strftime('%b-%y') for d in m6.index], fontsize=14, fontweight='bold')
        
        if cols_par: 
            ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=True, fontsize=10)
        else: 
            ax2.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), frameon=True, fontsize=10)
        
        st.pyplot(fig6)
    else:
        st.warning("⚠️ No hay datos evaluables de eficiencias para esta selección.")
