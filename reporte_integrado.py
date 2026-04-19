import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import matplotlib.image as mpimg
import textwrap
import re

# =========================================================================
# 1. CONFIGURACIÓN Y ESCUDO VISUAL TOTAL
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ocultar elementos de Streamlit Cloud y Cabecera */
    header, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="manage-app-button"], 
    #MainMenu, footer, .stAppDeployButton, .viewerBadge_container {display: none !important; visibility: hidden !important;}
    .block-container {padding-top: 1.5rem !important;}
    
    /* Contenedor del Panel de Control superior */
    #panel-control {
        background-color: #0E1117;
        padding-bottom: 20px;
        border-bottom: 3px solid #1E3A8A;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b = [pe.withStroke(linewidth=3, foreground='white')]
efecto_n = [pe.withStroke(linewidth=3, foreground='black')]

caja_v = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
caja_g = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
caja_b = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD (ACCESO RESTRINGIDO)
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        l1, l2, l3 = st.columns([1, 1, 1])
        with l2:
            try: st.image("LOGO OMBÚ.jpg", width=160)
            except: st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        st.markdown("<div style='text-align:center;'><h2 style='color:#1E3A8A;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</h2><p>Acceso Restringido - Control de Gestión</p></div>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in = st.text_input("Usuario Corporativo")
            p_in = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in:
                    st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")

if not st.session_state['autenticado']:
    mostrar_login(); st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE Y HERRAMIENTAS
# =========================================================================
def set_escala_y(ax, vmax, factor=1.4):
    ax.set_ylim(0, vmax * factor if vmax > 0 else 100)

def dibujar_meses(ax, n_meses):
    for i in range(n_meses):
        ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def cruce_robusto(sel, excel):
    if pd.isna(excel) or pd.isna(sel): return False
    s1, s2 = str(sel).upper(), str(excel).upper()
    for a,b in zip("ÁÉÍÓÚ", "AEIOU"): s1, s2 = s1.replace(a,b), s2.replace(a,b)
    l1, l2 = re.sub(r'[^A-Z0-9]', '', s1), re.sub(r'[^A-Z0-9]', '', s2)
    if not l1 or not l2: return False
    if l1 in l2 or l2 in l1: return True
    p1, p2 = set(re.findall(r'[A-Z0-9]{3,}', s1)), set(re.findall(r'[A-Z0-9]{3,}', s2))
    excl = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'LÍNEA', 'PLANTA', 'AREA', 'ÁREA', 'MAQUINA'}
    v1, v2 = p1 - excl, p2 - excl
    return v1.issubset(v2) or v2.issubset(v1) if v1 and v2 else False

def add_tendencia(ax, x, y):
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color='darkgray', linestyle=':', linewidth=4, path_effects=efecto_b, zorder=4, label='Tendencia')

def agregar_sello_agua(fig):
    try:
        img = mpimg.imread("LOGO OMBÚ.jpg")
        ax_logo = fig.add_axes([0.88, 0.02, 0.08, 0.08], zorder=1)
        ax_logo.imshow(img, alpha=0.35); ax_logo.axis('off')
    except: pass

def generar_accion_sugerida(detalle):
    d = str(detalle).upper()
    if '✅' in d: return "🎯 ACCIÓN GLOBAL"
    if any(x in d for x in ['ROTURA', 'FALLA', 'CORTE', 'MANTENIMIENTO', 'MECANICO', 'ELECTRICO', 'ROT']): return "⚙️ Revisar Equipo"
    if any(x in d for x in ['FALTA', 'MATERIAL', 'ESPERA', 'PUENTE', 'GRUA', 'ABASTECIMIENTO', 'INSUMO']): return "📦 Apurar Logística"
    if any(x in d for x in ['REPROCESO', 'CALIDAD', 'PLANO', 'ERROR', 'DEFECTO']): return "🔎 Ajustar Calidad"
    if any(x in d for x in ['LIMPIEZA', 'ORDEN', '5S']): return "🧹 Optimizar 5S"
    if any(x in d for x in ['REUNION', 'CHARLA', 'CAPACITACION', 'PERSONAL', 'AUSENCIA']): return "👥 Gestionar RRHH"
    if any(x in d for x in ['PREPARACION', 'SET UP', 'SETUP', 'CAMBIO']): return "⏱️ Reducir Set-Up"
    return "⚡ Investigar Causa"

# =========================================================================
# 4. CARGA Y EXTRACCIÓN DE DATOS
# =========================================================================
h_l, h_t, h_s = st.columns([1, 3, 1])
with h_l:
    try: st.image("LOGO OMBÚ.jpg", width=120)
    except: st.markdown("### OMBÚ")
with h_t:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
    st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</p>", unsafe_allow_html=True)
with h_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True): 
        st.session_state['autenticado'] = False; st.rerun()

try:
    df_ef = pd.read_excel("eficiencias.xlsx")
    df_im = pd.read_excel("improductivas.xlsx")
    df_ef.columns = df_ef.columns.str.strip()
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    if 'TIPO_PARADA' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), df_im.columns[0]): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), df_im.columns[0]): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'DETALLE' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c or 'DESC' in c), df_im.columns[0]): 'DETALLE'}, inplace=True)
    
    c_nom, c_ape = next((c for c in df_im.columns if 'NOMBRE' in c), None), next((c for c in df_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '') + ' ' + df_im[c_ape].astype(str).replace('nan', '')
    elif c_nom: df_im['OPERARIO'] = df_im[c_nom].astype(str).replace('nan', '')
    else: df_im['OPERARIO'] = "S/D"
    df_im['OPERARIO'] = df_im['OPERARIO'].str.strip().replace('', 'S/D')

    c_fec = next((c for c in df_im.columns if 'INICIO' in c or 'A3' in c or 'FECHA' in c), None)
    df_im['FECHA_EXACTA'] = pd.to_datetime(df_im[c_fec], errors='coerce') if c_fec else pd.NaT
    df_im['FECHA'] = df_im['FECHA_EXACTA'].dt.to_period('M').dt.to_timestamp()
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'], df_im['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y'), df_im['FECHA'].dt.strftime('%b-%Y')
except Exception as e:
    st.error(f"Error crítico cargando datos: {e}"); st.stop()

# =========================================================================
# 5. PANEL DE CONTROL: KPIs Y FILTROS LATERALES
# =========================================================================
st.markdown('<div id="panel-control">', unsafe_allow_html=True)

col_kpi, col_filtros = st.columns([3, 1], gap="large")

# --- LADO DERECHO: FILTROS VERTICALES ---
with col_filtros:
    st.markdown("<h4 style='margin-bottom:15px; color:#4B8BBE;'>🎛️ FILTROS MAESTROS</h4>", unsafe_allow_html=True)
    s_mes = st.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + list(df_ef['Mes_Str'].unique()))
    s_pl = st.multiselect("🏭 Planta", list(df_ef['Planta'].dropna().unique()))
    
    df_tl = df_ef[df_ef['Planta'].isin(s_pl)] if s_pl else df_ef
    s_li = st.multiselect("⚙️ Línea", list(df_tl['Linea'].dropna().unique()))
    
    df_tp = df_tl[df_tl['Linea'].isin(s_li)] if s_li else df_tl
    s_pu = st.multiselect("🛠️ Puesto", list(df_tp['Puesto_Trabajo'].dropna().unique()))

# --- PROCESAMIENTO DE DATOS SEGÚN FILTROS ---
df_ef_f, df_im_f = df_ef.copy(), df_im.copy()

if s_pl: df_ef_f = df_ef_f[df_ef_f['Planta'].isin(s_pl)]
if s_li: df_ef_f = df_ef_f[df_ef_f['Linea'].isin(s_li)]
if s_pu: df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(s_pu)]
if s_mes and "🎯 Acumulado YTD" not in s_mes: df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(s_mes)]

if not df_im_f.empty:
    if s_pl:
        col_pl = next((c for c in df_im_f.columns if 'PLANTA' in str(c).upper()), df_im_f.columns[0])
        df_im_f = df_im_f[df_im_f[col_pl].apply(lambda x: any(cruce_robusto(p, x) for p in s_pl))]
    if s_li:
        col_li = next((c for c in df_im_f.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), df_im_f.columns[1])
        df_im_f = df_im_f[df_im_f[col_li].apply(lambda x: any(cruce_robusto(l, x) for l in s_li))]
    if s_pu:
        col_pu = next((c for c in df_im_f.columns if 'PUESTO' in str(c).upper()), df_im_f.columns[2])
        df_im_f = df_im_f[df_im_f[col_pu].apply(lambda x: any(cruce_robusto(ps, x) for ps in s_pu))]
    if s_mes and "🎯 Acumulado YTD" not in s_mes: 
        df_im_f = df_im_f[df_im_f['Mes_Str'].isin(s_mes)]

# Lógica del Embudo para M1/M2 (Salida de Línea)
warn_linea = False
if s_pu: df_plot_1 = df_ef_f.copy()
elif s_li:
    df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_salida.empty: df_plot_1 = df_salida
    else: df_plot_1 = df_ef_f.copy(); warn_linea = True
else: df_plot_1 = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']

# --- CÁLCULOS PARA LOS CARTELES (KPIs) ---
tot_std = df_plot_1['HH_STD_TOTAL'].sum() if not df_plot_1.empty else 0
tot_disp = df_plot_1['HH_Disponibles'].sum() if not df_plot_1.empty else 0
tot_prod = df_plot_1['HH_Productivas_C/GAP'].sum() if ('HH_Productivas_C/GAP' in df_plot_1.columns and not df_plot_1.empty) else 0

kpi_ef_real = (tot_std / tot_disp * 100) if tot_disp > 0 else 0
kpi_ef_prod = (tot_std / tot_prod * 100) if tot_prod > 0 else 0

tot_costo = df_ef_f['Costo_Improd._$'].sum() if not df_ef_f.empty else 0
tot_hh_imp = df_ef_f['HH_Improductivas'].sum() if not df_ef_f.empty else 0

# Top 3 Puestos (Eficiencia Real) basado en df_ef_f para ver internas
top3_m1_html = "<div style='font-size:14px; color:#aaa; margin-top:10px;'>Sin datos suficientes</div>"
if not df_ef_f.empty:
    ag_puestos = df_ef_f.groupby('Puesto_Trabajo').agg({'HH_STD_TOTAL':'sum', 'HH_Disponibles':'sum'})
    ag_puestos['Ef'] = (ag_puestos['HH_STD_TOTAL'] / ag_puestos['HH_Disponibles'] * 100).fillna(0)
    top3_val = ag_puestos['Ef'].nlargest(3)
    if not top3_val.empty:
        filas_html = []
        for i, (p, v) in enumerate(top3_val.items(), 1):
            filas_html.append(f"<div style='display:flex; justify-content:space-between; margin-top:6px; padding:3px 0;'><span style='font-size:13px; font-weight:normal; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:140px;' title='{p}'>{i}. {p}</span><strong style='font-size:15px; color:#90CAF9;'>{v:.1f}%</strong></div>")
        top3_m1_html = "".join(filas_html)

# Top 3 Puestos (Improductivas)
top3_imp_html = "<div style='font-size:14px; color:#aaa; margin-top:10px;'>Sin datos suficientes</div>"
if not df_im_f.empty:
    c_pu_im = next((c for c in df_im_f.columns if 'PUESTO' in c), df_im_f.columns[2] if len(df_im_f.columns)>2 else None)
    if c_pu_im:
        ag_imp_puestos = df_im_f.groupby(c_pu_im)['HH_IMPRODUCTIVAS'].sum().nlargest(3)
        if not ag_imp_puestos.empty:
            filas_imp = []
            for i, (p, v) in enumerate(ag_imp_puestos.items(), 1):
                filas_imp.append(f"<div style='display:flex; justify-content:space-between; margin-top:6px; padding:3px 0;'><span style='font-size:13px; font-weight:normal; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:140px;' title='{p}'>{i}. {p}</span><strong style='font-size:15px; color:#FFCDD2;'>{v:.1f} HH</strong></div>")
            top3_imp_html = "".join(filas_imp)

# --- LADO IZQUIERDO: GRILLA HTML DE CARTELES ---
with col_kpi:
    html_grid = f"""
    <div style="display: grid; grid-template-columns: 1fr 1fr 1.3fr; grid-template-rows: 1fr 1.2fr; gap: 15px; height:100%;">
        
        <div style="background: linear-gradient(135deg, #e0e0e0, #f5f5f5); border: 2px solid #999; border-left: 8px solid #1E3A8A; padding: 15px; border-radius: 8px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); display:flex; flex-direction:column; justify-content:center; align-items:center;">
            <h4 style="margin:0; color: #1E3A8A; font-size:16px; text-transform:uppercase;">EFICIENCIA REAL</h4>
            <h2 style="margin:5px 0 0 0; color: #111; font-size:38px;">{kpi_ef_real:.1f}%</h2>
        </div>
        
        <div style="background: linear-gradient(135deg, #2E7D32, #4CAF50); border: 2px solid #1B5E20; border-left: 8px solid #A5D6A7; padding: 15px; border-radius: 8px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); display:flex; flex-direction:column; justify-content:center; align-items:center;">
            <h4 style="margin:0; color: white; font-size:16px; text-transform:uppercase;">EFICIENCIA PROD.</h4>
            <h2 style="margin:5px 0 0 0; color: white; font-size:38px;">{kpi_ef_prod:.1f}%</h2>
        </div>
        
        <div style="background: linear-gradient(135deg, #D32F2F, #E53935); border: 2px solid #B71C1C; padding: 20px; border-radius: 12px; grid-row: span 2; box-shadow: 2px 4px 15px rgba(211,47,47,0.4); display: flex; flex-direction: column; justify-content: center; text-align:center;">
            <h4 style="margin:0; color: white; font-size:22px; text-transform:uppercase;">COSTO HH IMPROD.</h4>
            <p style="margin:0; color: #FFCDD2; font-size:14px; font-weight:normal;">(Oportunidad Perdida)</p>
            <div style="background: rgba(255,255,255,0.1); padding: 15px; border-radius: 8px; margin-top: 15px;">
                <h2 style="margin:0; color: #FFEB3B; font-size:38px; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${tot_costo:,.0f}</h2>
            </div>
            <h4 style="margin:15px 0 0 0; color: white; font-size:22px;">{tot_hh_imp:,.0f} <span style="font-size:16px; font-weight:normal;">HH PERDIDAS</span></h4>
        </div>
        
        <div style="background: #0D47A1; color: white; padding: 15px; border-radius: 8px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); display:flex; flex-direction:column;">
            <h4 style="margin:0 0 10px 0; font-size:13px; color:#BBDEFB; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom:5px; text-align:center;">🏆 TOP 3 EF. REAL (PUESTOS)</h4>
            {top3_m1_html}
        </div>
        
        <div style="background: #B71C1C; color: white; padding: 15px; border-radius: 8px; box-shadow: 2px 4px 10px rgba(0,0,0,0.3); display:flex; flex-direction:column;">
            <h4 style="margin:0 0 10px 0; font-size:13px; color:#FFCDD2; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom:5px; text-align:center;">⚠️ TOP 3 MAYOR HH IMP.</h4>
            {top3_imp_html}
        </div>
        
    </div>
    """
    st.markdown(html_grid, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
t_enc = f"Filtros >> Planta: {'+'.join(s_pl) if s_pl else 'Todas'} | Línea: {'+'.join(s_li) if s_li else 'Todas'} | Puesto: {'+'.join(s_pu) if s_pu else 'Todos'}"

# =========================================================================
# 7. GRÁFICOS: MÉTRICAS 1 Y 2
# =========================================================================
col1, col2 = st.columns(2)

with col1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    if warn_linea: st.warning("⚠️ Esta Línea NO registra un 'Último Puesto'. Seleccione un Puesto para análisis preciso.")
    
    if not df_plot_1.empty:
        ag1 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'}).reset_index()
        ag1['Ef_Real'] = (ag1['HH_STD_TOTAL'] / ag1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig1, ax1 = plt.subplots(figsize=(14, 10)); ax1_line = ax1.twinx()
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig1.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag1))
        bs = ax1.bar(x_idx - 0.17, ag1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bd = ax1.bar(x_idx + 0.17, ag1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        set_escala_y(ax1, ag1['HH_Disponibles'].max(), 1.4) 
        ax1.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax1.bar_label(bd, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax1, len(x_idx))

        for i, bar in enumerate(bs):
            vu = int(ag1['Cant._Prod._A1'].iloc[i])
            if vu > 0: ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{vu} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)

        ax1_line.plot(x_idx, ag1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real', zorder=5)
        add_tendencia(ax1_line, x_idx, ag1['Ef_Real'])
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_idx[0], 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3)); ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag1['Ef_Real']):
            ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        agregar_sello_agua(fig1); st.pyplot(fig1)
    else: st.warning("⚠️ Sin datos para Eficiencia Real.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_plot_1.empty:
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Productivas_C/GAP': 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2 = plt.subplots(figsize=(14, 10)); ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag2))
        bs = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bp = ax2.bar(x_idx + 0.17, ag2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 1.4) 
        ax2.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax2.bar_label(bp, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax2, len(x_idx))

        ax2_line.plot(x_idx, ag2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Prod.', zorder=5)
        add_tendencia(ax2_line, x_idx, ag2['Ef_Prod'])
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_idx[0], 101, 'META = 100%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        ax2_line.set_ylim(0, max(110, ag2['Ef_Prod'].max()*1.3)); ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag2['Ef_Prod']):
            ax2_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax2.set_xticks(x_idx); ax2.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax2_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        agregar_sello_agua(fig2); st.pyplot(fig2)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")

# =========================================================================
# 8. GRÁFICOS: MÉTRICAS 3 Y 4
# =========================================================================
col3, col4 = st.columns(2)

with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        ag3 = df_ef_f.groupby('Fecha').agg({c_prod: 'sum', 'HH_Improductivas': 'sum', 'HH_Disponibles': 'sum'}).reset_index()
        ag3['Total_Decl'] = ag3[c_prod] + ag3['HH_Improductivas']
        
        fig3, ax3 = plt.subplots(figsize=(14, 10)); fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag3))
        
        bp = ax3.bar(x_idx, ag3[c_prod], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        bi = ax3.bar(x_idx, ag3['HH_Improductivas'], bottom=ag3[c_prod], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax3.bar_label(bp, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.bar_label(bi, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='HH DISPONIBLES', zorder=5)
        
        set_escala_y(ax3, ag3['HH_Disponibles'].max(), 1.4); dibujar_meses(ax3, len(x_idx))

        for i in range(len(x_idx)):
            hd, ht = ag3['HH_Disponibles'].iloc[i], ag3['Total_Decl'].iloc[i]; gap = hd - ht
            ax3.plot([i, i], [ht, hd], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            ax3.annotate(f"GAP:\n{int(gap)}", (i, ht + (gap/2)), color='firebrick', bbox=caja_b, ha='center', va='center', fontweight='bold', zorder=10)
            ax3.annotate(f"{int(hd)}", (i, hd + (ax3.get_ylim()[1]*0.05)), color='black', bbox=caja_b, ha='center', fontweight='bold', zorder=10)

        ax3.set_xticks(x_idx); ax3.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        agregar_sello_agua(fig3); st.pyplot(fig3)
    else: st.warning("⚠️ Sin datos para GAP.")

with col4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Valorización económica de la ineficiencia</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({'HH_Improductivas': 'sum', 'Costo_Improd._$': 'sum'}).reset_index()
        fig4, ax4 = plt.subplots(figsize=(14, 10)); ax4_line = ax4.twinx()
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig4.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_idx = np.arange(len(ag4))
        bi = ax4.bar(x_idx, ag4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4.bar_label(bi, padding=4, color='black', fontweight='bold', path_effects=efecto_b, zorder=4)
        set_escala_y(ax4, ag4['HH_Improductivas'].max(), 1.4) 
        
        ax4_line.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, label='COSTO ARS', zorder=5)
        add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'])
        ax4_line.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3)); ax4_line.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_line.get_yticks()], fontweight='bold')

        for i, val in enumerate(ag4['Costo_Improd._$']): ax4_line.annotate(f"${val:,.0f}", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax4.set_xticks(x_idx); ax4.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax4_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        agregar_sello_agua(fig4); st.pyplot(fig4)
    else: st.warning("⚠️ No hay datos económicos.")

st.markdown("---")

# =========================================================================
# 9. GRÁFICOS: MÉTRICAS 5 Y 6
# =========================================================================
col5, col6 = st.columns(2)

with col5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)

    if not df_im_f.empty:
        ag5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        n_m = df_im_f['FECHA'].nunique(); div = n_m if n_m > 0 else 1
        ag5['Prom_M'] = ag5['HH_IMPRODUCTIVAS'] / div
        ag5 = ag5.sort_values(by='Prom_M', ascending=False)
        ag5['Pct_Acu'] = (ag5['Prom_M'].cumsum() / ag5['Prom_M'].sum()) * 100

        fig5, ax5 = plt.subplots(figsize=(14, 10)); ax5_line = ax5.twinx()
        fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92); fig5.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_idx = np.arange(len(ag5))
        bp = ax5.bar(x_idx, ag5['Prom_M'], color='maroon', edgecolor='white', zorder=2)
        set_escala_y(ax5, ag5['Prom_M'].max(), 2.8)
        ax5.bar_label(bp, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        ax5_line.plot(x_idx, ag5['Pct_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax5_line.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        ax5_line.set_ylim(0, 110); ax5_line.yaxis.set_major_formatter(mtick.PercentFormatter()) 

        lbls = [textwrap.fill(str(t), 12) for t in ag5['TIPO_PARADA']]
        ax5.set_xticks(x_idx); ax5.set_xticklabels(lbls, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val in enumerate(ag5['Pct_Acu']): ax5_line.annotate(f"{val:.1f}%", (x_idx[i], val + 4), color='white', bbox=caja_g, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
        s_m = ag5['Prom_M'].sum()
        ax5.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{s_m:.1f} HH", transform=ax5.transAxes, bbox=caja_g, color='white', fontsize=15, ha='left', va='top', zorder=10)
        
        agregar_sello_agua(fig5); st.pyplot(fig5)
    else:
        st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Porcentaje histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)

    if not df_ef_f.empty:
        df_ef_f['K_Mes'] = df_ef_f['Fecha'].dt.strftime('%Y-%m')
        ag_disp = df_ef_f.groupby('K_Mes', as_index=False)['HH_Disponibles'].sum()

        if not df_im_f.empty:
            df_im_f['K_Mes'] = df_im_f['FECHA'].dt.strftime('%Y-%m')
            piv = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K_Mes', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6 = pd.merge(ag_disp, piv, on='K_Mes', how='left').fillna(0)
            list_c = [c for c in df6.columns if c not in ['HH_Disponibles', 'K_Mes']]
        else: 
            df6 = ag_disp.copy(); list_c = []
            
        df6['Suma_I'] = df6[list_c].sum(axis=1) if list_c else 0
        df6['Inc_%'] = (df6['Suma_I'] / df6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6['Fecha_O'] = pd.to_datetime(df6['K_Mes'] + '-01'); df6 = df6.sort_values(by='Fecha_O')

        fig6, ax6 = plt.subplots(figsize=(14, 10)); fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
        fig6.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_idx = np.arange(len(df6))
        
        if list_c:
            base_st = np.zeros(len(df6)); paleta = plt.cm.tab20.colors
            for i, c_nom in enumerate(list_c):
                vals = df6[c_nom].values
                bar_stack = ax6.bar(x_idx, vals, bottom=base_st, label=textwrap.fill(c_nom, 15), color=paleta[i % 20], edgecolor='white', zorder=2)
                lbls_stk = [f"{int(v)}" if v > 0 else "" for v in vals]
                ax6.bar_label(bar_stack, labels=lbls_stk, label_type='center', color='white', fontsize=9, fontweight='bold', path_effects=efecto_n)
                base_st += vals
            ax6.legend(loc='upper left', fontsize=8, ncol=3, framealpha=0.7)
        else:
            ax6.bar(x_idx, np.zeros(len(df6)), color='white')

        set_escala_y(ax6, df6['Suma_I'].max(), 1.8) 
        
        for i in range(len(x_idx)):
            v_i, v_d = df6['Suma_I'].iloc[i], df6['HH_Disponibles'].iloc[i]
            if v_i > 0: ax6.annotate(f"Imp: {int(v_i)}\nDisp: {int(v_d)}", (i, v_i + (ax6.get_ylim()[1]*0.02)), ha='center', bbox=caja_o, fontweight='bold', zorder=10)

        ax6_line = ax6.twinx()
        ax6_line.plot(x_idx, df6['Inc_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, label='% Incidencia', zorder=5)
        add_tendencia(ax6_line, x_idx, df6['Inc_%'])
        
        ax6_line.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_line.text(x_idx[0], 16, 'META = 15%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(df6['Inc_%']): 
            if df6['Suma_I'].iloc[i] > 0: ax6_line.annotate(f"{val:.1f}%", (x_idx[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_b, zorder=10)

        ax6.set_xticks(x_idx); ax6.set_xticklabels(df6['K_Mes'], fontsize=14, fontweight='bold')
        ax6_line.set_ylim(0, max(30, df6['Inc_%'].max() * 1.5)); ax6_line.legend(loc='upper right', bbox_to_anchor=(1, 1.02), frameon=True)
        agregar_sello_agua(fig6); st.pyplot(fig6)
    else: st.warning("⚠️ Sin datos históricos de eficiencia para cruzar.")

st.markdown("---")

# =========================================================================
# 10. FILA 4: DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO - METRICA 7)
# =========================================================================
st.header("7. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)")
st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Apertura de registros detallados con fecha, operario y motor de sugerencia de acciones</i></div>", unsafe_allow_html=True)

if not df_im_f.empty and 'DETALLE' in df_im_f.columns:
    motivos_disp = sorted(df_im_f['TIPO_PARADA'].dropna().unique())
    col_sel_m, _ = st.columns([1, 2])
    with col_sel_m:
        motivo_sel = st.selectbox("🔍 Filtrar Motivo a detallar:", ["Todos los motivos"] + list(motivos_disp))
    
    df_detalles = df_im_f.copy()
    if motivo_sel != "Todos los motivos":
        df_detalles = df_detalles[df_detalles['TIPO_PARADA'] == motivo_sel]
        
    if not df_detalles.empty:
        # Preparamos las columnas llenando vacíos para evitar que groupby borre filas
        df_detalles['FECHA_STR'] = df_detalles['FECHA_EXACTA'].dt.strftime('%d/%m/%Y').fillna('S/D')
        df_detalles['OPERARIO'] = df_detalles['OPERARIO'].fillna('S/D')
        df_detalles['DETALLE'] = df_detalles['DETALLE'].fillna('S/D')
        
        # Agrupación por Fecha, Operario y Detalle
        ag_det = df_detalles.groupby(['FECHA_STR', 'OPERARIO', 'DETALLE']).agg({'HH_IMPRODUCTIVAS': 'sum'}).reset_index()
        ag_det = ag_det.sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
        
        t_det = ag_det['HH_IMPRODUCTIVAS'].sum()
        ag_det['%'] = (ag_det['HH_IMPRODUCTIVAS'] / t_det) * 100 if t_det > 0 else 0
        ag_det['Acción Sugerida'] = ag_det['DETALLE'].apply(generar_accion_sugerida)
        
        # Orden y Renombrado Final
        ag_det = ag_det[['FECHA_STR', 'OPERARIO', 'DETALLE', 'HH_IMPRODUCTIVAS', '%', 'Acción Sugerida']]
        ag_det.columns = ['Fecha', 'Operario', 'Detalle Registrado', 'Subtotal HH', '%', 'Acción Sugerida']
        
        fila_tot = pd.DataFrame({
            'Fecha': ['---'], 'Operario': ['---'], 'Detalle Registrado': ['✅ TOTAL SUMATORIA'], 
            'Subtotal HH': [t_det], '%': [100.0], 'Acción Sugerida': ['🎯 ACCIÓN GLOBAL']
        })
        ag_det = pd.concat([ag_det, fila_tot], ignore_index=True)
        
        st.dataframe(ag_det, use_container_width=True, hide_index=True, column_config={
            "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), 
            "%": st.column_config.NumberColumn(format="%.1f %%")
        })
        
        csv_detalles = ag_det.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Detalle Operativo (CSV)", data=csv_detalles, file_name="Detalles_Operativos.csv", mime="text/csv", use_container_width=True, type="primary")
    else:
        st.info("No hay registros detallados para el motivo seleccionado en este periodo.")
else:
    st.info("No hay horas improductivas reportadas con la configuración actual para analizar detalles.")
