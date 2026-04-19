import streamlit as st, pandas as pd, numpy as np
import matplotlib.pyplot as plt, matplotlib.ticker as mtick
import matplotlib.patheffects as pe, matplotlib.image as mpimg
import textwrap, re

# =========================================================================
# 1. CONFIGURACIÓN Y ESCUDO VISUAL (LÍNEA 1)
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
    header, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="manage-app-button"], 
    #MainMenu, footer, .stAppDeployButton, .viewerBadge_container {display: none !important; visibility: hidden !important;}
    .block-container {padding-top: 1rem !important;}
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
        background-color: #0E1117 !important; z-index: 99999 !important;
        padding-top: 15px !important; padding-bottom: 15px !important; border-bottom: 3px solid #1E3A8A !important; 
    }
</style>
""", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
efecto_b, efecto_n = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
caja_v, caja_g = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o, caja_b = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SEGURIDAD
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True); c1, c2, c3 = st.columns([1, 1.8, 1])
    with c2:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        l1, l2, l3 = st.columns([1, 1, 1])
        with l2:
            try: st.image("LOGO OMBÚ.jpg", width=160)
            except: st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        st.markdown("<div style='text-align:center;'><h2 style='color:#1E3A8A;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</h2><p>Acceso Restringido - Control de Gestión</p></div>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in, p_in = st.text_input("Usuario Corporativo"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in: st.session_state['autenticado'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")

if not st.session_state['autenticado']: mostrar_login(); st.stop()

# =========================================================================
# 3. MOTOR INTELIGENTE
# =========================================================================
def set_escala_y(ax, vmax, factor=1.4): ax.set_ylim(0, vmax * factor if vmax > 0 else 100)
def dibujar_meses(ax, n_meses):
    for i in range(n_meses): ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

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
    if not v1 or not v2: return False
    return v1.issubset(v2) or v2.issubset(v1)

def add_tendencia(ax, x, y):
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color='darkgray', linestyle=':', linewidth=4, path_effects=efecto_b, zorder=4, label='Tendencia')

def agregar_sello_agua(fig):
    try:
        img = mpimg.imread("LOGO OMBÚ.jpg"); ax_logo = fig.add_axes([0.88, 0.02, 0.08, 0.08], zorder=1)
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
# 4. CARGA DE DATOS
# =========================================================================
h_l, h_t, h_s = st.columns([1, 3, 1])
with h_l:
    try: st.image("LOGO OMBÚ.jpg", width=120)
    except: st.markdown("### OMBÚ")
with h_t:
    st.title("TABLERO INTEGRADO - REPORTE C.G.P."); st.markdown("<p style='margin-top:-15px; font-weight:bold; color:gray;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</p>", unsafe_allow_html=True)
with h_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir del Tablero", use_container_width=True): st.session_state['autenticado'] = False; st.rerun()

try:
    df_ef, df_im = pd.read_excel("eficiencias.xlsx"), pd.read_excel("improductivas.xlsx")
    df_ef.columns, df_im.columns = df_ef.columns.str.strip(), [str(c).strip().upper() for c in df_im.columns]
    
    if 'TIPO_PARADA' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), None): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'FECHA' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'FECHA' in c or 'INICIO' in c), None): 'FECHA'}, inplace=True)
    if 'DETALLE' not in df_im.columns: df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c or 'DESC' in c), None): 'DETALLE'}, inplace=True)
    
    c_nom, c_ape = next((c for c in df_im.columns if 'NOMBRE' in c), None), next((c for c in df_im.columns if 'APELLIDO' in c), None)
    if c_nom and c_ape: df_im['OPERARIO'] = df_im[c_nom].astype(str) + ' ' + df_im[c_ape].astype(str)
    elif c_nom: df_im['OPERARIO'] = df_im[c_nom].astype(str)
    else: df_im['OPERARIO'] = "S/D"

    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im['FECHA_EXACTA'] = pd.to_datetime(df_im['FECHA'], errors='coerce')
    df_im['FECHA'] = df_im['FECHA_EXACTA'].dt.to_period('M').dt.to_timestamp()
    
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'], df_im['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y'), df_im['FECHA'].dt.strftime('%b-%Y')
except Exception as e: st.error(f"Error crítico: {e}"); st.stop()

# =========================================================================
# 5. FILTROS EN CASCADA
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True); st.markdown("### 🔍 Configuración del Escenario")
    f1, f2, f3, f4 = st.columns(4)
    with f1: s_pl = st.multiselect("🏭 Planta", list(df_ef['Planta'].dropna().unique()), placeholder="Todas")
    df_tl = df_ef[df_ef['Planta'].isin(s_pl)] if s_pl else df_ef
    with f2: s_li = st.multiselect("⚙️ Línea", list(df_tl['Linea'].dropna().unique()), placeholder="Todas")
    df_tp = df_tl[df_tl['Linea'].isin(s_li)] if s_li else df_tl
    with f3: s_pu = st.multiselect("🛠️ Puesto de Trabajo", list(df_tp['Puesto_Trabajo'].dropna().unique()), placeholder="Todos")
    with f4: s_mes = st.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + list(df_ef['Mes_Str'].unique()), placeholder="Todos")

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
    if s_mes and "🎯 Acumulado YTD" not in s_mes: df_im_f = df_im_f[df_im_f['Mes_Str'].isin(s_mes)]

t_enc = f"Filtros: {'+'.join(s_pl) if s_pl else 'Todas'} | {'+'.join(s_li) if s_li else 'Todas'} | {'+'.join(s_pu) if s_pu else 'Todos'}"
st.markdown("---")

# =========================================================================
# 6. FILA 1: MÉTRICAS 1 Y 2
# =========================================================================
col1, col2 = st.columns(2)
warn_linea = False
if s_pu: df_plot_1 = df_ef_f.copy()
elif s_li:
    df_salida = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_salida.empty: df_plot_1 = df_salida
    else: df_plot_1 = df_ef_f.copy(); warn_linea = True
else: df_plot_1 = df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI']

with col1:
    st.header("1. EFICIENCIA REAL"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISP)</i></div>", unsafe_allow_html=True)
    if warn_linea: st.warning("⚠️ Esta Línea NO registra un 'Último Puesto'. Seleccione un Puesto para análisis preciso.")
    if not df_plot_1.empty:
        ag1 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'}).reset_index()
        ag1['Ef_Real'] = (ag1['HH_STD_TOTAL'] / ag1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        fig1, ax1 = plt.subplots(figsize=(14, 10)); ax1_line = ax1.twinx()
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig1.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag1))
        bs = ax1.bar(x_idx - 0.17, ag1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD', zorder=2)
        bd = ax1.bar(x_idx + 0.17, ag1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISP', zorder=2)
        set_escala_y(ax1, ag1['HH_Disponibles'].max(), 1.4); ax1.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3); ax1.bar_label(bd, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax1, len(x_idx))
        for i, bar in enumerate(bs):
            vu = int(ag1['Cant._Prod._A1'].iloc[i])
            if vu > 0: ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{vu} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)
        ax1_line.plot(x_idx, ag1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic.', zorder=5); add_tendencia(ax1_line, x_idx, ag1['Ef_Real'])
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1); ax1_line.text(x_idx[0], 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3)); ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, val in enumerate(ag1['Ef_Real']): ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True); agregar_sello_agua(fig1); st.pyplot(fig1)
    else: st.warning("⚠️ Sin datos.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PROD)</i></div>", unsafe_allow_html=True)
    if not df_plot_1.empty:
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Productivas_C/GAP': 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        fig2, ax2 = plt.subplots(figsize=(14, 10)); ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag2))
        bs = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD', zorder=2)
        bp = ax2.bar(x_idx + 0.17, ag2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PROD', zorder=2)
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 1.4); ax2.bar_label(bs, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3); ax2.bar_label(bp, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3); dibujar_meses(ax2, len(x_idx))
        ax2_line.plot(x_idx, ag2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic.', zorder=5); add_tendencia(ax2_line, x_idx, ag2['Ef_Prod'])
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1); ax2_line.text(x_idx[0], 101, 'META = 100%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        ax2_line.set_ylim(0, max(110, ag2['Ef_Prod'].max()*1.3)); ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, val in enumerate(ag2['Ef_Prod']): ax2_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
        ax2.set_xticks(x_idx); ax2.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax2_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True); agregar_sello_agua(fig2); st.pyplot(fig2)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")

# =========================================================================
# 7. FILA 2: MÉTRICAS 3 Y 4
# =========================================================================
col3, col4 = st.columns(2)
with col3:
    st.header("3. GAP HH GLOBAL"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Desvío entre Horas Disponibles y Declaradas</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        ag3 = df_ef_f.groupby('Fecha').agg({c_prod: 'sum', 'HH_Improductivas': 'sum', 'HH_Disponibles': 'sum'}).reset_index(); ag3['Tot'] = ag3[c_prod] + ag3['HH_Improductivas']
        fig3, ax3 = plt.subplots(figsize=(14, 10)); fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig3.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag3))
        bp = ax3.bar(x_idx, ag3[c_prod], color='darkgreen', edgecolor='white', label='PROD', zorder=2); bi = ax3.bar(x_idx, ag3['HH_Improductivas'], bottom=ag3[c_prod], color='firebrick', edgecolor='white', label='IMP', zorder=2)
        ax3.bar_label(bp, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4); ax3.bar_label(bi, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='DISP', zorder=5)
        set_escala_y(ax3, ag3['HH_Disponibles'].max(), 1.4); dibujar_meses(ax3, len(x_idx))
        for i in range(len(x_idx)):
            hd, ht = ag3['HH_Disponibles'].iloc[i], ag3['Tot'].iloc[i]; gap = hd - ht
            ax3.plot([i, i], [ht, hd], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            ax3.annotate(f"GAP:\n{int(gap)}", (i, ht + (gap/2)), color='firebrick', bbox=caja_b, ha='center', va='center', fontweight='bold', zorder=10)
            ax3.annotate(f"{int(hd)}", (i, hd + (ax3.get_ylim()[1]*0.05)), color='black', bbox=caja_b, ha='center', fontweight='bold', zorder=10)
        ax3.set_xticks(x_idx); ax3.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold'); ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True); agregar_sello_agua(fig3); st.pyplot(fig3)
    else: st.warning("⚠️ Sin datos.")

with col4:
    st.header("4. COSTOS IMPRODUCTIVOS"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Valorización de la ineficiencia</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({'HH_Improductivas': 'sum', 'Costo_Improd._$': 'sum'}).reset_index()
        fig4, ax4 = plt.subplots(figsize=(14, 10)); ax4_line = ax4.twinx(); fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig4.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag4))
        bi = ax4.bar(x_idx, ag4['HH_Improductivas'], color='darkred', edgecolor='white', label='IMP', zorder=2); ax4.bar_label(bi, padding=4, color='black', fontweight='bold', path_effects=efecto_b, zorder=4)
        set_escala_y(ax4, ag4['HH_Improductivas'].max(), 1.4) 
        ax4_line.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, label='COSTO', zorder=5); add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'])
        ax4_line.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3)); ax4_line.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_line.get_yticks()], fontweight='bold')
        t_p, t_h = ag4['Costo_Improd._$'].sum(), ag4['HH_Improductivas'].sum()
        ax4.text(0.5, 0.90, f"COSTO ACUMULADO ARS\n${t_p:,.0f}\nTOTAL: {t_h:,.0f} HH IMP", transform=ax4.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=caja_o, weight='bold', zorder=10)
        for i, val in enumerate(ag4['Costo_Improd._$']): ax4_line.annotate(f"${val:,.0f}", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)
        ax4.set_xticks(x_idx); ax4.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold'); ax4.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax4_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True); agregar_sello_agua(fig4); st.pyplot(fig4)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")

# =========================================================================
# 8. FILA 3: MÉTRICAS 5 Y 6 
# =========================================================================
col5, col6 = st.columns(2)
with col5:
    st.header("5. PARETO DE CAUSAS"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Distribución de motivos de pérdida (80/20)</i></div>", unsafe_allow_html=True)
    if not df_im_f.empty:
        ag5 = df_im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        nm = df_im_f['FECHA'].nunique(); div = nm if nm > 0 else 1
        ag5['Prom'] = ag5['HH_IMPRODUCTIVAS'] / div; ag5 = ag5.sort_values(by='Prom', ascending=False); ag5['Pct'] = (ag5['Prom'].cumsum() / ag5['Prom'].sum()) * 100
        fig5, ax5 = plt.subplots(figsize=(14, 10)); ax5_line = ax5.twinx(); fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92); fig5.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(ag5))
        bp = ax5.bar(x_idx, ag5['Prom'], color='maroon', edgecolor='white', zorder=2); set_escala_y(ax5, ag5['Prom'].max(), 2.8); ax5.bar_label(bp, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        ax5_line.plot(x_idx, ag5['Pct'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax5_line.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1); ax5_line.set_ylim(0, 110); ax5_line.yaxis.set_major_formatter(mtick.PercentFormatter()) 
        ax5.set_xticks(x_idx); ax5.set_xticklabels([textwrap.fill(str(t), 12) for t in ag5['TIPO_PARADA']], rotation=90, fontsize=12, fontweight='bold')
        for i, val in enumerate(ag5['Pct']): ax5_line.annotate(f"{val:.1f}%", (x_idx[i], val + 4), color='white', bbox=caja_g, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
        ax5.text(0.02, 0.96, f"PROMEDIO MENSUAL: {ag5['Prom'].sum():.1f} HH", transform=ax5.transAxes, bbox=caja_g, color='white', fontsize=15, ha='left', va='top', zorder=10)
        agregar_sello_agua(fig5); st.pyplot(fig5)
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl = ag5.copy(); th = df_tbl['HH_IMPRODUCTIVAS'].sum(); df_tbl['%'] = (df_tbl['HH_IMPRODUCTIVAS'] / th) * 100
        df_tbl = pd.concat([df_tbl, pd.DataFrame({'TIPO_PARADA': ['✅ TOTAL'], 'HH_IMPRODUCTIVAS': [th], 'Prom': [df_tbl['Prom'].sum()], 'Pct': [100.0], '%': [100.0]})], ignore_index=True)
        st.dataframe(df_tbl.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), use_container_width=True, hide_index=True)
    else: st.success("✅ ¡Felicitaciones! Cero horas improductivas en este periodo.")

with col6:
    st.header("6. EVOLUCIÓN INCIDENCIA %"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Histórico de HH Improductivas sobre Disponibles</i></div>", unsafe_allow_html=True)
    if not df_ef_f.empty:
        df_ef_f['K'] = df_ef_f['Fecha'].dt.strftime('%Y-%m'); ad = df_ef_f.groupby('K', as_index=False)['HH_Disponibles'].sum()
        if not df_im_f.empty:
            df_im_f['K'] = df_im_f['FECHA'].dt.strftime('%Y-%m'); piv = pd.pivot_table(df_im_f, values='HH_IMPRODUCTIVAS', index='K', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6 = pd.merge(ad, piv, on='K', how='left').fillna(0); lc = [c for c in df6.columns if c not in ['HH_Disponibles', 'K']]
        else: df6 = ad.copy(); lc = []
        df6['S'] = df6[lc].sum(axis=1) if lc else 0; df6['I'] = (df6['S'] / df6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0); df6['O'] = pd.to_datetime(df6['K'] + '-01'); df6 = df6.sort_values(by='O')
        fig6, ax6 = plt.subplots(figsize=(14, 10)); fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig6.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        x_idx = np.arange(len(df6))
        if lc:
            base = np.zeros(len(df6)); pal = plt.cm.tab20.colors
            for i, c_nom in enumerate(lc):
                vals = df6[c_nom].values; bar_stk = ax6.bar(x_idx, vals, bottom=base, label=textwrap.fill(c_nom, 15), color=pal[i % 20], edgecolor='white', zorder=2)
                ax6.bar_label(bar_stk, labels=[f"{int(v)}" if v > 0 else "" for v in vals], label_type='center', color='white', fontsize=9, fontweight='bold', path_effects=efecto_n); base += vals
            ax6.legend(loc='upper left', fontsize=8, ncol=3, framealpha=0.7)
        else: ax6.bar(x_idx, np.zeros(len(df6)), color='white')
        set_escala_y(ax6, df6['S'].max(), 1.8) 
        for i in range(len(x_idx)):
            if df6['S'].iloc[i] > 0: ax6.annotate(f"Imp: {int(df6['S'].iloc[i])}\nDisp: {int(df6['HH_Disponibles'].iloc[i])}", (i, df6['S'].iloc[i] + (ax6.get_ylim()[1]*0.02)), ha='center', bbox=caja_o, fontweight='bold', zorder=10)
        ax6_line = ax6.twinx(); ax6_line.plot(x_idx, df6['I'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, label='% Incidencia', zorder=5); add_tendencia(ax6_line, x_idx, df6['I'])
        ax6_line.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1); ax6_line.text(x_idx[0], 16, 'META = 15%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        for i, val in enumerate(df6['I']): 
            if df6['S'].iloc[i] > 0: ax6_line.annotate(f"{val:.1f}%", (x_idx[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_b, zorder=10)
        ax6.set_xticks(x_idx); ax6.set_xticklabels(df6['K'], fontsize=14, fontweight='bold'); ax6_line.set_ylim(0, max(30, df6['I'].max() * 1.5)); ax6_line.legend(loc='upper right', bbox_to_anchor=(1, 1.02), frameon=True); agregar_sello_agua(fig6); st.pyplot(fig6)
    else: st.warning("⚠️ Sin datos históricos.")

st.markdown("---")

# =========================================================================
# 10. FILA 4: DETALLES CON FECHA, OPERARIO Y ACCIONES
# =========================================================================
st.header("7. DETALLES DE IMPRODUCTIVIDAD"); st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Apertura de registros con motor de acciones y responsables</i></div>", unsafe_allow_html=True)
if not df_im_f.empty and 'DETALLE' in df_im_f.columns:
    c_sel, _ = st.columns([1, 2]); sel_mot = c_sel.selectbox("🔍 Filtrar Motivo:", ["Todos"] + sorted(df_im_f['TIPO_PARADA'].dropna().unique()))
    df_det = df_im_f[df_im_f['TIPO_PARADA'] == sel_mot] if sel_mot != "Todos" else df_im_f.copy()
    if not df_det.empty:
        ag_det = df_det.groupby(['FECHA_EXACTA', 'OPERARIO', 'DETALLE']).agg({'HH_IMPRODUCTIVAS': 'sum'}).reset_index().sort_values('HH_IMPRODUCTIVAS', ascending=False)
        ag_det['FECHA_EXACTA'] = ag_det['FECHA_EXACTA'].dt.strftime('%d/%m/%Y').fillna('S/D')
        t_det = ag_det['HH_IMPRODUCTIVAS'].sum(); ag_det['%'] = (ag_det['HH_IMPRODUCTIVAS'] / t_det) * 100
        ag_det['Acción Sugerida'] = ag_det['DETALLE'].apply(generar_accion_sugerida)
        ag_det = pd.concat([ag_det, pd.DataFrame({'FECHA_EXACTA': ['---'], 'OPERARIO': ['---'], 'DETALLE': ['✅ TOTAL SUMATORIA'], 'HH_IMPRODUCTIVAS': [t_det], '%': [100.0], 'Acción Sugerida': ['🎯 ACCIÓN GLOBAL']})], ignore_index=True)
        st.dataframe(ag_det.rename(columns={'FECHA_EXACTA': 'Fecha', 'OPERARIO': 'Operario', 'DETALLE': 'Detalle Registrado', 'HH_IMPRODUCTIVAS': 'Subtotal HH'}), use_container_width=True, hide_index=True, column_config={"Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), "%": st.column_config.NumberColumn(format="%.1f %%")})
        st.download_button("📥 Descargar Detalle (CSV)", ag_det.to_csv(index=False).encode('utf-8'), "Detalles_Operativos.csv", "text/csv", use_container_width=True, type="primary")
    else: st.info("No hay registros para este motivo.")
else: st.info("No hay horas improductivas reportadas con la configuración actual para analizar detalles.")
