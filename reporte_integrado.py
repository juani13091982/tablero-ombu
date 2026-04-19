import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import textwrap
import re

# =========================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESCUDO VISUAL TOTAL (LINEA 1)
# =========================================================================
st.set_page_config(
    page_title="C.G.P. Reporte Integrado - Ombú S.A.", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# BLOQUEO NUCLEAR: Oculta GitHub, Share, Deploy, Menú y Decoración superior
st.markdown("""
    <style>
        /* Ocultar elementos de Streamlit Cloud */
        header, [data-testid="stHeader"], .stAppToolbar, .stAppDeployButton, 
        [data-testid="stDecoration"], #MainMenu, footer, .viewerBadge_container,
        [data-testid="manage-app-button"] {
            display: none !important; 
            visibility: hidden !important; 
            height: 0px !important; 
            width: 0px !important;
            opacity: 0 !important;
        }
        
        /* Ajuste de márgenes para ganar espacio */
        .block-container {
            padding-top: 0rem !important;
            padding-bottom: 1rem !important;
        }

        /* Panel de filtros fijo en la parte superior (Sticky) */
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
""", unsafe_allow_html=True)

# Configuración Maestra de Fuentes para Gráficos Gerenciales
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold', 
    'axes.labelweight': 'bold', 
    'axes.titleweight': 'bold', 
    'figure.titlesize': 18
})

# Efectos de contorno para etiquetas de alta legibilidad
efecto_b = [pe.withStroke(linewidth=3, foreground='white')]
efecto_n = [pe.withStroke(linewidth=3, foreground='black')]

# Estilos de cajas de texto (KPIs)
caja_v = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5)
caja_g = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
caja_o = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5)
caja_b = dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

# =========================================================================
# 2. SISTEMA DE SEGURIDAD (ACCESO RESTRINGIDO)
# =========================================================================
USUARIOS_PERMITIDOS = {"acceso.ombu": "Gestion2026"}

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

def mostrar_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_v1, col_login, col_v2 = st.columns([1, 1.8, 1])
    
    with col_login:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        
        inner_l, inner_c, inner_r = st.columns([1, 1, 1])
        with inner_c:
            try: st.image("LOGO OMBÚ.jpg", width=160)
            except: st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        
        st.markdown("""
            <div style='text-align:center; margin-top:-10px; margin-bottom:20px;'>
                <h2 style='margin:0; color:#1E3A8A; font-weight:bold;'>GESTIÓN INDUSTRIAL PRODUCTIVA OMBÚ S.A.</h2>
                <p style='color:#666; font-weight:bold;'>Acceso Restringido - Control de Gestión</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_login"):
            st.markdown("<h4 style='text-align: center;'>🔒 Iniciar Sesión</h4>", unsafe_allow_html=True)
            u_in = st.text_input("Usuario Corporativo")
            p_in = st.text_input("Contraseña", type="password")
            
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if u_in in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[u_in] == p_in:
                    st.session_state['autenticado'] = True
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas. Verifique los datos.")

if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. MOTOR MATEMÁTICO (TENDENCIAS Y ESCALAS)
# =========================================================================
def set_escala_y(ax, vmax, factor=1.4):
    """Ajusta el eje Y para que los datos ocupen casi toda la pantalla (factor bajo)"""
    if vmax > 0: ax.set_ylim(0, vmax * factor)
    else: ax.set_ylim(0, 100)

def dibujar_meses(ax, n_meses):
    for i in range(n_meses):
        ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def cruce_robusto(sel, excel):
    """Busca coincidencias de nombres entre archivos evitando falsos positivos"""
    if pd.isna(excel) or pd.isna(sel): return False
    s1, s2 = str(sel).upper(), str(excel).upper()
    for a,b in zip("ÁÉÍÓÚ", "AEIOU"): s1, s2 = s1.replace(a,b), s2.replace(a,b)
    
    l1, l2 = re.sub(r'[^A-Z0-9]', '', s1), re.sub(r'[^A-Z0-9]', '', s2)
    if not l1 or not l2: return False
    if l1 in l2 or l2 in l1: return True
    
    p1, p2 = set(re.findall(r'[A-Z0-9]{3,}', s1)), set(re.findall(r'[A-Z0-9]{3,}', s2))
    excl = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'AREA', 'MAQUINA'}
    v1, v2 = p1 - excl, p2 - excl
    return v1.issubset(v2) or v2.issubset(v1) if v1 and v2 else False

def add_tendencia(ax, x, y):
    """Inyecta línea de tendencia lineal punteada"""
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color='darkgray', linestyle=':', linewidth=4, path_effects=efecto_b, zorder=4, label='Tendencia')

# =========================================================================
# 4. CARGA DE BASES EXCEL
# =========================================================================
try:
    df_ef = pd.read_excel("eficiencias.xlsx")
    df_im = pd.read_excel("improductivas.xlsx")
    
    df_ef.columns = df_ef.columns.str.strip()
    df_im.columns = [str(c).strip().upper() for c in df_im.columns]
    
    # Auto-identificador de columnas clave
    if 'TIPO_PARADA' not in df_im.columns:
        df_im.rename(columns={next(c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in df_im.columns:
        df_im.rename(columns={next(c for c in df_im.columns if 'HH' in c and 'IMP' in c): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'FECHA' not in df_im.columns:
        df_im.rename(columns={next(c for c in df_im.columns if 'FECHA' in c): 'FECHA'}, inplace=True)
    if 'DETALLE' not in df_im.columns:
        df_im.rename(columns={next((c for c in df_im.columns if 'DETALLE' in c or 'OBS' in c), df_im.columns[0]): 'DETALLE'}, inplace=True)
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha']).dt.to_period('M').dt.to_timestamp()
    df_im['FECHA'] = pd.to_datetime(df_im['FECHA']).dt.to_period('M').dt.to_timestamp()
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_im['Mes_Str'] = df_im['FECHA'].dt.strftime('%b-%Y')
except Exception as e:
    st.error(f"Error crítico cargando archivos: {e}"); st.stop()

# =========================================================================
# 5. PANEL DE FILTROS (STICKY)
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    f1, f2, f3, f4 = st.columns(4)
    with f1: s_pl = st.multiselect("🏭 Planta", list(df_ef['Planta'].dropna().unique()))
    df_tl = df_ef[df_ef['Planta'].isin(s_pl)] if s_pl else df_ef
    with f2: s_li = st.multiselect("⚙️ Línea", list(df_tl['Linea'].dropna().unique()))
    df_tp = df_tl[df_tl['Linea'].isin(s_li)] if s_li else df_tl
    with f3: s_pu = st.multiselect("🛠️ Puesto", list(df_tp['Puesto_Trabajo'].dropna().unique()))
    with f4: s_me = st.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + list(df_ef['Mes_Str'].unique()))

# PROCESO DE FILTRADO FINAL
ef_f, im_f = df_ef.copy(), df_im.copy()
if s_pl: ef_f = ef_f[ef_f['Planta'].isin(s_pl)]
if s_li: ef_f = ef_f[ef_f['Linea'].isin(s_li)]
if s_pu: ef_f = ef_f[ef_f['Puesto_Trabajo'].isin(s_pu)]
if s_me and "🎯 Acumulado YTD" not in s_me: ef_f = ef_f[ef_f['Mes_Str'].isin(s_me)]

if not im_f.empty:
    if s_pl: im_f = im_f[im_f.iloc[:,0].apply(lambda x: any(cruce_robusto(p, x) for p in s_pl))]
    if s_li: 
        cli = next((c for c in im_f.columns if 'LINEA' in c), im_f.columns[1])
        im_f = im_f[im_f[cli].apply(lambda x: any(cruce_robusto(l, x) for l in s_li))]
    if s_pu:
        cpu = next((c for c in im_f.columns if 'PUESTO' in c), im_f.columns[2])
        im_f = im_f[im_f[cpu].apply(lambda x: any(cruce_robusto(p, x) for p in s_pu))]
    if s_me and "🎯 Acumulado YTD" not in s_me: im_f = im_f[im_f['Mes_Str'].isin(s_me)]

t_enc = f"Filtros >> Planta: {'+'.join(s_pl) if s_pl else 'Todas'} | Línea: {'+'.join(s_li) if s_li else 'Todas'} | Puesto: {'+'.join(s_pu) if s_pu else 'Todos'}"
st.markdown("---")

# =========================================================================
# 6. FILA 1: MÉTRICAS 1 Y 2 (LÓGICA DE SALIDA ESTRICTA)
# =========================================================================
c1, c2 = st.columns(2)

# REGLA DE NEGOCIO: Embudo de Eficiencia
warn_linea = False
if s_pu:
    df_m1_plot = ef_f.copy() # Si se elige un puesto, se evalúa ese puesto.
elif s_li:
    df_check = ef_f[ef_f['Es_Ultimo_Puesto'] == 'SI']
    if not df_check.empty:
        df_m1_plot = df_check # Prioridad: Salida de línea.
    else:
        df_m1_plot = ef_f.copy() # Excepción: Línea sin terminales (Limpieza).
        warn_linea = True
else:
    df_m1_plot = ef_f[ef_f['Es_Ultimo_Puesto'] == 'SI'] # Vista General: solo salidas.

with c1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<p style='color:#aaa; font-size:14px;'>(∑ HH STD / ∑ HH DISPONIBLES)</p>", unsafe_allow_html=True)
    if warn_linea: st.warning("⚠️ Esta Línea NO registra un 'Último Puesto'. Seleccione un Puesto para un análisis de rendimiento preciso.")
    
    if not df_m1_plot.empty:
        ag1 = df_m1_plot.groupby('Fecha').agg({'HH_STD_TOTAL':'sum','HH_Disponibles':'sum','Cant._Prod._A1':'sum'}).reset_index()
        ag1['Ef'] = (ag1['HH_STD_TOTAL']/ag1['HH_Disponibles']).replace([np.inf,-np.inf],0).fillna(0)*100
        
        fig1, ax = plt.subplots(figsize=(14,10)); ax2 = ax.twinx()
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig1.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        
        x_idx = np.arange(len(ag1)); w=0.35
        bs = ax.bar(x_idx-w/2, ag1['HH_STD_TOTAL'], w, color='midnightblue', edgecolor='white', label='HH STD', zorder=2)
        bd = ax.bar(x_idx+w/2, ag1['HH_Disponibles'], w, color='black', edgecolor='white', label='HH DISP', zorder=2)
        
        set_escala_y(ax, ag1['HH_Disponibles'].max(), 1.4)
        ax.bar_label(bs, padding=4, fontweight='bold', path_effects=efecto_b, fmt='%.0f')
        ax.bar_label(bd, padding=4, fontweight='bold', path_effects=efecto_b, fmt='%.0f')
        dibujar_meses(ax, len(x_idx))

        for i, b in enumerate(bs):
            if ag1['Cant._Prod._A1'].iloc[i] > 0: 
                ax.text(b.get_x()+b.get_width()/2, b.get_height()*0.05, f"{int(ag1['Cant._Prod._A1'].iloc[i])} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n)

        ax2.plot(x_idx, ag1['Ef'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic.', zorder=5)
        add_tendencia(ax2, x_idx, ag1['Ef'])
        ax2.axhline(85, color='darkgreen', linestyle='--', linewidth=3)
        ax2.set_ylim(0, max(100, ag1['Ef'].max()*1.3)); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i,v in enumerate(ag1['Ef']): ax2.annotate(f"{v:.1f}%", (x_idx[i],v+5), color='white', bbox=caja_g, ha='center')
        ax.set_xticks(x_idx); ax.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y')); st.pyplot(fig1)

with c2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<p style='color:#aaa; font-size:14px;'>(∑ HH STD / ∑ HH PRODUCTIVAS)</p>", unsafe_allow_html=True)
    if not df_m1_plot.empty:
        ag2 = df_m1_plot.groupby('Fecha').agg({'HH_STD_TOTAL':'sum','HH_Productivas_C/GAP':'sum'}).reset_index()
        ag2['Ef'] = (ag2['HH_STD_TOTAL']/ag2['HH_Productivas_C/GAP']).replace([np.inf,-np.inf],0).fillna(0)*100
        
        fig2, ax = plt.subplots(figsize=(14,10)); ax2 = ax.twinx()
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        
        x_idx = np.arange(len(ag2))
        bs = ax.bar(x_idx-0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD', zorder=2)
        bp = ax.bar(x_idx+0.17, ag2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PROD', zorder=2)
        
        set_escala_y(ax, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 1.4)
        ax.bar_label(bs, padding=4, fontweight='bold', path_effects=efecto_b, fmt='%.0f')
        ax.bar_label(bp, padding=4, fontweight='bold', path_effects=efecto_b, fmt='%.0f')
        dibujar_meses(ax, len(x_idx))

        ax2.plot(x_idx, ag2['Ef'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic.', zorder=5)
        add_tendencia(ax2, x_idx, ag2['Ef'])
        ax2.axhline(100, color='darkgreen', linestyle='--', linewidth=3)
        ax2.set_ylim(0, max(110, ag2['Ef'].max()*1.3)); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i,v in enumerate(ag2['Ef']): ax2.annotate(f"{v:.1f}%", (x_idx[i],v+5), color='white', bbox=caja_g, ha='center')
        ax.set_xticks(x_idx); ax.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y')); st.pyplot(fig2)

st.markdown("---")

# =========================================================================
# 7. FILA 2: MÉTRICAS 3 Y 4
# =========================================================================
c3, c4 = st.columns(2)

with c3:
    st.header("3. GAP HH GLOBAL")
    if not ef_f.empty:
        cp = 'HH_Productivas' if 'HH_Productivas' in ef_f.columns else 'HH Productivas'
        ag3 = ef_f.groupby('Fecha').agg({cp:'sum','HH_Improductivas':'sum','HH_Disponibles':'sum'}).reset_index()
        ag3['Tot'] = ag3[cp] + ag3['HH_Improductivas']
        
        fig3, ax = plt.subplots(figsize=(14, 10)); fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig3.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x_idx = np.arange(len(ag3))
        
        bp = ax.bar(x_idx, ag3[cp], color='darkgreen', edgecolor='white', label='PROD', zorder=2)
        bi = ax.bar(x_idx, ag3['HH_Improductivas'], bottom=ag3[cp], color='firebrick', edgecolor='white', label='IMP', zorder=2)
        ax.bar_label(bp, label_type='center', color='white', fontweight='bold', fmt='%.0f')
        ax.bar_label(bi, label_type='center', color='white', fontweight='bold', fmt='%.0f')
        
        ax.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='DISP', zorder=5)
        set_escala_y(ax, ag3['HH_Disponibles'].max(), 1.4)
        
        for i in range(len(x_idx)):
            hd, ht = ag3['HH_Disponibles'].iloc[i], ag3['Tot'].iloc[i]; gap = hd - ht
            ax.plot([i, i], [ht, hd], color='dimgray', linewidth=5, alpha=0.6)
            ax.annotate(f"GAP:\n{int(gap)}", (i, ht + (gap/2)), color='firebrick', bbox=caja_b, ha='center', va='center', fontweight='bold')
            ax.annotate(f"{int(hd)}", (i, hd + (ax.get_ylim()[1]*0.05)), color='black', bbox=caja_b, ha='center', fontweight='bold')
        ax.set_xticks(x_idx); ax.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y')); st.pyplot(fig3)

with c4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    if not ef_f.empty:
        ag4 = ef_f.groupby('Fecha').agg({'HH_Improductivas': 'sum', 'Costo_Improd._$': 'sum'}).reset_index()
        fig4, ax = plt.subplots(figsize=(14, 10)); ax2 = ax.twinx(); fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig4.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        
        x_idx = np.arange(len(ag4))
        bi = ax.bar(x_idx, ag4['HH_Improductivas'], color='darkred', edgecolor='white', zorder=2); ax.bar_label(bi, padding=4, fontweight='bold', path_effects=efecto_b)
        set_escala_y(ax, ag4['HH_Improductivas'].max(), 1.4)
        
        ax2.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, zorder=5)
        add_tendencia(ax2, x_idx, ag4['Costo_Improd._$'])
        ax2.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3)); ax2.set_yticklabels([f'${int(v/1000000)}M' for v in ax2.get_yticks()])
        
        ax.text(0.5, 0.90, f"COSTO ACUMULADO: ${ag4['Costo_Improd._$'].sum():,.0f}\nTOTAL: {ag4['HH_Improductivas'].sum():,.0f} HH IMP", transform=ax.transAxes, ha='center', va='top', fontsize=18, bbox=caja_o)
        for i,v in enumerate(ag4['Costo_Improd._$']): ax2.annotate(f"${v:,.0f}", (x_idx[i],v+5), color='white', bbox=caja_g, ha='center')
        ax.set_xticks(x_idx); ax.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y')); st.pyplot(fig4)

st.markdown("---")

# =========================================================================
# 8. FILA 3: MÉTRICAS 5 Y 6
# =========================================================================
c5, c6 = st.columns(2)

with c5:
    st.header("5. PARETO DE CAUSAS")
    if not im_f.empty:
        ag5 = im_f.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        nm = im_f['FECHA'].nunique() or 1; ag5['Prom'] = ag5['HH_IMPRODUCTIVAS'] / nm
        ag5 = ag5.sort_values(by='Prom', ascending=False); ag5['Pct'] = (ag5['Prom'].cumsum() / ag5['Prom'].sum()) * 100
        
        fig5, ax = plt.subplots(figsize=(14, 10)); ax2 = ax.twinx(); fig5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92)
        x_idx = np.arange(len(ag5))
        bp = ax.bar(x_idx, ag5['Prom'], color='maroon', edgecolor='white', zorder=2); ax.bar_label(bp, padding=4, fontweight='bold', fmt='%.1f')
        set_escala_y(ax, ag5['Prom'].max(), 1.4)
        
        ax2.plot(x_idx, ag5['Pct'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax2.axhline(80, color='gray', linestyle='--'); ax2.set_ylim(0, 110); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        ax.set_xticks(x_idx); ax.set_xticklabels([textwrap.fill(str(t), 12) for t in ag5['TIPO_PARADA']], rotation=90)
        st.pyplot(fig5)
        st.dataframe(ag5.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), use_container_width=True, hide_index=True)

with c6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    if not ef_f.empty:
        ef_f['K'] = ef_f['Fecha'].dt.strftime('%Y-%m'); ad = ef_f.groupby('K', as_index=False)['HH_Disponibles'].sum()
        if not im_f.empty:
            im_f['K'] = im_f['FECHA'].dt.strftime('%Y-%m')
            pv = pd.pivot_table(im_f, values='HH_IMPRODUCTIVAS', index='K', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            df6 = pd.merge(ad, pv, on='K', how='left').fillna(0); lc = [c for c in df6.columns if c not in ['HH_Disponibles', 'K']]
        else: df6 = ad.copy(); lc = []
        
        df6['S'] = df6[lc].sum(axis=1) if lc else 0; df6['I'] = (df6['S']/df6['HH_Disponibles']*100).replace([np.inf,-np.inf],0).fillna(0)
        df6['O'] = pd.to_datetime(df6['K']+'-01'); df6 = df6.sort_values('O')
        
        fig6, ax = plt.subplots(figsize=(14, 10)); ax2 = ax.twinx(); fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        x_idx = np.arange(len(df6))
        
        if lc:
            base = np.zeros(len(df6)); pal = plt.cm.tab20.colors
            for i, c in enumerate(lc):
                vals = df6[c].values
                bar = ax.bar(x_idx, vals, bottom=base, label=textwrap.fill(c, 15), color=pal[i%20], edgecolor='white', zorder=2)
                ax.bar_label(bar, labels=[f"{int(v)}" if v>0 else "" for v in vals], label_type='center', color='white', fontsize=9, fontweight='bold', path_effects=efecto_n)
                base += vals
        
        ax.set_ylim(0, df6['S'].max()*1.8 if not df6.empty else 100)
        for i in range(len(x_idx)):
            if df6['S'].iloc[i] > 0: ax.annotate(f"Imp: {int(df6['S'].iloc[i])}\nDisp: {int(df6['HH_Disponibles'].iloc[i])}", (i, df6['S'].iloc[i]+2), ha='center', bbox=caja_o, fontweight='bold')
        
        ax2.plot(x_idx, df6['I'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, zorder=5)
        add_tendencia(ax2, x_idx, df6['I']); ax2.axhline(15, color='darkgreen', linestyle='--', linewidth=3)
        ax2.set_ylim(0, max(25, df6['I'].max()*1.3)); ax.set_xticks(x_idx); ax.set_xticklabels(df6['K']); st.pyplot(fig6)

# =========================================================================
# 9. FILA 4: DETALLES DE IMPRODUCTIVIDAD
# =========================================================================
st.markdown("---")
st.header("7. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)")
if not im_f.empty:
    sel_mot = st.selectbox("🔍 Filtrar Motivo a detallar:", ["Todos"] + sorted(list(im_f['TIPO_PARADA'].unique())))
    df_det = im_f[im_f['TIPO_PARADA'] == sel_mot] if sel_mot != "Todos" else im_f
    st.dataframe(df_det[['FECHA', 'TIPO_PARADA', 'DETALLE', 'HH_IMPRODUCTIVAS']].sort_values('HH_IMPRODUCTIVAS', ascending=False), use_container_width=True, hide_index=True)
