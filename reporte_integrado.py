import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patheffects as pe
import textwrap
import re

# =========================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# =========================================================================
st.set_page_config(page_title="C.G.P. Reporte Integrado - Ombú", layout="wide", initial_sidebar_state="collapsed")

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
                    st.error("❌ Credenciales incorrectas.")

if not st.session_state['autenticado']:
    mostrar_login()
    st.stop()

# =========================================================================
# 3. ESTILOS VISUALES Y FILTROS STICKY (CSS)
# =========================================================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden !important;} header {visibility: hidden !important;} footer {visibility: hidden !important;}
    div[data-testid="stVerticalBlock"] > div:has(#filtro-ribbon) {
        position: -webkit-sticky !important; position: sticky !important; top: 0px !important;
        background-color: #0E1117 !important; z-index: 99999 !important;
        padding-top: 15px !important; padding-bottom: 15px !important; border-bottom: 3px solid #1E3A8A !important; 
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
# 4. MOTOR INTELIGENTE Y TENDENCIAS
# =========================================================================
def set_escala_y(ax, vmax, factor=1.4):
    if vmax > 0: ax.set_ylim(0, vmax * factor)
    else: ax.set_ylim(0, 100)

def dibujar_meses(ax, n_meses):
    for i in range(n_meses):
        ax.axvline(x=i, color='lightgray', linestyle='--', linewidth=1, zorder=0)

def cruce_robusto(sel, excel):
    if pd.isna(excel) or pd.isna(sel): return False
    s1 = str(sel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    s2 = str(excel).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    l1, l2 = re.sub(r'[^A-Z0-9]', '', s1), re.sub(r'[^A-Z0-9]', '', s2)
    if not l1 or not l2: return False
    if l1 in l2 or l2 in l1: return True
    
    p1, p2 = set(re.findall(r'[A-Z0-9]{3,}', s1)), set(re.findall(r'[A-Z0-9]{3,}', s2))
    excl = {'SECTOR', 'PUESTO', 'TRABAJO', 'LINEA', 'PLANTA', 'AREA', 'ÁREA', 'MAQUINA'}
    v1, v2 = p1 - excl, p2 - excl
    
    if not v1 or not v2: return False
    if v1.issubset(v2) or v2.issubset(v1): return True
    return False

def add_tendencia(ax, x, y):
    if len(x) > 1:
        z = np.polyfit(x, y.astype(float), 1); p = np.poly1d(z)
        ax.plot(x, p(x), color='darkgray', linestyle=':', linewidth=4, path_effects=efecto_b, zorder=4, label='Tendencia')

# =========================================================================
# 5. HEADER Y CARGA DE DATOS
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
    
    if 'TIPO_PARADA' not in df_im.columns:
        cmot = next((c for c in df_im.columns if 'TIPO' in c or 'MOTIVO' in c or 'CAUSA' in c), None)
        if cmot: df_im.rename(columns={cmot: 'TIPO_PARADA'}, inplace=True)
            
    if 'HH_IMPRODUCTIVAS' not in df_im.columns:
        chs = next((c for c in df_im.columns if 'HH' in c and 'IMP' in c), None)
        if chs: df_im.rename(columns={chs: 'HH_IMPRODUCTIVAS'}, inplace=True)
            
    if 'FECHA' not in df_im.columns:
        cfec = next((c for c in df_im.columns if 'FECHA' in c), None)
        if cfec: df_im.rename(columns={cfec: 'FECHA'}, inplace=True)

    # Identificación automática de la columna DETALLE
    if 'DETALLE' not in df_im.columns:
        cdet = next((c for c in df_im.columns if 'DETALLE' in c or 'OBSERVACION' in c or 'DESCRIPCION' in c), None)
        if cdet: df_im.rename(columns={cdet: 'DETALLE'}, inplace=True)
    
    df_ef['Fecha'] = pd.to_datetime(df_ef['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    df_im['FECHA'] = pd.to_datetime(df_im['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    
    df_ef['Es_Ultimo_Puesto'] = df_ef['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper()
    df_ef['Mes_Str'] = df_ef['Fecha'].dt.strftime('%b-%Y')
    df_im['Mes_Str'] = df_im['FECHA'].dt.strftime('%b-%Y')
    
except Exception as e:
    st.error(f"Error crítico en datos: {e}"); st.stop()

# =========================================================================
# 6. FILTROS EN CASCADA
# =========================================================================
with st.container():
    st.markdown('<div id="filtro-ribbon"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Configuración del Escenario")
    
    f1, f2, f3, f4 = st.columns(4)
    
    with f1: 
        s_pl = st.multiselect("🏭 Planta", list(df_ef['Planta'].dropna().unique()), placeholder="Todas")
    df_tl = df_ef[df_ef['Planta'].isin(s_pl)] if s_pl else df_ef
    
    with f2: 
        s_li = st.multiselect("⚙️ Línea", list(df_tl['Linea'].dropna().unique()), placeholder="Todas")
    df_tp = df_tl[df_tl['Linea'].isin(s_li)] if s_li else df_tl
    
    with f3: 
        s_pu = st.multiselect("🛠️ Puesto de Trabajo", list(df_tp['Puesto_Trabajo'].dropna().unique()), placeholder="Todos")
        
    with f4: 
        s_mes = st.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + list(df_ef['Mes_Str'].unique()), placeholder="Todos")

# =========================================================================
# 7. PROCESAMIENTO FINAL (FILTRADO DE BASES)
# =========================================================================
df_ef_f = df_ef.copy()
df_im_f = df_im.copy()

if s_pl: df_ef_f = df_ef_f[df_ef_f['Planta'].isin(s_pl)]
if s_li: df_ef_f = df_ef_f[df_ef_f['Linea'].isin(s_li)]
if s_pu: df_ef_f = df_ef_f[df_ef_f['Puesto_Trabajo'].isin(s_pu)]
if s_mes and "🎯 Acumulado YTD" not in s_mes: df_ef_f = df_ef_f[df_ef_f['Mes_Str'].isin(s_mes)]

if s_pl and not df_im_f.empty:
    col_pl = next((c for c in df_im_f.columns if 'PLANTA' in str(c).upper()), df_im_f.columns[0] if len(df_im_f.columns) > 0 else None)
    if col_pl: df_im_f = df_im_f[df_im_f[col_pl].apply(lambda x: any(cruce_robusto(p, x) for p in s_pl))]

if s_li and not df_im_f.empty:
    col_li = next((c for c in df_im_f.columns if 'LINEA' in str(c).upper() or 'LÍNEA' in str(c).upper()), df_im_f.columns[1] if len(df_im_f.columns) > 1 else None)
    if col_li: df_im_f = df_im_f[df_im_f[col_li].apply(lambda x: any(cruce_robusto(l, x) for l in s_li))]

if s_pu and not df_im_f.empty:
    col_pu = next((c for c in df_im_f.columns if 'PUESTO' in str(c).upper()), df_im_f.columns[2] if len(df_im_f.columns) > 2 else None)
    if col_pu: df_im_f = df_im_f[df_im_f[col_pu].apply(lambda x: any(cruce_robusto(ps, x) for ps in s_pu))]

if s_mes and "🎯 Acumulado YTD" not in s_mes and not df_im_f.empty:
    df_im_f = df_im_f[df_im_f['Mes_Str'].isin(s_mes)]

t_enc = f"Filtros: {'+'.join(s_pl) if s_pl else 'Todas'} | {'+'.join(s_li) if s_li else 'Todas'} | {'+'.join(s_pu) if s_pu else 'Todos'}"
st.markdown("---")

# =========================================================================
# 8. FILA 1: MÉTRICAS 1 Y 2 (EFICIENCIAS)
# =========================================================================
col1, col2 = st.columns(2)

with col1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></div>", unsafe_allow_html=True)
    
    df_plot_1 = df_ef_f.copy() if (s_pu or s_li) else (df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI'] if not df_ef_f[df_ef_f['Es_Ultimo_Puesto'] == 'SI'].empty else df_ef_f.copy())
    
    if not df_plot_1.empty:
        ag1 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Disponibles': 'sum', 'Cant._Prod._A1': 'sum'}).reset_index()
        ag1['Ef_Real'] = (ag1['HH_STD_TOTAL'] / ag1['HH_Disponibles']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig1, ax1 = plt.subplots(figsize=(14, 10))
        ax1_line = ax1.twinx()
        fig1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig1.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag1))
        bar_s = ax1.bar(x_idx - 0.17, ag1['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        bar_d = ax1.bar(x_idx + 0.17, ag1['HH_Disponibles'], 0.35, color='black', edgecolor='white', label='HH DISPONIBLES', zorder=2)
        
        set_escala_y(ax1, ag1['HH_Disponibles'].max(), 1.4) 
        ax1.bar_label(bar_s, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax1.bar_label(bar_d, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax1, len(x_idx))

        for i, bar in enumerate(bar_s):
            val_u = int(ag1['Cant._Prod._A1'].iloc[i])
            if val_u > 0: ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height()*0.05, f"{val_u} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, fontweight='bold', path_effects=efecto_n, zorder=4)

        ax1_line.plot(x_idx, ag1['Ef_Real'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Real', zorder=5)
        add_tendencia(ax1_line, x_idx, ag1['Ef_Real'])
        
        ax1_line.axhline(85, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax1_line.text(x_idx[0], 86, 'META = 85%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        
        ax1_line.set_ylim(0, max(100, ag1['Ef_Real'].max()*1.3)) 
        ax1_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag1['Ef_Real']):
            ax1_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax1.set_xticks(x_idx); ax1.set_xticklabels(ag1['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax1_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        st.pyplot(fig1)
    else: st.warning("⚠️ Sin datos para Eficiencia Real.")

with col2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></div>", unsafe_allow_html=True)
    
    if not df_plot_1.empty:
        ag2 = df_plot_1.groupby('Fecha').agg({'HH_STD_TOTAL': 'sum', 'HH_Productivas_C/GAP': 'sum'}).reset_index()
        ag2['Ef_Prod'] = (ag2['HH_STD_TOTAL'] / ag2['HH_Productivas_C/GAP']).replace([np.inf, -np.inf], 0).fillna(0) * 100
        
        fig2, ax2 = plt.subplots(figsize=(14, 10))
        ax2_line = ax2.twinx()
        fig2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); fig2.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag2))
        b_s2 = ax2.bar(x_idx - 0.17, ag2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD TOTAL', zorder=2)
        b_p2 = ax2.bar(x_idx + 0.17, ag2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        
        set_escala_y(ax2, max(ag2['HH_STD_TOTAL'].max(), ag2['HH_Productivas_C/GAP'].max()), 1.4) 
        ax2.bar_label(b_s2, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        ax2.bar_label(b_p2, padding=4, color='black', fontweight='bold', path_effects=efecto_b, fmt='%.0f', zorder=3)
        dibujar_meses(ax2, len(x_idx))

        ax2_line.plot(x_idx, ag2['Ef_Prod'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=efecto_b, label='% Efic. Prod.', zorder=5)
        add_tendencia(ax2_line, x_idx, ag2['Ef_Prod'])
        
        ax2_line.axhline(100, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax2_line.text(x_idx[0], 101, 'META = 100%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        
        ax2_line.set_ylim(0, max(110, ag2['Ef_Prod'].max()*1.3)) 
        ax2_line.yaxis.set_major_formatter(mtick.PercentFormatter())

        for i, val in enumerate(ag2['Ef_Prod']):
            ax2_line.annotate(f"{val:.1f}%", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax2.set_xticks(x_idx); ax2.set_xticklabels(ag2['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax2.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax2_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        st.pyplot(fig2)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")

# =========================================================================
# 10. FILA 2: MÉTRICAS 3 Y 4 (BRECHA Y COSTOS)
# =========================================================================
col3, col4 = st.columns(2)

with col3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Desvío entre Horas Disponibles y Declaradas Totales</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        c_prod = 'HH_Productivas' if 'HH_Productivas' in df_ef_f.columns else 'HH Productivas'
        ag3 = df_ef_f.groupby('Fecha').agg({c_prod: 'sum', 'HH_Improductivas': 'sum', 'HH_Disponibles': 'sum'}).reset_index()
        ag3['Total_Decl'] = ag3[c_prod] + ag3['HH_Improductivas']
        
        fig3, ax3 = plt.subplots(figsize=(14, 10))
        fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig3.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')
        
        x_idx = np.arange(len(ag3))
        
        b_p = ax3.bar(x_idx, ag3[c_prod], color='darkgreen', edgecolor='white', label='HH PRODUCTIVAS', zorder=2)
        b_i = ax3.bar(x_idx, ag3['HH_Improductivas'], bottom=ag3[c_prod], color='firebrick', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax3.bar_label(b_p, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        ax3.bar_label(b_i, label_type='center', color='white', fontweight='bold', fmt='%.0f', zorder=4)
        
        ax3.plot(x_idx, ag3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=efecto_b, label='HH DISPONIBLES', zorder=5)
        
        set_escala_y(ax3, ag3['HH_Disponibles'].max(), 1.4) 
        dibujar_meses(ax3, len(x_idx))

        for i in range(len(x_idx)):
            h_dis = ag3['HH_Disponibles'].iloc[i]; h_dec = ag3['Total_Decl'].iloc[i]; gap = h_dis - h_dec
            ax3.plot([i, i], [h_dec, h_dis], color='dimgray', linewidth=5, alpha=0.6, zorder=3)
            ax3.annotate(f"GAP:\n{int(gap)}", (i, h_dec + (gap/2)), color='firebrick', bbox=caja_b, ha='center', va='center', fontweight='bold', zorder=10)
            ax3.annotate(f"{int(h_dis)}", (i, h_dis + (ax3.get_ylim()[1]*0.05)), color='black', bbox=caja_b, ha='center', fontweight='bold', zorder=10)

        ax3.set_xticks(x_idx); ax3.set_xticklabels(ag3['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=3, frameon=True)
        st.pyplot(fig3)
    else: st.warning("⚠️ Sin datos para GAP.")

with col4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Valorización económica de la ineficiencia</i></div>", unsafe_allow_html=True)
    
    if not df_ef_f.empty:
        ag4 = df_ef_f.groupby('Fecha').agg({'HH_Improductivas': 'sum', 'Costo_Improd._$': 'sum'}).reset_index()
        fig4, ax4 = plt.subplots(figsize=(14, 10)); ax4_line = ax4.twinx()
        fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92)
        fig4.suptitle(t_enc, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray', fontweight='bold')

        x_idx = np.arange(len(ag4))
        
        bar_imp = ax4.bar(x_idx, ag4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMPRODUCTIVAS', zorder=2)
        ax4.bar_label(bar_imp, padding=4, color='black', fontweight='bold', path_effects=efecto_b, zorder=4)
        
        set_escala_y(ax4, ag4['HH_Improductivas'].max(), 1.4) 
        ax4_line.plot(x_idx, ag4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=efecto_b, label='COSTO ARS', zorder=5)
        add_tendencia(ax4_line, x_idx, ag4['Costo_Improd._$'])
        
        ax4_line.set_ylim(0, max(1000, ag4['Costo_Improd._$'].max() * 1.3)) 
        ax4_line.set_yticklabels([f'${int(x/1000000)}M' for x in ax4_line.get_yticks()], fontweight='bold')

        tot_pesos = ag4['Costo_Improd._$'].sum(); tot_hh_imp = ag4['HH_Improductivas'].sum()
        ax4.text(0.5, 0.90, f"COSTO TOTAL ACUMULADO ARS\n${tot_pesos:,.0f}\nTOTAL: {tot_hh_imp:,.0f} HH IMP", transform=ax4.transAxes, ha='center', va='top', fontsize=18, color='black', bbox=caja_o, weight='bold', zorder=10)

        for i, val in enumerate(ag4['Costo_Improd._$']):
            ax4_line.annotate(f"${val:,.0f}", (x_idx[i], val + 5), color='white', bbox=caja_g, ha='center', fontweight='bold', zorder=10)

        ax4.set_xticks(x_idx); ax4.set_xticklabels(ag4['Fecha'].dt.strftime('%b-%y'), fontsize=14, fontweight='bold')
        ax4.legend(loc='lower left', bbox_to_anchor=(0, 1.02), ncol=2, frameon=True); ax4_line.legend(loc='lower right', bbox_to_anchor=(1, 1.02), frameon=True)
        st.pyplot(fig4)
    else: st.warning("⚠️ No hay datos económicos.")

st.markdown("---")

# =========================================================================
# 11. FILA 3: MÉTRICAS 5 Y 6 (PARETO E INCIDENCIA)
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
        bar_p = ax5.bar(x_idx, ag5['Prom_M'], color='maroon', edgecolor='white', zorder=2)
        set_escala_y(ax5, ag5['Prom_M'].max(), 2.8)
        ax5.bar_label(bar_p, padding=4, color='black', fontweight='bold', fmt='%.1f', zorder=4)
        
        ax5_line.plot(x_idx, ag5['Pct_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=efecto_b, zorder=5)
        ax5_line.axhline(80, color='gray', linestyle='--', linewidth=2, zorder=1)
        ax5_line.set_ylim(0, 110); ax5_line.yaxis.set_major_formatter(mtick.PercentFormatter()) 

        lbls = [textwrap.fill(str(t), 12) for t in ag5['TIPO_PARADA']]
        ax5.set_xticks(x_idx); ax5.set_xticklabels(lbls, rotation=90, fontsize=12, fontweight='bold')
        
        for i, val in enumerate(ag5['Pct_Acu']): ax5_line.annotate(f"{val:.1f}%", (x_idx[i], val + 4), color='white', bbox=caja_g, ha='center', va='bottom', fontsize=11, rotation=45, zorder=10)
        sum_m = ag5['Prom_M'].sum()
        ax5.text(0.02, 0.96, f"SUMA PROMEDIO MENSUAL\n{sum_m:.1f} HH", transform=ax5.transAxes, bbox=caja_g, color='white', fontsize=15, ha='left', va='top', zorder=10)
        st.pyplot(fig5)
        
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        df_tbl = ag5.copy(); t_hs = df_tbl['HH_IMPRODUCTIVAS'].sum()
        df_tbl['% sobre Selección'] = (df_tbl['HH_IMPRODUCTIVAS'] / t_hs) * 100
        
        f_tot = pd.DataFrame({'TIPO_PARADA': ['✅ TOTAL'], 'HH_IMPRODUCTIVAS': [t_hs], 'Prom_M': [df_tbl['Prom_M'].sum()], 'Pct_Acu': [100.0], '% sobre Selección': [100.0]})
        df_tbl = pd.concat([df_tbl, f_tot], ignore_index=True)
        st.dataframe(df_tbl.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA': 'Motivo'}), use_container_width=True, hide_index=True, column_config={"Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), "% sobre Selección": st.column_config.NumberColumn(format="%.1f %%")})
    else: st.success("✅ Cero horas improductivas.")

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
        else: df6 = ag_disp.copy(); list_c = []
            
        df6['Suma_I'] = df6[list_c].sum(axis=1) if list_c else 0
        df6['Inc_%'] = (df6['Suma_I'] / df6['HH_Disponibles'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        df6['Fecha_O'] = pd.to_datetime(df6['K_Mes'] + '-01'); df6 = df6.sort_values(by='Fecha_O')

        fig6, ax6 = plt.subplots(figsize=(14, 10))
        fig6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92) 
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
            v_i = df6['Suma_I'].iloc[i]; v_d = df6['HH_Disponibles'].iloc[i]
            if v_i > 0: ax6.annotate(f"Imp: {int(v_i)}\nDisp: {int(v_d)}", (i, v_i + (ax6.get_ylim()[1]*0.02)), ha='center', bbox=caja_o, fontweight='bold', zorder=10)

        ax6_line = ax6.twinx()
        ax6_line.plot(x_idx, df6['Inc_%'], color='red', marker='o', markersize=12, linewidth=6, path_effects=efecto_b, label='% Incidencia', zorder=5)
        add_tendencia(ax6_line, x_idx, df6['Inc_%'])
        
        ax6_line.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1)
        ax6_line.text(x_idx[0], 16, 'META = 15%', color='white', bbox=caja_v, fontsize=14, fontweight='bold', zorder=10)
        
        for i, val in enumerate(df6['Inc_%']): ax6_line.annotate(f"{val:.1f}%", (x_idx[i], val + 2), color='red', ha='center', fontsize=16, fontweight='bold', path_effects=efecto_b, zorder=10)

        ax6.set_xticks(x_idx); ax6.set_xticklabels(df6['K_Mes'], fontsize=14, fontweight='bold')
        ax6_line.set_ylim(0, max(30, df6['Inc_%'].max() * 1.5))
        ax6_line.legend(loc='upper right', bbox_to_anchor=(1, 1.02), frameon=True)
        st.pyplot(fig6)
    else: st.warning("⚠️ Sin datos históricos.")

st.markdown("---")

# =========================================================================
# 12. FILA 4: ANÁLISIS PROFUNDO DE CAUSAS (DETALLES)
# =========================================================================
st.header("7. DETALLES DE IMPRODUCTIVIDAD (MESA DE TRABAJO)")
st.markdown("<div style='min-height:25px; font-size:14px; color:#aaa;'><i>Apertura de los registros detallados correspondientes a los filtros superiores</i></div>", unsafe_allow_html=True)

if not df_im_f.empty and 'DETALLE' in df_im_f.columns:
    motivos_disponibles = sorted(df_im_f['TIPO_PARADA'].dropna().unique())
    
    col_sel_motivo, col_espacio = st.columns([1, 2])
    with col_sel_motivo:
        motivo_seleccionado = st.selectbox("🔍 Seleccionar Motivo a detallar:", ["Todos los motivos"] + list(motivos_disponibles))
    
    df_detalles = df_im_f.copy()
    if motivo_seleccionado != "Todos los motivos":
        df_detalles = df_detalles[df_detalles['TIPO_PARADA'] == motivo_seleccionado]
        
    if not df_detalles.empty:
        agrup_detalles = df_detalles.groupby('DETALLE').agg({'HH_IMPRODUCTIVAS': 'sum'}).reset_index()
        agrup_detalles = agrup_detalles.sort_values(by='HH_IMPRODUCTIVAS', ascending=False)
        
        gran_total_detalles = agrup_detalles['HH_IMPRODUCTIVAS'].sum()
        agrup_detalles['% sobre Motivo'] = (agrup_detalles['HH_IMPRODUCTIVAS'] / gran_total_detalles) * 100
        
        fila_tot_detalles = pd.DataFrame({'DETALLE': ['✅ TOTAL SUMATORIA'], 'HH_IMPRODUCTIVAS': [gran_total_detalles], '% sobre Motivo': [100.0]})
        agrup_detalles = pd.concat([agrup_detalles, fila_tot_detalles], ignore_index=True)
        
        st.dataframe(
            agrup_detalles.rename(columns={'HH_IMPRODUCTIVAS': 'Subtotal HH', 'DETALLE': 'Detalle Registrado'}), 
            use_container_width=True, 
            hide_index=True, 
            column_config={
                "Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), 
                "% sobre Motivo": st.column_config.NumberColumn(format="%.1f %%")
            }
        )
        
        # Botón para descargar solo la tabla de detalles filtrada
        csv_detalles = agrup_detalles.to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Descargar Detalle Operativo (CSV)", data=csv_detalles, file_name="Detalles_Improductividad.csv", mime="text/csv", use_container_width=True)
    else:
        st.info("No hay registros detallados para el motivo seleccionado en este periodo.")
else:
    st.info("La base de datos de Improductivas no posee la columna 'Detalle', o no hay horas improductivas reportadas con la configuración de filtros actual.")
