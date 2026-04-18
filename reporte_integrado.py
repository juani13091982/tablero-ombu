import streamlit as st, pandas as pd, numpy as np, matplotlib.pyplot as plt, matplotlib.ticker as mtick, matplotlib.patheffects as pe, textwrap, re

st.set_page_config(page_title="Tablero Ombú", layout="wide", initial_sidebar_state="collapsed")

if 'auth' not in st.session_state: st.session_state['auth'] = False

def login():
    st.markdown("<br><br>", unsafe_allow_html=True); _, c, _ = st.columns([1, 1.8, 1])
    with c:
        st.markdown("<div style='background-color:#1E3A8A; padding:5px; border-radius:10px 10px 0px 0px;'></div>", unsafe_allow_html=True)
        try: st.image("LOGO OMBÚ.jpg", width=160)
        except: st.markdown("<h2 style='text-align:center;'>OMBÚ</h2>", unsafe_allow_html=True)
        st.markdown("<div style='text-align:center;'><h2 style='color:#1E3A8A;'>GESTIÓN INDUSTRIAL OMBÚ S.A.</h2><p>Acceso Restringido</p></div>", unsafe_allow_html=True)
        with st.form("l"):
            u, p = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if u=="acceso.ombu" and p=="Gestion2026": st.session_state['auth'] = True; st.rerun()
                else: st.error("❌ Credenciales incorrectas.")
if not st.session_state['auth']: login(); st.stop()

st.markdown("<style>#MainMenu, header, footer {visibility: hidden !important;} div[data-testid='stVerticalBlock'] > div:has(#filtro-ribbon) {position: sticky !important; top: 0px !important; background-color: #0E1117 !important; z-index: 99999 !important; padding: 15px; border-bottom: 3px solid #1E3A8A !important;}</style>", unsafe_allow_html=True)

plt.rcParams.update({'font.size': 14, 'font.weight': 'bold', 'axes.labelweight': 'bold', 'axes.titleweight': 'bold', 'figure.titlesize': 18})
cw, cb = [pe.withStroke(linewidth=3, foreground='white')], [pe.withStroke(linewidth=3, foreground='black')]
bx_v, bx_g, bx_o, bx_b = dict(boxstyle="round,pad=0.3", fc="darkgreen", ec="white", lw=1.5), dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5), dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.5), dict(boxstyle="round,pad=0.3", fc="white", ec="black", lw=1.5)

def fuzzy(s, v):
    if pd.isna(s) or pd.isna(v): return False
    s1, s2 = str(s).upper(), str(v).upper()
    for a,b in zip("ÁÉÍÓÚ","AEIOU"): s1, s2 = s1.replace(a,b), s2.replace(a,b)
    if s1 in s2 or s2 in s1: return True
    w1, w2 = set(re.findall(r'\w+', s1)), set(re.findall(r'\w+', s2))
    return bool((w1 & w2) - {'SECTOR','PUESTO','LINEA','PLANTA','DE','LA','EL'})

c_l, c_t, c_s = st.columns([1, 3, 1])
with c_l:
    try: st.image("LOGO OMBÚ.jpg", width=120)
    except: st.markdown("### OMBÚ")
with c_t: st.title("TABLERO INTEGRADO - REPORTE C.G.P.")
with c_s:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Salir", use_container_width=True): st.session_state['auth'] = False; st.rerun()

try:
    d_e, d_i = pd.read_excel("eficiencias.xlsx"), pd.read_excel("improductivas.xlsx")
    d_e.columns, d_i.columns = d_e.columns.str.strip(), [str(c).strip().upper() for c in d_i.columns]
    if 'TIPO_PARADA' not in d_i.columns: d_i.rename(columns={next(c for c in d_i.columns if 'TIPO' in c or 'MOTIVO' in c): 'TIPO_PARADA'}, inplace=True)
    if 'HH_IMPRODUCTIVAS' not in d_i.columns: d_i.rename(columns={next(c for c in d_i.columns if 'HH' in c and 'IMP' in c): 'HH_IMPRODUCTIVAS'}, inplace=True)
    if 'FECHA' not in d_i.columns: d_i.rename(columns={next(c for c in d_i.columns if 'FECHA' in c): 'FECHA'}, inplace=True)
    d_e['Fecha'], d_i['FECHA'] = pd.to_datetime(d_e['Fecha'], errors='coerce').dt.to_period('M').dt.to_timestamp(), pd.to_datetime(d_i['FECHA'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    d_e['Es_Ultimo_Puesto'], d_e['M'], d_i['M'] = d_e['Es_Ultimo_Puesto'].astype(str).str.strip().str.upper(), d_e['Fecha'].dt.strftime('%b-%Y'), d_i['FECHA'].dt.strftime('%b-%Y')
except Exception as e: st.error(f"Error: {e}"); st.stop()

st.markdown('<div id="filtro-ribbon"></div><h3>🔍 Escenario</h3>', unsafe_allow_html=True)
f1, f2, f3, f4 = st.columns(4)
s_pl = f1.multiselect("🏭 Planta", list(d_e['Planta'].dropna().unique()))
s_li = f2.multiselect("⚙️ Línea", list(d_e[d_e['Planta'].isin(s_pl)]['Linea'].dropna().unique() if s_pl else d_e['Linea'].dropna().unique()))
s_pu = f3.multiselect("🛠️ Puesto", list(d_e[d_e['Linea'].isin(s_li)]['Puesto_Trabajo'].dropna().unique() if s_li else d_e['Puesto_Trabajo'].dropna().unique()))
s_me = f4.multiselect("📅 Mes", ["🎯 Acumulado YTD"] + list(d_e['M'].unique()))

ef, im = d_e.copy(), d_i.copy()
if s_pl: ef = ef[ef['Planta'].isin(s_pl)]; im = im[im.iloc[:,0].apply(lambda x: any(fuzzy(p, x) for p in s_pl))]
if s_li: 
    ef = ef[ef['Linea'].isin(s_li)]
    cl = next((c for c in im.columns if 'LINEA' in c), im.columns[1])
    im = im[im[cl].apply(lambda x: any(fuzzy(l, x) for l in s_li))]
if s_pu: 
    ef = ef[ef['Puesto_Trabajo'].isin(s_pu)]
    cp = next((c for c in im.columns if 'PUESTO' in c), im.columns[2])
    im = im[im[cp].apply(lambda x: any(fuzzy(p, x) for p in s_pu))]
if s_me and "🎯 Acumulado YTD" not in s_me: ef, im = ef[ef['M'].isin(s_me)], im[im['M'].isin(s_me)]

tit = f"Filtros: {'+'.join(s_pl) if s_pl else 'Todas'} > {'+'.join(s_li) if s_li else 'Todas'} > {'+'.join(s_pu) if s_pu else 'Todos'}"
st.markdown("---")

c1, c2 = st.columns(2)
with c1:
    st.header("1. EFICIENCIA REAL")
    st.markdown("<p style='color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH DISPONIBLES)</i></p>", unsafe_allow_html=True)
    d1 = ef.copy() if s_pu else ef[ef['Es_Ultimo_Puesto'] == 'SI']
    if not d1.empty:
        a1 = d1.groupby('Fecha').agg({'HH_STD_TOTAL':'sum','HH_Disponibles':'sum','Cant._Prod._A1':'sum'}).reset_index()
        a1['Ef'] = (a1['HH_STD_TOTAL']/a1['HH_Disponibles']).replace([np.inf,-np.inf],0).fillna(0)*100
        f1, ax = plt.subplots(figsize=(14,10)); ax2=ax.twinx(); f1.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); f1.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(a1)); w=0.35
        b1=ax.bar(x-w/2, a1['HH_STD_TOTAL'], w, color='midnightblue', edgecolor='white', label='HH STD')
        b2=ax.bar(x+w/2, a1['HH_Disponibles'], w, color='black', edgecolor='white', label='HH DISP')
        ax.set_ylim(0, a1['HH_Disponibles'].max()*2.6 if not a1.empty else 100); ax.bar_label(b1, padding=4, color='black', path_effects=cw, fmt='%.0f'); ax.bar_label(b2, padding=4, color='black', path_effects=cw, fmt='%.0f')
        for i in range(len(x)): ax.axvline(i, color='lightgray', linestyle='--')
        for i, b in enumerate(b1):
            if a1['Cant._Prod._A1'].iloc[i] > 0: ax.text(b.get_x()+b.get_width()/2, b.get_height()*0.05, f"{int(a1['Cant._Prod._A1'].iloc[i])} UND", rotation=90, color='white', ha='center', va='bottom', fontsize=18, path_effects=cb)
        ax2.plot(x, a1['Ef'], color='dimgray', marker='o', markersize=12, linewidth=4, path_effects=cw, label='% Efic. Real')
        ax2.axhline(85, color='darkgreen', linestyle='--', linewidth=3); ax2.text(x[0], 86, 'META = 85%', color='white', bbox=bx_v)
        ax2.set_ylim(0, max(120, a1['Ef'].max()*1.8)); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, v in enumerate(a1['Ef']): ax2.annotate(f"{v:.1f}%", (x[i], v+5), color='white', bbox=bx_g, ha='center')
        ax.set_xticks(x); ax.set_xticklabels(a1['Fecha'].dt.strftime('%b-%y')); ax.legend(loc='lower left', bbox_to_anchor=(0,1.02), ncol=2); ax2.legend(loc='lower right', bbox_to_anchor=(1,1.02))
        st.pyplot(f1)
    else: st.warning("⚠️ Sin datos.")

with c2:
    st.header("2. EFICIENCIA PRODUCTIVA")
    st.markdown("<p style='color:#aaa;'><i>Fórmula: (∑ HH STD / ∑ HH PRODUCTIVAS)</i></p>", unsafe_allow_html=True)
    if not d1.empty:
        a2 = d1.groupby('Fecha').agg({'HH_STD_TOTAL':'sum','HH_Productivas_C/GAP':'sum'}).reset_index()
        a2['Ef'] = (a2['HH_STD_TOTAL']/a2['HH_Productivas_C/GAP']).replace([np.inf,-np.inf],0).fillna(0)*100
        f2, ax = plt.subplots(figsize=(14,10)); ax2=ax.twinx(); f2.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); f2.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(a2))
        b1=ax.bar(x-0.17, a2['HH_STD_TOTAL'], 0.35, color='midnightblue', edgecolor='white', label='HH STD')
        b2=ax.bar(x+0.17, a2['HH_Productivas_C/GAP'], 0.35, color='darkgreen', edgecolor='white', label='HH PROD')
        ax.set_ylim(0, max(a2['HH_STD_TOTAL'].max(), a2['HH_Productivas_C/GAP'].max())*2.6 if not a2.empty else 100)
        ax.bar_label(b1, padding=4, color='black', path_effects=cw, fmt='%.0f'); ax.bar_label(b2, padding=4, color='black', path_effects=cw, fmt='%.0f')
        for i in range(len(x)): ax.axvline(i, color='lightgray', linestyle='--')
        ax2.plot(x, a2['Ef'], color='dimgray', marker='s', markersize=12, linewidth=4, path_effects=cw, label='% Efic. Prod.')
        ax2.axhline(100, color='darkgreen', linestyle='--', linewidth=3); ax2.text(x[0], 101, 'META = 100%', color='white', bbox=bx_v)
        ax2.set_ylim(0, max(150, a2['Ef'].max()*1.8)); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        for i, v in enumerate(a2['Ef']): ax2.annotate(f"{v:.1f}%", (x[i], v+5), color='white', bbox=bx_g, ha='center')
        ax.set_xticks(x); ax.set_xticklabels(a2['Fecha'].dt.strftime('%b-%y')); ax.legend(loc='lower left', bbox_to_anchor=(0,1.02), ncol=2)
        st.pyplot(f2)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")
c3, c4 = st.columns(2)
with c3:
    st.header("3. GAP HH GLOBAL")
    st.markdown("<p style='color:#aaa;'><i>Desvío Horas Disponibles vs Declaradas</i></p>", unsafe_allow_html=True)
    if not ef.empty:
        cp = 'HH_Productivas' if 'HH_Productivas' in ef.columns else 'HH Productivas'
        a3 = ef.groupby('Fecha').agg({cp:'sum','HH_Improductivas':'sum','HH_Disponibles':'sum'}).reset_index()
        a3['Tot'] = a3[cp] + a3['HH_Improductivas']
        f3, ax = plt.subplots(figsize=(14,10)); f3.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); f3.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(a3))
        ax.bar(x, a3[cp], color='darkgreen', edgecolor='white', label='PROD')
        ax.bar(x, a3['HH_Improductivas'], bottom=a3[cp], color='firebrick', edgecolor='white', label='IMP')
        ax.plot(x, a3['HH_Disponibles'], color='black', marker='D', markersize=12, linewidth=4, path_effects=cw, label='DISP')
        ax.set_ylim(0, a3['HH_Disponibles'].max()*2.6 if not a3.empty else 100)
        for i in range(len(x)):
            hd, ht = a3['HH_Disponibles'].iloc[i], a3['Tot'].iloc[i]
            ax.plot([i,i], [ht,hd], color='dimgray', linewidth=5, alpha=0.6)
            ax.annotate(f"GAP:\n{int(hd-ht)}", (i, ht+5), color='firebrick', bbox=bx_b, ha='center', va='bottom')
            ax.annotate(f"{int(hd)}", (i, hd+(ax.get_ylim()[1]*0.08)), color='black', bbox=bx_b, ha='center')
        ax.set_xticks(x); ax.set_xticklabels(a3['Fecha'].dt.strftime('%b-%y')); ax.legend(loc='lower left', bbox_to_anchor=(0,1.02), ncol=3)
        st.pyplot(f3)
    else: st.warning("⚠️ Sin datos.")

with c4:
    st.header("4. COSTOS IMPRODUCTIVOS")
    st.markdown("<p style='color:#aaa;'><i>Valorización de HH Perdidas</i></p>", unsafe_allow_html=True)
    if not ef.empty:
        a4 = ef.groupby('Fecha').agg({'HH_Improductivas':'sum','Costo_Improd._$':'sum'}).reset_index()
        f4, ax = plt.subplots(figsize=(14,10)); ax2=ax.twinx(); f4.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); f4.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(a4))
        bi = ax.bar(x, a4['HH_Improductivas'], color='darkred', edgecolor='white', label='HH IMP')
        ax.bar_label(bi, padding=4, color='black', path_effects=cw); ax.set_ylim(0, a4['HH_Improductivas'].max()*2.6 if not a4.empty else 100)
        ax2.plot(x, a4['Costo_Improd._$'], color='maroon', marker='s', markersize=12, linewidth=5, path_effects=cw, label='COSTO ARS')
        ax2.set_ylim(0, max(1000, a4['Costo_Improd._$'].max()*1.8)); ax2.set_yticklabels([f'${int(v/1000000)}M' for v in ax2.get_yticks()])
        ax.text(0.5, 0.90, f"COSTO TOTAL: ${a4['Costo_Improd._$'].sum():,.0f}", transform=ax.transAxes, ha='center', va='top', fontsize=18, bbox=bx_o)
        for i, v in enumerate(a4['Costo_Improd._$']): ax2.annotate(f"${v:,.0f}", (x[i], v+5), color='white', bbox=bx_g, ha='center')
        ax.set_xticks(x); ax.set_xticklabels(a4['Fecha'].dt.strftime('%b-%y')); ax.legend(loc='lower left', bbox_to_anchor=(0,1.02), ncol=2)
        st.pyplot(f4)
    else: st.warning("⚠️ Sin datos.")

st.markdown("---")
c5, c6 = st.columns(2)

with c5:
    st.header("5. PARETO DE CAUSAS")
    st.markdown("<p style='color:#aaa;'><i>Distribución de motivos (80/20)</i></p>", unsafe_allow_html=True)
    if not im.empty:
        a5 = im.groupby('TIPO_PARADA')['HH_IMPRODUCTIVAS'].sum().reset_index()
        a5['Prom_M'] = a5['HH_IMPRODUCTIVAS'] / (im['FECHA'].nunique() or 1)
        a5 = a5.sort_values(by='Prom_M', ascending=False)
        a5['%_Acu'] = (a5['Prom_M'].cumsum() / a5['Prom_M'].sum()) * 100
        f5, ax = plt.subplots(figsize=(14,10)); ax2=ax.twinx(); f5.subplots_adjust(top=0.86, bottom=0.28, left=0.08, right=0.92); f5.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(a5))
        bp = ax.bar(x, a5['Prom_M'], color='maroon', edgecolor='white')
        ax.set_ylim(0, a5['Prom_M'].max()*2.8 if not a5.empty else 100); ax.bar_label(bp, padding=4, color='black', fmt='%.1f')
        ax2.plot(x, a5['%_Acu'], color='red', marker='D', markersize=10, linewidth=4, path_effects=cw)
        ax2.axhline(80, color='gray', linestyle='--'); ax2.set_ylim(0, 200); ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
        ax.set_xticks(x); ax.set_xticklabels([textwrap.fill(str(t), 12) for t in a5['TIPO_PARADA']], rotation=90, fontsize=12)
        for i, v in enumerate(a5['%_Acu']): ax2.annotate(f"{v:.1f}%", (x[i], v+4), color='white', bbox=bx_g, ha='center', va='bottom', fontsize=11, rotation=45)
        ax.text(0.02, 0.96, f"SUMA PROMEDIO: {a5['Prom_M'].sum():.1f} HH", transform=ax.transAxes, bbox=bx_g, color='white', fontsize=15, ha='left', va='top')
        st.pyplot(f5)
        
        st.markdown("### 🛠️ Mesa de Trabajo e Impacto")
        dt = a5.copy(); th = dt['HH_IMPRODUCTIVAS'].sum()
        dt['%'] = (dt['HH_IMPRODUCTIVAS']/th)*100
        dt = pd.concat([dt, pd.DataFrame({'TIPO_PARADA':['✅ TOTAL'], 'HH_IMPRODUCTIVAS':[th], 'Prom_M':[dt['Prom_M'].sum()], '%_Acu':[100.0], '%':[100.0]})], ignore_index=True)
        st.dataframe(dt.rename(columns={'HH_IMPRODUCTIVAS':'Subtotal HH', 'TIPO_PARADA':'Causa'}), use_container_width=True, hide_index=True, column_config={"Subtotal HH": st.column_config.NumberColumn(format="%.1f ⏱️"), "%": st.column_config.NumberColumn(format="%.1f %%")})
        st.download_button("📥 Descargar Plan (CSV)", dt.to_csv(index=False).encode('utf-8'), "Plan_Ombu.csv", "text/csv")
    else: st.success("✅ Cero horas improductivas.")

with c6:
    st.header("6. EVOLUCIÓN INCIDENCIA %")
    st.markdown("<p style='color:#aaa;'><i>% HH Improductivas sobre Disponibles</i></p>", unsafe_allow_html=True)
    if not ef.empty:
        ef['K'] = ef['Fecha'].dt.strftime('%Y-%m')
        ad = ef.groupby('K', as_index=False)['HH_Disponibles'].sum()
        if not im.empty:
            im['K'] = im['FECHA'].dt.strftime('%Y-%m')
            pv = pd.pivot_table(im, values='HH_IMPRODUCTIVAS', index='K', columns='TIPO_PARADA', aggfunc='sum').fillna(0).reset_index()
            d6 = pd.merge(ad, pv, on='K', how='left').fillna(0)
            lc = [c for c in d6.columns if c not in ['HH_Disponibles', 'K']]
        else: d6 = ad.copy(); lc = []
        d6['Sum'] = d6[lc].sum(axis=1) if lc else 0
        d6['Inc'] = (d6['Sum'] / d6['HH_Disponibles'] * 100).replace([np.inf,-np.inf],0).fillna(0)
        d6['Ord'] = pd.to_datetime(d6['K']+'-01')
        d6 = d6.sort_values('Ord')
        
        f6, ax = plt.subplots(figsize=(14, 10)); ax2=ax.twinx(); f6.subplots_adjust(top=0.86, bottom=0.22, left=0.08, right=0.92); f6.suptitle(tit, x=0.08, y=0.98, ha='left', fontsize=8, color='dimgray')
        x = np.arange(len(d6))
        if lc:
            b_s = np.zeros(len(d6)); pal = plt.cm.tab20.colors
            for i, c in enumerate(lc):
                vc = d6[c].values
                ax.bar(x, vc, bottom=b_s, label=c, color=pal[i%20], edgecolor='white', zorder=2)
                b_s += vc
        else: ax.bar(x, np.zeros(len(d6)), color='white')
        ax.set_ylim(0, d6['Sum'].max()*2.2 if not d6.empty else 100)
        for i in range(len(x)):
            vi, vd = d6['Sum'].iloc[i], d6['HH_Disponibles'].iloc[i]
            if vi > 0: ax.annotate(f"Imp: {int(vi)}\nDisp: {int(vd)}", (i, vi+(ax.get_ylim()[1]*0.05)), ha='center', bbox=bx_o, zorder=10)
        
        ax2.plot(x, d6['Inc'], color='red', marker='o', markersize=12, linewidth=6, path_effects=cw, label='% Inc', zorder=5)
        ax2.axhline(15, color='darkgreen', linestyle='--', linewidth=3, zorder=1); ax2.text(x[0], 16, 'META = 15%', color='white', bbox=bx_v, zorder=10)
        for i, v in enumerate(d6['Inc']): ax2.annotate(f"{v:.1f}%", (x[i], v+2), color='red', ha='center', path_effects=cw, zorder=10)
        ax.set_xticks(x); ax.set_xticklabels(d6['K']); ax2.set_ylim(0, max(30, d6['Inc'].max()*1.8))
        st.pyplot(f6)
    else: st.warning("⚠️ Sin datos.")
