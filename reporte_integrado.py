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
css_styles = """
<style>
/* Forzar que todos los contenedores padres permitan elementos sticky (Inmóviles) */
.stApp, .main, .block-container, [data-testid="stAppViewBlockContainer"], [data-testid="stMain"] {
    overflow: visible !important;
}

/* Seleccionar el contenedor de los 4 filtros y fijarlo */
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

if "admin" not in st.query_params:
    css_styles += """
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    """
else:
    css_styles += """
    div[data-testid="stHorizontalBlock"]:nth-of-type(2) {
        top: 55px !important; 
    }
    """

css_styles += "</style>"
st.markdown(css_styles, unsafe_allow_html=True)

# Regla Innegociable: Tamaños de fuente grandes y en negrita (AUMENTADOS)
plt.rcParams.update({
    'font.size': 14, 
    'font.weight': 'bold',
    'axes.labelweight': 'bold',
    'axes.titleweight': 'bold',
    'figure.titlesize': 18, 
    'figure.titleweight': 'bold'
})

# Efectos de contorno (Anti-Overlap de texto) y bboxes
outline_white = [path_effects.withStroke(linewidth=3, foreground='white')]
outline_black = [path_effects.withStroke(linewidth=3, foreground='black')]
bbox_gray = dict(boxstyle="round,pad=0.3", fc="dimgray", ec="white", lw=1.5)
bbox_yellow = dict(boxstyle="round,pad=0.4", fc="gold", ec="black", lw=1.
