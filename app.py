import streamlit as st
import pandas as pd
from datetime import date
from processor import procesar_v2_fijo

st.set_page_config(page_title="Gestión de Turnos Fijos", layout="wide")
st.title("📊 Mapa de Cobertura con Coordinadores Fijos")

t_file = st.sidebar.file_uploader("Subir Turnos", type=['xlsx', 'csv'])
v_file = st.sidebar.file_uploader("Subir Ventas", type=['xlsx', 'csv'])
d_ini = st.sidebar.date_input("Inicio", date(2025, 11, 1))
d_fin = st.sidebar.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("Procesar Reporte"):
    if t_file and v_file:
        try:
            df_final, lista_nombres = procesar_v2_fijo(v_file, t_file, d_ini, d_fin)
            st.dataframe(df_final, use_container_width=True)
        except Exception as e:
            st.error(f"Error al procesar: {e}")
