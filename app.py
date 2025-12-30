import streamlit as st
import pandas as pd
from datetime import date
from processor import process_all

st.set_page_config(page_title="Dashboard Ventas Airport", layout="wide")
st.title("📊 Control de Ventas por Coordinador (Fijo)")

turnos_f = st.sidebar.file_uploader("Excel de Turnos", type=['xlsx', 'csv'])
ventas_f = st.sidebar.file_uploader("Excel de Ventas", type=['xlsx', 'csv'])
d_ini = st.sidebar.date_input("Inicio", date(2025, 11, 1))
d_fin = st.sidebar.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("Procesar"):
    if turnos_f and ventas_f:
        df_matriz, df_diario = process_all(ventas_f, turnos_f, d_ini, d_fin)
        
        tab1, tab2 = st.tabs(["⏰ Matriz Horaria (Juan Perez Fijo)", "📅 Resumen Diario"])
        with tab1:
            st.write("Debajo de cada Coordinador aparecerá el nombre solo si está operativo.")
            st.dataframe(df_matriz)
        with tab2:
            st.dataframe(df_diario)
    else:
        st.error("Sube los archivos.")
