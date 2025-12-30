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
        df_final, lista_nombres = procesar_v2_fijo(v_file, t_file, d_ini, d_fin)
        
        st.info("💡 Cada columna pertenece a un único coordinador. Si no aparece nombre, es porque no tenía turno.")
        
        # Mostrar leyenda de quién es quién
        cols = st.columns(len(lista_nombres))
        for i, nombre in enumerate(lista_nombres):
            cols[i].metric(f"Columna {i+1}", nombre)

        st.dataframe(df_final, use_container_width=True)
        
        # Excel
        import io
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button("Descargar Reporte", buf.getvalue(), "reporte_fijo.xlsx")
