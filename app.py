import streamlit as st
import pandas as pd
from datetime import date
from processor import process_all, generate_styled_excel

st.set_page_config(page_title="Reporte Cabify", layout="wide")

# Cabecera Minimalista
st.markdown("""
<style>
    .main-header {font-family: 'Arial'; color: #7145D6; font-size: 32px; font-weight: bold;}
    .sub-header {font-family: 'Arial'; color: #333333; font-size: 18px;}
</style>
<div class='main-header'>Reporte de Productividad Airport</div>
<div class='sub-header'>Gestión de Coordinadores y Ventas</div>
<hr style='border: 1px solid #7145D6;'>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Configuración")
    t_file = st.file_uploader("1. Archivo Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("2. Archivo Ventas", type=['xlsx', 'csv'])
    st.divider()
    d_ini = st.date_input("Fecha Inicio", date(2025, 12, 1))
    d_fin = st.date_input("Fecha Fin", date(2025, 12, 31))

if st.button("Generar Informe"):
    if t_file and v_file:
        try:
            res = process_all(v_file, t_file, d_ini, d_fin)
            
            if res[0] is None:
                st.error("❌ Error: No se encontró la columna 'createdAt_local' o 'date' en el archivo de Ventas.")
            else:
                df_h, df_d, df_t, df_s = res
                
                st.success("Reporte generado exitosamente.")
                
                tab1, tab2, tab3 = st.tabs(["Matriz Horaria", "Totales y Turnos", "Franjas Compartidas"])
                
                with tab1:
                    st.dataframe(df_h, use_container_width=True)
                with tab2:
                    c1, c2 = st.columns([2, 1])
                    c1.write("##### Resumen Diario")
                    c1.dataframe(df_d, use_container_width=True)
                    c2.write("##### Métricas Globales")
                    c2.dataframe(df_t, use_container_width=True)
                with tab3:
                    st.write("##### Análisis de Competencia (Horas)")
                    st.dataframe(df_s, use_container_width=True)

                # Generar Excel Estilizado
                excel_bytes = generate_styled_excel({
                    'Matriz_Horaria': df_h,
                    'Resumen_Diario': df_d,
                    'Totales_Periodo': df_t,
                    'Franjas_Compartidas': df_s
                })
                
                st.download_button(
                    label="📥 Descargar Excel Estilo Cabify",
                    data=excel_bytes,
                    file_name=f"Reporte_Cabify_{d_ini}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error inesperado: {e}")
    else:
        st.warning("Carga ambos archivos para continuar.")
