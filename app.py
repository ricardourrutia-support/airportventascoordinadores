import streamlit as st
import pandas as pd
from datetime import date
from processor import process_all, generate_styled_excel

st.set_page_config(page_title="Airport Pro", layout="wide")
st.title("📊 Reporte de Coordinadores - Estilo Cabify")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Cabify_Logo.svg/1200px-Cabify_Logo.svg.png", width=150) # Logo opcional si quieres
    st.header("Carga de Datos")
    t_file = st.file_uploader("Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Ventas", type=['xlsx', 'csv'])
    st.divider()
    d_ini = st.date_input("Desde", date(2025, 11, 1))
    d_fin = st.date_input("Hasta", date(2025, 11, 30))

if st.sidebar.button("Generar Reporte Oficial"):
    if t_file and v_file:
        try:
            # Procesar
            df_h, df_d, df_t, df_s = process_all(v_file, t_file, d_ini, d_fin)
            
            st.success("✅ Datos procesados correctamente.")

            # Mostrar Previsualización
            tab1, tab2, tab3 = st.tabs(["Matriz Horaria", "Resumen Diario", "Indicadores Clave"])
            
            with tab1:
                st.dataframe(df_h, height=400)
            with tab2:
                st.dataframe(df_d)
            with tab3:
                col1, col2 = st.columns(2)
                col1.write("##### Totales y Turnos")
                col1.table(df_t)
                col2.write("##### Análisis de Franjas Compartidas")
                col2.table(df_s)

            # Generar Excel Estilizado
            excel_data = generate_styled_excel({
                'Matriz_Horaria': df_h,
                'Resumen_Diario': df_d,
                'Totales_Periodo': df_t,
                'Franjas_Compartidas': df_s
            })
            
            st.download_button(
                label="📥 Descargar Reporte Estilo Cabify (.xlsx)",
                data=excel_data,
                file_name=f"Reporte_Cabify_{d_ini}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
    else:
        st.warning("Por favor carga ambos archivos para continuar.")
