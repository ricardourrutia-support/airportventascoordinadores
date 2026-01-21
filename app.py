import streamlit as st
import pandas as pd
from datetime import date
from processor import process_all, generate_styled_excel

st.set_page_config(page_title="Reporte Cabify", layout="wide")

# Estilos CSS
st.markdown("""
<style>
    .main-header {font-family: 'Arial'; color: #7145D6; font-size: 32px; font-weight: bold;}
    .sub-header {font-family: 'Arial'; color: #333333; font-size: 18px;}
    .info-box {background-color: #F0EBFF; padding: 15px; border-radius: 10px; border-left: 5px solid #7145D6;}
    /* Optimización de tablas */
    .stDataFrame {font-size: 12px;}
</style>
<div class='main-header'>Reporte de Productividad Airport</div>
<div class='sub-header'>Gestión de Coordinadores, Ventas y Horarios Especiales</div>
<hr style='border: 1px solid #7145D6;'>
""", unsafe_allow_html=True)

# Cacheamos la función pesada para que no se ejecute al cambiar de tab
@st.cache_data(show_spinner=False)
def run_processing(v_file, t_file, d_ini, d_fin):
    return process_all(v_file, t_file, d_ini, d_fin)

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Cabify_Logo.svg/1200px-Cabify_Logo.svg.png", width=120)
    st.header("Configuración")
    t_file = st.file_uploader("1. Archivo Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("2. Archivo Ventas", type=['xlsx', 'csv'])
    st.divider()
    d_ini = st.date_input("Inicio", date(2025, 12, 1))
    d_fin = st.date_input("Fin", date(2025, 12, 31))
    
    st.info("""
    **Reglas de Loza/Colación:**
    • Turno 10:00 -> OFF: 10-11 y 14-16
    • Turno 05:00 -> OFF: 11-14
    • Turno 21:00 -> OFF: 06-09
    """)

if st.button("Generar Reporte Oficial"):
    if t_file and v_file:
        try:
            with st.spinner("Procesando datos... (Esto será rápido)"):
                # Llamada a la función con caché
                res = run_processing(v_file, t_file, d_ini, d_fin)
            
            if res is None or res[0] is None:
                st.error("Error: Archivo de ventas inválido o sin fechas.")
            else:
                df_h, df_d, df_t, df_s = res
                
                st.success("✅ Procesamiento completado.")
                
                st.markdown("""
                <div class='info-box'>
                <b>Nota:</b> Coordinadores con <b>(*)</b> están en horario de Loza/Colación y no reciben ventas en esa hora.
                </div>
                """, unsafe_allow_html=True)
                
                tab1, tab2, tab3 = st.tabs(["Matriz Horaria", "Totales y Turnos", "Franjas Compartidas"])
                
                with tab1:
                    st.dataframe(df_h, use_container_width=True, height=500)
                with tab2:
                    c1, c2 = st.columns([2, 1])
                    c1.write("##### Resumen Diario")
                    c1.dataframe(df_d, use_container_width=True)
                    c2.write("##### Indicadores Globales")
                    c2.dataframe(df_t, use_container_width=True)
                with tab3:
                    st.write("##### Análisis de Competencia (Horas Hábiles de Venta)")
                    st.dataframe(df_s, use_container_width=True)

                # Generar Excel (este no se cachea porque retorna bytes)
                excel_bytes = generate_styled_excel({
                    'Matriz_Horaria': df_h,
                    'Resumen_Diario': df_d,
                    'Totales_Periodo': df_t,
                    'Franjas_Compartidas': df_s
                })
                
                st.download_button(
                    label="📥 Descargar Reporte Estilo Cabify (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Reporte_Cabify_{d_ini}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error técnico: {e}")
    else:
        st.warning("Carga ambos archivos para continuar.")
