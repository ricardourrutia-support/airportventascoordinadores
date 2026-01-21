import streamlit as st
import pandas as pd
from datetime import date
from processor import process_all, generate_styled_excel

st.set_page_config(page_title="Airport Pro", layout="wide")

st.markdown("""
<style>
    .header {color: #7145D6; font-size: 30px; font-weight: bold;}
    .info-box {background-color: #F0EBFF; padding: 15px; border-radius: 10px; border-left: 5px solid #7145D6;}
</style>
<div class='header'>Gestión de Coordinadores Airport</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Cabify_Logo.svg/1200px-Cabify_Logo.svg.png", width=120)
    st.header("Configuración")
    t_file = st.file_uploader("1. Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("2. Ventas", type=['xlsx', 'csv'])
    st.divider()
    d_ini = st.date_input("Inicio", date(2025, 12, 1))
    d_fin = st.date_input("Fin", date(2025, 12, 31))
    
    st.markdown("""
    <div style='font-size: 12px; color: grey; margin-top: 20px;'>
    <b>Reglas de Loza/Colación aplicadas:</b><br>
    • Turno 10:00 -> Off: 10-11, 14-16<br>
    • Turno 05:00 -> Off: 11-14<br>
    • Turno 21:00 -> Off: 06-09
    </div>
    """, unsafe_allow_html=True)

if st.button("Generar Reporte Oficial"):
    if t_file and v_file:
        try:
            res = process_all(v_file, t_file, d_ini, d_fin)
            
            if res[0] is None:
                st.error("Error: Archivo de ventas inválido o sin fechas.")
            else:
                df_h, df_d, df_t, df_s = res
                
                st.markdown(f"""
                <div class='info-box'>
                <b>Proceso Completado</b><br>
                Se han aplicado las exclusiones de horarios de gestión. 
                Los coordinadores en Loza/Colación aparecen marcados con <b>(*)</b> en la matriz horaria y no comisionan en esas horas.
                </div>
                """, unsafe_allow_html=True)
                
                tab1, tab2, tab3 = st.tabs(["Matriz Horaria", "Totales y Turnos", "Franjas Compartidas"])
                
                with tab1:
                    st.dataframe(df_h, use_container_width=True, height=500)
                with tab2:
                    c1, c2 = st.columns([1, 1])
                    c1.write("##### Resumen Diario")
                    c1.dataframe(df_d, use_container_width=True)
                    c2.write("##### Indicadores Globales")
                    c2.dataframe(df_t, use_container_width=True)
                with tab3:
                    st.write("##### Análisis de Competencia (Horas Hábiles de Venta)")
                    st.dataframe(df_s, use_container_width=True)

                excel_bytes = generate_styled_excel({
                    'Matriz_Horaria': df_h,
                    'Resumen_Diario': df_d,
                    'Totales_Periodo': df_t,
                    'Franjas_Compartidas': df_s
                })
                
                st.download_button("📥 Descargar Reporte Final (.xlsx)", excel_bytes, f"Reporte_Cabify_{d_ini}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error técnico: {e}")
    else:
        st.warning("Carga ambos archivos.")
