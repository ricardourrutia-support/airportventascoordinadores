import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import process_all

st.set_page_config(page_title="Gestión CL Airport", layout="wide")
st.title("📊 Análisis de Ventas por Coordinador Fijo")

st.sidebar.header("📁 Carga de Datos")
turnos_file = st.sidebar.file_uploader("Subir Turnos", type=['csv', 'xlsx'])
ventas_file = st.sidebar.file_uploader("Subir Ventas", type=['csv', 'xlsx'])

st.sidebar.header("📅 Filtro de Fechas")
d_start = st.sidebar.date_input("Inicio", date(2025, 11, 1))
d_end = st.sidebar.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Generar Reportes"):
    if turnos_file and ventas_file:
        try:
            # Procesar archivos
            t_path = "temp_turnos.xlsx" if turnos_file.name.endswith('.xlsx') else "temp_turnos.csv"
            v_path = "temp_ventas.xlsx" if ventas_file.name.endswith('.xlsx') else "temp_ventas.csv"
            with open(t_path, "wb") as f: f.write(turnos_file.getbuffer())
            with open(v_path, "wb") as f: f.write(ventas_file.getbuffer())
            
            df_hourly, df_daily, df_total = process_all(v_path, t_path, d_start, d_end)
            
            st.success("✅ Reporte generado con éxito.")
            
            tab1, tab2, tab3 = st.tabs(["⏰ Matriz Horaria (C1-C6)", "📅 Ventas Diarias", "👤 Resumen Periodo"])
            
            with tab1:
                st.subheader("Mapa Horario de Cobertura y Ventas")
                st.dataframe(df_hourly)
                
            with tab2:
                st.subheader("Total Ventas por Día")
                st.dataframe(df_daily)
                
            with tab3:
                st.subheader("Resumen General del Periodo")
                st.dataframe(df_total)
                
            # Descarga Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_hourly.to_excel(writer, sheet_name='Matriz', index=False)
                df_daily.to_excel(writer, sheet_name='Diario', index=False)
                df_total.to_excel(writer, sheet_name='Resumen', index=False)
            
            st.download_button("📥 Descargar Reporte (.xlsx)", output.getvalue(), "Reporte_Airport.xlsx")
            
        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.error("⚠️ Sube ambos archivos.")
