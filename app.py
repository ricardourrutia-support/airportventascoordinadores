import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
from processor import process_all

st.set_page_config(page_title="Airport Coordinators Dashboard", layout="wide")

st.title("📊 Gestión y Productividad de Coordinadores - CL Airport")

st.sidebar.header("📁 Carga de Archivos")
turnos_file = st.sidebar.file_uploader("Base de Turnos (CSV/XLSX)", type=['csv', 'xlsx'])
ventas_file = st.sidebar.file_uploader("Base de Ventas (CSV/XLSX)", type=['csv', 'xlsx'])

st.sidebar.header("📅 Filtro de Fechas")
d_start = st.sidebar.date_input("Fecha Inicio", date(2025, 11, 1))
d_end = st.sidebar.date_input("Fecha Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Procesar Reportes"):
    if turnos_file and ventas_file:
        try:
            # Guardar temporalmente para procesar
            t_ext = ".xlsx" if turnos_file.name.endswith('.xlsx') else ".csv"
            v_ext = ".xlsx" if ventas_file.name.endswith('.xlsx') else ".csv"
            t_path = "temp_turnos" + t_ext
            v_path = "temp_ventas" + v_ext
            
            with open(t_path, "wb") as f: f.write(turnos_file.getbuffer())
            with open(v_path, "wb") as f: f.write(ventas_file.getbuffer())
            
            # Llamada al procesador
            result = process_all(v_path, t_path, d_start, d_end)
            
            if result[0] is None:
                st.error("❌ Error: No se encontró la columna 'createdAt_local' o 'date' en el archivo de ventas.")
            else:
                df_hourly, df_daily, df_total, df_shared = result
                
                st.success("✅ Procesamiento completado con éxito.")
                
                tab1, tab2, tab3, tab4 = st.tabs(["⏰ Matriz Horaria", "📅 Resumen Diario", "👤 Resumen Total", "🤝 Franjas Compartidas"])
                
                with tab1:
                    st.subheader("Matriz de Cobertura y Ventas por Franja Horaria")
                    st.dataframe(df_hourly)
                    
                with tab2:
                    st.subheader("Ventas Totales por Día")
                    st.dataframe(df_daily)
                    
                with tab3:
                    st.subheader("Resumen Total (Ventas + Turnos)")
                    st.dataframe(df_total)

                with tab4:
                    st.subheader("Análisis de Competencia en Turnos")
                    st.write("Cantidad de franjas horarias trabajadas según nivel de concurrencia.")
                    st.dataframe(df_shared)
                    
                # Descarga Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_hourly.to_excel(writer, sheet_name='Matriz Horaria', index=False)
                    df_daily.to_excel(writer, sheet_name='Resumen Diario', index=False)
                    df_total.to_excel(writer, sheet_name='Resumen Total', index=False)
                    df_shared.to_excel(writer, sheet_name='Franjas Compartidas', index=False)
                
                st.download_button(
                    label="📥 Descargar Reporte Consolidado (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"Reporte_Coordinadores_{d_start}_{d_end}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
        except Exception as e:
            st.error(f"❌ Error crítico: {e}")
    else:
        st.error("⚠️ Por favor cargue ambos archivos.")
