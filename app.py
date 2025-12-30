import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import process_all

st.set_page_config(page_title="Airport Dashboard", layout="wide")
st.title("📊 Gestión y Productividad - CL Airport")

st.sidebar.header("📁 Carga de Archivos")
turnos_file = st.sidebar.file_uploader("Subir Turnos (CSV/XLSX)", type=['csv', 'xlsx'])
ventas_file = st.sidebar.file_uploader("Subir Ventas (CSV/XLSX)", type=['csv', 'xlsx'])

st.sidebar.header("📅 Rango de Fechas")
d_start = st.sidebar.date_input("Inicio", date(2025, 11, 1))
d_end = st.sidebar.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Procesar Reportes"):
    if turnos_file and ventas_file:
        try:
            # Guardar archivos temporales
            with open("temp_turnos.csv", "wb") as f: f.write(turnos_file.getbuffer())
            with open("temp_ventas.csv", "wb") as f: f.write(ventas_file.getbuffer())
            
            df_hourly, df_daily, df_total = process_all("temp_ventas.csv", "temp_turnos.csv", d_start, d_end)
            st.success("✅ Datos procesados correctamente.")
            
            tab1, tab2, tab3 = st.tabs(["⏰ Matriz Horaria", "📅 Resumen Diario", "👤 Resumen Total"])
            
            with tab1:
                st.subheader("Mapa de Cobertura por Franja Horaria")
                st.dataframe(df_hourly, use_container_width=True)
                
            with tab2:
                st.subheader("Productividad Diaria por Coordinador")
                st.dataframe(df_daily, use_container_width=True)
                
            with tab3:
                st.subheader("Totales Acumulados del Periodo")
                st.dataframe(df_total, use_container_width=True)
                
            # Excel download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_hourly.to_excel(writer, sheet_name='Matriz Horaria', index=False)
                df_daily.to_excel(writer, sheet_name='Resumen Diario', index=False)
                df_total.to_excel(writer, sheet_name='Resumen Total', index=False)
            
            st.download_button(
                label="📥 Descargar Reporte Completo (.xlsx)",
                data=output.getvalue(),
                file_name=f"Reporte_Airport_{d_start}_{d_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Error al procesar: {e}")
    else:
        st.error("⚠️ Sube ambos archivos para continuar.")
