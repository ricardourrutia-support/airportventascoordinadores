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
            t_path, v_path = "temp_turnos.csv", "temp_ventas.csv"
            with open(t_path, "wb") as f: f.write(turnos_file.getbuffer())
            with open(v_path, "wb") as f: f.write(ventas_file.getbuffer())
            
            df_hourly, df_daily, df_total, df_shared = process_all(v_path, t_path, d_start, d_end)
            st.success("✅ Procesamiento completado.")
            
            tab1, tab2, tab3, tab4 = st.tabs(["⏰ Matriz Horaria", "📅 Resumen Diario", "👤 Resumen Total", "🤝 Franjas Compartidas"])
            with tab1: st.dataframe(df_hourly)
            with tab2: st.dataframe(df_daily)
            with tab3: st.dataframe(df_total)
            with tab4: st.dataframe(df_shared)
                
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_hourly.to_excel(writer, sheet_name='Matriz Horaria', index=False)
                df_daily.to_excel(writer, sheet_name='Resumen Diario', index=False)
                df_total.to_excel(writer, sheet_name='Resumen Total', index=False)
                df_shared.to_excel(writer, sheet_name='Franjas Compartidas', index=False)
            
            st.download_button(label="📥 Descargar Reporte Consolidado (.xlsx)", data=output.getvalue(), file_name="Reporte_Ventas_Coordinadores.xlsx")
        except Exception as e:
            st.error(f"❌ Error: {e}")
