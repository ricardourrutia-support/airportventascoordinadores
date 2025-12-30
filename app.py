import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import procesar_final_airport

st.set_page_config(page_title="Gestión Airport Pro", layout="wide")
st.title("📊 Reporte Consolidado de Ventas Airport")

with st.sidebar:
    t_file = st.file_uploader("Subir Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Subir Ventas", type=['xlsx', 'csv'])
    d_ini = st.date_input("Inicio", date(2025, 11, 1))
    d_fin = st.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Generar Reportes"):
    if t_file and v_file:
        try:
            df_matriz, df_diario, df_periodo = procesar_final_airport(v_file, t_file, d_ini, d_fin)
            
            tab1, tab2, tab3 = st.tabs(["⏰ Matriz Horaria", "📅 Resumen Diario", "🏆 Totales Periodo"])
            
            with tab1:
                st.dataframe(df_matriz, use_container_width=True)
            with tab2:
                st.dataframe(df_diario, use_container_width=True)
            with tab3:
                st.table(df_periodo)

            # --- LÓGICA DE DESCARGA MULTI-PESTAÑA ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_matriz.to_excel(writer, sheet_name='Matriz_Horaria', index=False)
                df_diario.to_excel(writer, sheet_name='Resumen_Diario', index=False)
                df_periodo.to_excel(writer, sheet_name='Totales_Periodo', index=False)
            
            st.download_button(
                label="📥 Descargar Reporte Completo (Excel)",
                data=output.getvalue(),
                file_name=f"Reporte_Ventas_{d_ini}_al_{d_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
