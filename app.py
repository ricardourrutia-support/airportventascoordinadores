import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import procesar_final_airport

st.set_page_config(page_title="Gestión Airport", layout="wide")
st.title("📊 Control de Ventas y Cobertura Fija")

with st.sidebar:
    t_file = st.file_uploader("Subir Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Subir Ventas", type=['xlsx', 'csv'])
    d_ini = st.date_input("Inicio", date(2025, 11, 1))
    d_fin = st.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Procesar"):
    if t_file and v_file:
        try:
            df_m, df_na_h, df_na_d, df_p = procesar_final_airport(v_file, t_file, d_ini, d_fin)
            
            tab1, tab2, tab3 = st.tabs(["⏰ Matriz Coordinadores", "⚠️ No Asignados", "🏆 Resumen Periodo"])
            
            with tab1:
                st.subheader("Mapa de Cobertura y Ventas (Casilleros Fijos)")
                st.dataframe(df_m)
                
            with tab2:
                st.subheader("Reporte de Ventas sin Coordinador")
                col1, col2 = st.columns(2)
                col1.write("**Resumen Diario**")
                col1.dataframe(df_na_d)
                col2.metric("Total No Asignado", f"${df_na_d['Venta No Asignada'].sum():,.0f}")
                st.write("**Detalle por Franja Horaria**")
                st.dataframe(df_na_h)

            with tab3:
                st.subheader("Totales por Coordinador")
                st.table(df_p)

            # Descarga Excel con todas las pestañas
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_m.to_excel(writer, sheet_name='Matriz_Coordinadores', index=False)
                df_na_h.to_excel(writer, sheet_name='No_Asignados_Horario', index=False)
                df_na_d.to_excel(writer, sheet_name='No_Asignados_Diario', index=False)
                df_p.to_excel(writer, sheet_name='Totales_Periodo', index=False)
            
            st.download_button("📥 Descargar Reporte Completo", output.getvalue(), "Reporte_Airport.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
