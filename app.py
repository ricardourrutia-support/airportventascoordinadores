import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import procesar_final_airport

st.set_page_config(page_title="Gestión Airport", layout="wide")
st.title("📊 Control de Ventas y Cobertura Fija")

with st.sidebar:
    st.header("1. Cargar Archivos")
    t_file = st.file_uploader("Subir Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Subir Ventas", type=['xlsx', 'csv'])
    st.header("2. Rango de Fechas")
    d_ini = st.date_input("Inicio", date(2025, 11, 1))
    d_fin = st.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Procesar Reporte"):
    if t_file and v_file:
        try:
            # Procesar recibiendo 5 dataframes (incluyendo el de franjas compartidas)
            df_m, df_na_h, df_na_d, df_p, df_shared = procesar_final_airport(v_file, t_file, d_ini, d_fin)
            
            tab1, tab2, tab3, tab4 = st.tabs(["⏰ Matriz Coordinadores", "⚠️ No Asignados", "🏆 Resumen Periodo", "🤝 Franjas Compartidas"])
            
            with tab1:
                st.subheader("Mapa Horario (Casilleros Fijos)")
                st.dataframe(df_m)
                
            with tab2:
                st.subheader("Ventas sin Coordinador en Turno")
                c1, c2 = st.columns(2)
                c1.write("**Resumen Diario**")
                c1.dataframe(df_na_d)
                c2.metric("Pérdida Total", f"${df_na_d['Venta No Asignada'].sum():,.0f}")
                st.write("**Detalle por Franja Horaria**")
                st.dataframe(df_na_h)

            with tab3:
                st.subheader("Liquidación y Turnos")
                st.write("Incluye total de ventas y cantidad de días trabajados.")
                st.table(df_p)
                
            with tab4:
                st.subheader("Análisis de Competencia en Turnos")
                st.write("Cantidad de horas (franjas) trabajadas según concurrencia.")
                st.table(df_shared)

            # Preparar descarga Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_m.to_excel(writer, sheet_name='Matriz_Horaria', index=False)
                df_na_h.to_excel(writer, sheet_name='Sin_Turno_Horario', index=False)
                df_na_d.to_excel(writer, sheet_name='Sin_Turno_Diario', index=False)
                df_p.to_excel(writer, sheet_name='Totales_Periodo', index=False)
                df_shared.to_excel(writer, sheet_name='Franjas_Compartidas', index=False)
            
            st.download_button(
                label="📥 Descargar Reporte Completo",
                data=output.getvalue(),
                file_name=f"Reporte_Ventas_Airport.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error técnico: {e}")
    else:
        st.error("Por favor, sube ambos archivos.")
