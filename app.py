import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Gestión de Ventas Airport", layout="wide")

st.title("📊 Sistema de Asignación de Ventas por Coordinador")

# SECCIÓN DE CARGA
with st.sidebar:
    st.header("Configuración")
    turnos_file = st.file_uploader("Archivo de TURNOS", type=["xlsx"])
    ventas_file = st.file_uploader("Archivo de VENTAS", type=["xlsx"])
    
    fecha_inicio = st.date_input("Desde")
    fecha_fin = st.date_input("Hasta")

if st.button("🚀 Procesar Datos"):
    if not turnos_file or not ventas_file:
        st.error("Por favor, sube ambos archivos.")
    else:
        try:
            turnos = load_turnos(turnos_file)
            df_ventas = pd.read_excel(ventas_file) # O pd.read_csv si es el caso
            
            df_detallado, df_totales, df_franjas = asignar_ventas(
                df_ventas, turnos, fecha_inicio, fecha_fin
            )

            if df_detallado is not None:
                tab1, tab2, tab3 = st.tabs(["Resumen General", "Vista por Franjas", "Detalle de Ventas"])
                
                with tab1:
                    st.subheader("💰 Venta Total Asignada por Coordinador")
                    st.dataframe(df_totales, use_container_width=True)
                
                with tab2:
                    st.subheader("⏰ Visualización por Franjas Horarias")
                    st.write("Muestra cuántos coordinadores hubo y qué se vendió en cada hora.")
                    st.dataframe(df_franjas, use_container_width=True)
                
                with tab3:
                    st.subheader("📋 Registro Detallado")
                    st.dataframe(df_detallado)

                # DESCARGA
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_totales.to_excel(writer, sheet_name="Resumen", index=False)
                    df_franjas.to_excel(writer, sheet_name="Franjas_Horarias", index=False)
                    df_detallado.to_excel(writer, sheet_name="Detalle_Completo", index=False)
                
                st.download_button(
                    label="📥 Descargar Reporte Completo",
                    data=buffer.getvalue(),
                    file_name="reporte_coordinadores.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.warning("No se encontraron ventas para el rango seleccionado.")
        except Exception as e:
            st.error(f"Error en el proceso: {e}")
