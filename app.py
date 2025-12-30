import streamlit as st
import pandas as pd
import io
from processor import load_turnos, procesar_operacion_maestra

st.set_page_config(page_title="Sistema Aeropuerto Pro", layout="wide")
st.title("📊 Control de Productividad y Comisiones")

# SIDEBAR PARA CONFIGURACIÓN
with st.sidebar:
    st.header("1. Carga de Archivos")
    f_turnos = st.file_uploader("Subir Turnos (Excel)", type=["xlsx"])
    f_ventas = st.file_uploader("Subir Ventas (Excel)", type=["xlsx"])
    
    st.header("2. Rango de Análisis")
    f_inicio = st.date_input("Desde")
    f_fin = st.date_input("Hasta")

# PROCESAMIENTO
if st.button("🚀 Generar Reporte Maestro"):
    if not f_turnos or not f_ventas:
        st.error("Faltan archivos para procesar.")
    else:
        try:
            turnos = load_turnos(f_turnos)
            df_ventas = pd.read_excel(f_ventas)
            
            # Llamada a la función robusta
            resultados = procesar_operacion_maestra(df_ventas, turnos, f_inicio, f_fin)

            if "error" in resultados:
                st.warning(resultados["error"])
            else:
                tab1, tab2, tab3 = st.tabs(["📅 Reporte Diario", "📈 Resumen General", "⚙️ Mapeo"])

                with tab1:
                    st.subheader("Registro Diario por Columnas Fijas")
                    st.dataframe(resultados["reporte_diario"], use_container_width=False)

                with tab2:
                    st.subheader("Indicadores Acumulados")
                    st.dataframe(resultados["resumen_gral"], use_container_width=True)

                with tab3:
                    st.info("Referencia de posiciones de coordinadores en el reporte.")
                    st.table(resultados["mapeo"])

                # DESCARGA EXCEL
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    resultados["reporte_diario"].to_excel(writer, sheet_name="Diario_C1_C6", index=False)
                    resultados["resumen_gral"].to_excel(writer, sheet_name="Totales_Periodo", index=False)
                
                st.download_button(
                    label="📥 Descargar Reporte Completo",
                    data=buf.getvalue(),
                    file_name=f"Reporte_Productividad_{f_inicio}.xlsx",
                    mime="application/vnd.ms-excel"
                )

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
