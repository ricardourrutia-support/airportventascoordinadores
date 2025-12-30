import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas_completo

st.set_page_config(page_title="Reporte Operativo Airport", layout="wide")
st.title("📊 Análisis de Productividad por Coordinador")

with st.sidebar:
    st.header("Carga de Datos")
    t_file = st.file_uploader("Archivo de Turnos", type=["xlsx"])
    v_file = st.file_uploader("Base de Ventas", type=["xlsx"])
    f_i = st.date_input("Desde")
    f_f = st.date_input("Hasta")

if st.button("🚀 Generar Reporte"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_v = pd.read_excel(v_file)
            
            res = asignar_ventas_completo(df_v, turnos, f_i, f_f)

            if "error" in res:
                st.error(res["error"])
            else:
                tab1, tab2, tab3 = st.tabs(["📅 Reporte Diario", "🏆 Resumen General", "📍 Mapeo de Columnas"])

                with tab1:
                    st.subheader("Registro Diario de Indicadores")
                    st.write("Cada columna de coordinador es fija para el mismo agente durante todo el periodo.")
                    st.dataframe(res["reporte_diario"], use_container_width=False)

                with tab2:
                    st.subheader("Indicadores Totales por Agente")
                    st.dataframe(res["resumen_gral"], use_container_width=True)

                with tab3:
                    st.info("Esta tabla indica qué número de columna se le asignó a cada coordinador en el reporte.")
                    st.table(res["coordinadores_fijos"])

                # EXCEL
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    res["reporte_diario"].to_excel(w, sheet_name="Reporte_Diario", index=False)
                    res["resumen_gral"].to_excel(w, sheet_name="Totales", index=False)
                
                st.download_button("📥 Descargar Reporte en Excel", buf.getvalue(), "Reporte_Productividad.xlsx")

        except Exception as e:
            st.error(f"Error en el proceso: {e}")
