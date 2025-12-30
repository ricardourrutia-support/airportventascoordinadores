import streamlit as st
import pandas as pd
import io
from processor import load_turnos, procesar_maestro_airport

st.set_page_config(page_title="Airport Pro", layout="wide")
st.title("📊 Control de Productividad Aeropuerto")

with st.sidebar:
    st.header("Carga de Datos")
    f_turnos = st.file_uploader("Excel Turnos", type=["xlsx"])
    f_ventas = st.file_uploader("Excel Ventas", type=["xlsx"])
    f_ini = st.date_input("Inicio")
    f_fin = st.date_input("Fin")

if st.button("🚀 Procesar Reporte Completo"):
    if f_turnos and f_ventas:
        try:
            turnos = load_turnos(f_turnos)
            df_v = pd.read_excel(f_ventas)
            res = procesar_maestro_airport(df_v, turnos, f_ini, f_fin)

            if "error" in res:
                st.warning(res["error"])
            else:
                t1, t2 = st.tabs(["📅 Reporte Diario (C1-C6)", "📈 Resumen General"])
                with t1:
                    st.dataframe(res["diario"], use_container_width=False)
                with t2:
                    st.dataframe(res["resumen"], use_container_width=True)
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    res["diario"].to_excel(writer, sheet_name="Diario", index=False)
                    res["resumen"].to_excel(writer, sheet_name="Resumen", index=False)
                st.download_button("📥 Descargar Excel", buf.getvalue(), "Reporte_Productividad.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
