import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas_avanzado

st.set_page_config(page_title="Reporte Aeropuerto", layout="wide")
st.title("📊 Reporte de Productividad: Coordinadores")

with st.sidebar:
    st.header("Carga")
    t_file = st.file_uploader("Excel Turnos", type=["xlsx"])
    v_file = st.file_uploader("Excel Ventas", type=["xlsx"])
    f_i = st.date_input("Desde")
    f_f = st.date_input("Hasta")

if st.button("🚀 Generar Reporte"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_v = pd.read_excel(v_file)
            
            data = asignar_ventas_avanzado(df_v, turnos, f_i, f_f)

            if "error" in data:
                st.warning(data["error"])
            else:
                tab1, tab2 = st.tabs(["📅 Reporte Diario (Fijo)", "🏆 Resumen Acumulado"])
                
                with tab1:
                    st.subheader("Registro por Día y Coordinador")
                    st.dataframe(data["diario"], use_container_width=False)

                with tab2:
                    st.subheader("Indicadores de Periodo")
                    st.dataframe(data["resumen"], use_container_width=True)

                # Exportar
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    data["diario"].to_excel(w, sheet_name="Reporte_Diario", index=False)
                    data["resumen"].to_excel(w, sheet_name="Resumen_General", index=False)
                st.download_button("📥 Descargar Excel", buf.getvalue(), "Reporte_Productividad.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
