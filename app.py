import streamlit as st
import pandas as pd
import io
from processor import load_turnos, procesar_reporte_productividad

st.set_page_config(page_title="Airport Pro Dashboard", layout="wide")
st.title("📊 Reporte de Productividad de Coordinadores")

# PANEL LATERAL
with st.sidebar:
    st.header("1. Carga de Datos")
    f_turnos = st.file_uploader("Excel de Turnos", type=["xlsx"])
    f_ventas = st.file_uploader("Excel de Ventas", type=["xlsx"])
    
    st.header("2. Filtros")
    f_ini = st.date_input("Fecha Inicio")
    f_fin = st.date_input("Fecha Fin")

# ACCIÓN PRINCIPAL
if st.button("🚀 Generar Reporte de Comisiones"):
    if f_turnos and f_ventas:
        try:
            # 1. Cargar datos
            turnos = load_turnos(f_turnos)
            df_v = pd.read_excel(f_ventas)
            
            # 2. Procesar con la lógica robusta
            res = procesar_reporte_productividad(df_v, turnos, f_ini, f_fin)

            if "error" in res:
                st.warning(res["error"])
            else:
                t1, t2, t3 = st.tabs(["📅 Reporte Diario (C1-C6)", "📈 Resumen General", "⚙️ Mapeo de Agentes"])

                with t1:
                    st.subheader("Registro Diario Multivariable")
                    st.dataframe(res["reporte_diario"], use_container_width=False)

                with t2:
                    st.subheader("Totales Acumulados")
                    st.dataframe(res["resumen_gral"], use_container_width=True)

                with t3:
                    st.info("Referencia de posiciones fijas en las columnas del reporte.")
                    st.table(res["mapeo"])

                # EXPORTACIÓN
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    res["reporte_diario"].to_excel(writer, sheet_name="Reporte_Diario", index=False)
                    res["resumen_gral"].to_excel(writer, sheet_name="Resumen_General", index=False)
                
                st.download_button(
                    label="📥 Descargar Reporte Excel",
                    data=buf.getvalue(),
                    file_name=f"Reporte_Productividad_{f_ini}.xlsx",
                    mime="application/vnd.ms-excel"
                )
        except Exception as e:
            st.error(f"Error técnico durante el procesamiento: {e}")
    else:
        st.error("Por favor, cargue ambos archivos Excel para continuar.")
