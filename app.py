import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Dashboard Coordinadores", layout="wide")

st.title("🚀 Control Operativo de Ventas y Turnos")

with st.sidebar:
    st.header("Carga de Datos")
    t_file = st.file_uploader("Turnos (.xlsx)", type=["xlsx"])
    v_file = st.file_uploader("Ventas (.xlsx)", type=["xlsx"])
    f_i = st.date_input("Inicio")
    f_f = st.date_input("Fin")

if st.button("Generar Reportes"):
    if t_file and v_file:
        turnos = load_turnos(t_file)
        df_ventas = pd.read_excel(v_file)
        
        det, res, sin_age, visual = asignar_ventas(df_ventas, turnos, f_i, f_f)

        if det is not None:
            # MÉTRICAS RÁPIDAS
            c1, c2, c3 = st.columns(3)
            c1.metric("Ventas Totales", f"${det.drop_duplicates(subset=['hora_exacta'])['venta_original'].sum():,.0f}")
            c2.metric("Ventas Asignadas", f"${res['venta_asignada'].sum():,.0f}")
            c3.metric("Ventas SIN AGENTE", f"${sin_age['ventas_perdidas'].sum():,.0f}", delta_color="inverse")

            tab1, tab2, tab3 = st.tabs(["👀 Vista Supervisores", "⚠️ Sin Asignar", "💰 Totales"])

            with tab1:
                st.subheader("Distribución de Coordinadores por Franja")
                st.write("Usa esta tabla para ver rápidamente quién cubrió cada horario.")
                st.dataframe(visual, use_container_width=True)

            with tab2:
                st.subheader("Reporte de Ventas no reclamables (Sin Agente)")
                st.warning("Estas ventas ocurrieron en horarios donde nadie tenía turno asignado.")
                st.dataframe(sin_age, use_container_width=True)

            with tab3:
                st.subheader("Liquidación por Coordinador")
                st.table(res.style.format({"venta_asignada": "${:,.0f}"}))

            # DESCARGA
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                visual.to_excel(w, sheet_name="Vista_Supervisores", index=False)
                sin_age.to_excel(w, sheet_name="Sin_Agentes", index=False)
                res.to_excel(w, sheet_name="Totales", index=False)
            
            st.download_button("📥 Descargar Excel para Objeciones", buf.getvalue(), "reporte_operativo.xlsx")
