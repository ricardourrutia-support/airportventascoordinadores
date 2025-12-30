import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Gestión Airport", layout="wide")
st.title("📊 Control de Ventas y Cobertura")

with st.sidebar:
    st.header("Carga de Datos")
    t_file = st.file_uploader("Turnos", type=["xlsx"])
    v_file = st.file_uploader("Ventas", type=["xlsx"])
    f_i = st.date_input("Inicio")
    f_f = st.date_input("Fin")

if st.button("Procesar"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_v = pd.read_excel(v_file)
            
            # ESTA ES LA LÍNEA 22 - AHORA TIENE 5 VARIABLES
            det, res, res_sin, visual, det_sin = asignar_ventas(df_v, turnos, f_i, f_f)

            if det is not None:
                tab1, tab2, tab3 = st.tabs(["📋 Franjas y Turnos", "⚠️ Sin Asignar", "💰 Pagos"])
                with tab1:
                    st.dataframe(visual, use_container_width=True)
                with tab2:
                    st.metric("Total No Asignado", f"${res_sin['ventas_totales_perdidas'].sum():,.0f}")
                    st.subheader("Resumen por Franja")
                    st.dataframe(res_sin, use_container_width=True)
                    st.subheader("Detalle por Venta")
                    st.dataframe(det_sin[["fecha", "hora_exacta", "venta_original"]], use_container_width=True)
                with tab3:
                    st.dataframe(res, use_container_width=True)
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    visual.to_excel(w, sheet_name="Visual", index=False)
                    res_sin.to_excel(w, sheet_name="Resumen_Sin", index=False)
                    det_sin.to_excel(w, sheet_name="Detalle_Sin", index=False)
                    res.to_excel(w, sheet_name="Pagos", index=False)
                st.download_button("Descargar Reporte", buf.getvalue(), "Reporte.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
