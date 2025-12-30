import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Gestión Airport", layout="wide")
st.title("📊 Control de Ventas y Cobertura de Coordinadores")

with st.sidebar:
    st.header("Carga de Archivos")
    t_file = st.file_uploader("Excel de Turnos", type=["xlsx"])
    v_file = st.file_uploader("Excel de Ventas", type=["xlsx"])
    f_i = st.date_input("Fecha Inicio")
    f_f = st.date_input("Fecha Fin")

if st.button("🚀 Generar Reportes y Análisis"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_ventas = pd.read_excel(v_file)
            
            # SOLUCIÓN AL ERROR: Ahora recibimos 5 variables (det, res, res_sin, visual, det_sin)
            det, res, res_sin, visual, det_sin = asignar_ventas(df_ventas, turnos, f_i, f_f)

            if det is not None:
                tab1, tab2, tab3 = st.tabs(["📋 Vista de Franjas y Turnos", "⚠️ Reporte Sin Asignar", "💰 Resumen de Pagos"])

                with tab1:
                    st.subheader("Visualización por Franja Horaria")
                    st.write("Detalle de coordinadores activos y sus turnos por cada hora.")
                    st.dataframe(visual, use_container_width=True)

                with tab2:
                    st.subheader("Análisis de Ventas Sin Coordinador")
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        st.write("**Resumen por Hora**")
                        st.dataframe(res_sin, use_container_width=True)
                        st.metric("Total No Asignado", f"${res_sin['ventas_totales_perdidas'].sum():,.0f}")
                    with c2:
                        st.write("**Registro de Ventas Huérfanas**")
                        st.dataframe(det_sin[["fecha", "hora_exacta", "venta_original"]], use_container_width=True)

                with tab3:
                    st.subheader("Total a Pagar por Coordinador")
                    st.table(res.style.format({"venta_asignada": "${:,.0f}"}))

                # Exportación Excel con todas las pestañas
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    visual.to_excel(w, sheet_name="Franjas_y_Coordinadores", index=False)
                    res_sin.to_excel(w, sheet_name="Resumen_Sin_Asignar", index=False)
                    det_sin.to_excel(w, sheet_name="Ventas_No_Asignadas", index=False)
                    res.to_excel(w, sheet_name="Totales_Pago", index=False)
                
                st.download_button("📥 Descargar Reporte Operativo", buf.getvalue(), f"Reporte_Airport.xlsx")
            else:
                st.warning("No hay datos para el rango seleccionado.")
        except Exception as e:
            st.error(f"Se produjo un error en el procesamiento: {e}")
