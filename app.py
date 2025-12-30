import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Airport Sales Tracker", layout="wide")

st.title("📊 Control de Operación y Ventas")

with st.sidebar:
    st.header("Carga de Datos")
    t_file = st.file_uploader("Subir Turnos", type=["xlsx"])
    v_file = st.file_uploader("Subir Ventas", type=["xlsx"])
    f_i = st.date_input("Inicio")
    f_f = st.date_input("Fin")

if st.button("🚀 Ejecutar Análisis"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_v = pd.read_excel(v_file)
            
            # Llamada robusta: devuelve un diccionario
            resultados = asignar_ventas(df_v, turnos, f_i, f_f)

            if "error" in resultados:
                st.warning(resultados["error"])
            else:
                # Extraemos los datos del diccionario
                visual = resultados["vista_visual"]
                res_sin = resultados["sin_agente_resumen"]
                det_sin = resultados["sin_agente_detalle"]
                pagos = resultados["resumen_pagos"]

                tab1, tab2, tab3 = st.tabs(["📋 Control de Franjas", "⚠️ Ventas Sin Agente", "💰 Resumen Pagos"])

                with tab1:
                    st.subheader("Distribución de Turnos por Hora")
                    st.dataframe(visual, use_container_width=True)

                with tab2:
                    st.error(f"Total No Asignado: ${res_sin['ventas_totales_no_asignadas'].sum():,.0f}")
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        st.write("**Resumen por Franja**")
                        st.dataframe(res_sin, use_container_width=True)
                    with c2:
                        st.write("**Detalle de Ventas Huérfanas**")
                        st.dataframe(det_sin[["fecha", "hora_exacta", "venta_original"]], use_container_width=True)

                with tab3:
                    st.subheader("Total Asignado a Coordinadores")
                    st.table(pagos.style.format({"venta_asignada": "${:,.0f}"}))

                # Botón de Descarga
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    visual.to_excel(w, sheet_name="Franjas", index=False)
                    res_sin.to_excel(w, sheet_name="Resumen_Sin_Asignar", index=False)
                    pagos.to_excel(w, sheet_name="Pagos", index=False)
                st.download_button("📥 Descargar Excel", buf.getvalue(), "Reporte_Completo.xlsx")

        except Exception as e:
            st.error(f"Error crítico: {e}")
