import streamlit as st
import pandas as pd
import io
from processor import load_turnos, asignar_ventas

st.set_page_config(page_title="Airport Sales Tracker", layout="wide")
st.title("📊 Gestión de Ventas y Coordinadores")

with st.sidebar:
    st.header("Entrada de Datos")
    t_file = st.file_uploader("Turnos (.xlsx)", type=["xlsx"])
    v_file = st.file_uploader("Ventas (.xlsx)", type=["xlsx"])
    f_i = st.date_input("Fecha Inicio")
    f_f = st.date_input("Fecha Fin")

if st.button("🚀 Analizar"):
    if t_file and v_file:
        try:
            turnos = load_turnos(t_file)
            df_v = pd.read_excel(v_file)
            
            # Recibimos el diccionario
            res = asignar_ventas(df_v, turnos, f_i, f_f)

            if "error" in res:
                st.error(res["error"])
            else:
                # EXTRACCIÓN SEGURA DE DATAFRAMES
                df_visual = res["visual"]
                df_sin_res = res["sin_res"]
                df_sin_det = res["sin_det"]
                df_pagos = res["pagos"]

                t1, t2, t3 = st.tabs(["📋 Vista Supervisores", "⚠️ Ventas Sin Agente", "💰 Pagos"])

                with t1:
                    st.subheader("Turnos por Franja Horaria")
                    st.dataframe(df_visual, use_container_width=True)

                with t2:
                    st.subheader("Reporte de Ventas No Asignadas")
                    st.metric("Total No Asignado", f"${df_sin_res['total_perdido'].sum():,.0f}")
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        st.write("**Resumen**")
                        st.dataframe(df_sin_res, use_container_width=True)
                    with c2:
                        st.write("**Detalle Individual**")
                        # Solo mostramos columnas relevantes si hay datos
                        if not df_sin_det.empty:
                            st.dataframe(df_sin_det[["fecha", "hora_exacta", "venta_original"]], use_container_width=True)
                        else:
                            st.success("¡No hay ventas sin asignar!")

                with t3:
                    st.subheader("Liquidación por Coordinador")
                    st.table(df_pagos.style.format({"venta_asignada": "${:,.0f}"}))

                # EXCEL
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                    df_visual.to_excel(w, sheet_name="Turnos_Franjas", index=False)
                    df_sin_res.to_excel(w, sheet_name="Resumen_Sin_Asignar", index=False)
                    df_sin_det.to_excel(w, sheet_name="Detalle_Ventas_Sin_Asignar", index=False)
                    df_pagos.to_excel(w, sheet_name="Pagos", index=False)
                
                st.download_button("📥 Descargar Reporte Completo", buf.getvalue(), "Reporte_Airport.xlsx")

        except Exception as e:
            st.error(f"Error en el proceso: {e}")
