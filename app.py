import streamlit as st
import pandas as pd
import io
from datetime import date
from processor import procesar_maestro

st.set_page_config(page_title="Airport Dashboard", layout="wide")
st.title("📊 Control de Ventas y Cobertura Fija")

with st.sidebar:
    t_file = st.file_uploader("Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Ventas", type=['xlsx', 'csv'])
    d_ini = st.date_input("Inicio", date(2025, 11, 1))
    d_fin = st.date_input("Fin", date(2025, 11, 30))

if st.sidebar.button("🚀 Procesar"):
    if t_file and v_file:
        df_m, df_na_h, df_na_d, df_p = procesar_maestro(v_file, t_file, d_ini, d_fin)
        
        tab1, tab2, tab3 = st.tabs(["⏰ Matriz Coordinadores", "⚠️ Ventas No Asignadas", "🏆 Totales Periodo"])
        
        with tab1:
            st.subheader("Mapa de Cobertura (Casilleros Fijos)")
            st.dataframe(df_m, use_container_width=True)
            
        with tab2:
            st.subheader("Ventas sin Coordinador en Turno")
            c1, c2 = st.columns(2)
            c1.write("**Por Día**")
            c1.dataframe(df_na_d)
            c2.metric("Total No Asignado", f"${df_na_d['Venta No Asignada'].sum():,.0f}")
            st.write("**Detalle por Franja Horaria**")
            st.dataframe(df_na_h)

        with tab3:
            st.subheader("Resumen de Ventas por Coordinador")
            st.table(df_p)

        # Excel con todas las pestañas solicitadas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_m.to_excel(writer, sheet_name='Matriz_Coordinadores', index=False)
            df_na_h.to_excel(writer, sheet_name='No_Asignados_Franja', index=False)
            df_na_d.to_excel(writer, sheet_name='No_Asignados_Diario', index=False)
            df_p.to_excel(writer, sheet_name='Totales_Periodo', index=False)
        
        st.download_button("📥 Descargar Reporte Completo", output.getvalue(), f"Reporte_Airport_{d_ini}.xlsx")
