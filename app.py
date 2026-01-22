import streamlit as st
import pandas as pd
from datetime import date
from processor import load_data_once, generate_initial_state_matrix, calculate_metrics_dynamic, generate_styled_excel

st.set_page_config(page_title="Simulador Cabify", layout="wide")

st.markdown("""
<style>
    .header {color: #7145D6; font-size: 28px; font-weight: bold; font-family: Arial;}
    .sub {color: #666; font-size: 16px; font-family: Arial;}
    .info {background-color: #F0EBFF; padding: 10px; border-radius: 5px; border-left: 4px solid #7145D6;}
</style>
<div class='header'>Simulador de Gestión Airport</div>
<div class='sub'>Ajusta la operación y recalcula comisiones y KPIs en tiempo real.</div>
<hr style='border: 1px solid #7145D6;'>
""", unsafe_allow_html=True)

if 'data_loaded' not in st.session_state: st.session_state.data_loaded = False
if 'state_matrix' not in st.session_state: st.session_state.state_matrix = None

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Cabify_Logo.svg/1200px-Cabify_Logo.svg.png", width=120)
    st.header("Configuración")
    t_file = st.file_uploader("1. Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("2. Ventas", type=['xlsx', 'csv'])
    st.divider()
    d_ini = st.date_input("Inicio", date(2025, 12, 1))
    d_fin = st.date_input("Fin", date(2025, 12, 31))
    
    st.info("""
    **Reglas de Loza/Colación (Corregidas):**
    • Turno 10:00 -> OFF: 10-11 y 14-16
    • Turno 05:00 -> OFF: 11-14
    • Turno 21:00 -> OFF: 05-08 (3 horas)
    """)
    
    if st.button("🔄 Cargar / Reiniciar"):
        if t_file and v_file:
            with st.spinner("Cargando y aplicando reglas..."):
                df_sales, turnos, names = load_data_once(v_file, t_file)
                if df_sales is not None:
                    initial_mx = generate_initial_state_matrix(turnos, names, d_ini, d_fin)
                    st.session_state.sales_df = df_sales
                    st.session_state.turnos = turnos
                    st.session_state.names = names
                    st.session_state.state_matrix = initial_mx
                    st.session_state.data_loaded = True
                    st.rerun()
                else:
                    st.error("Error al leer ventas (falta columna fecha).")

if st.session_state.data_loaded:
    tab1, tab2, tab3 = st.tabs(["🎛️ Simulación Interactiva", "📊 Resultados Finales", "📥 Exportar"])
    
    with tab1:
        st.markdown("""
        <div class='info'>
        <b>Instrucciones:</b><br>
        ✅ <b>Activado:</b> Coordinador vendiendo (Comisiona).<br>
        ⬜ <b>Desactivado:</b> Coordinador en Loza/Colación (No comisiona, aparece con *).
        </div>
        """, unsafe_allow_html=True)
        
        visible_cols = ["Fecha", "Hora"] + st.session_state.names
        
        col_cfg = {
            "Fecha": st.column_config.TextColumn(disabled=True),
            "Hora": st.column_config.TextColumn(disabled=True)
        }
        for n in st.session_state.names:
            col_cfg[n] = st.column_config.CheckboxColumn(n)
            
        edited_mx = st.data_editor(
            st.session_state.state_matrix,
            column_config=col_cfg,
            column_order=visible_cols,
            height=500,
            use_container_width=True,
            hide_index=True
        )
        st.session_state.state_matrix = edited_mx
        
    df_h, df_d, df_t, df_s = calculate_metrics_dynamic(
        st.session_state.sales_df, 
        st.session_state.turnos, 
        st.session_state.names, 
        st.session_state.state_matrix, 
        d_ini, d_fin
    )
    
    with tab2:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown("##### Matriz Resultante")
            st.dataframe(df_h, use_container_width=True)
        with c2:
            st.markdown("##### Liquidación & KPIs")
            st.dataframe(
                df_t.style.format({
                    "Ventas Totales": "${:,.0f}", 
                    "Comisión (2%)": "${:,.0f}",
                    "Promedio Venta/Hora": "${:,.0f}"
                }), 
                use_container_width=True
            )
            st.divider()
            st.markdown("##### Análisis de Competencia")
            st.dataframe(df_s, use_container_width=True)
            
    with tab3:
        st.success("Reporte listo para descargar con las modificaciones manuales aplicadas.")
        excel_data = generate_styled_excel({
            'Matriz_Horaria': df_h,
            'Resumen_Diario': df_d,
            'Totales_Comisiones': df_t,
            'Franjas_Compartidas': df_s
        })
        st.download_button("📥 Descargar Excel Final", excel_data, "Simulacion_Airport.xlsx")

else:
    st.info("👈 Carga los archivos en el menú lateral.")
