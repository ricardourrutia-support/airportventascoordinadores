import streamlit as st
import pandas as pd
from datetime import date
from processor import load_data_once, generate_initial_state_matrix, calculate_metrics_dynamic, generate_styled_excel

st.set_page_config(page_title="Simulador Airport", layout="wide")

# Estilos
st.markdown("""
<style>
    .header {color: #7145D6; font-size: 28px; font-weight: bold;}
    .sub {color: #666; font-size: 16px;}
    .highlight {background-color: #F3F0FF; padding: 10px; border-radius: 5px; border: 1px solid #7145D6;}
</style>
<div class='header'>Simulador de Gestión Airport</div>
<div class='sub'>Ajusta la operación y recalcula comisiones en tiempo real.</div>
<hr>
""", unsafe_allow_html=True)

# 1. INICIALIZACIÓN DE ESTADO
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'state_matrix' not in st.session_state:
    st.session_state.state_matrix = None

# 2. SIDEBAR - CONFIGURACIÓN
with st.sidebar:
    st.header("Configuración")
    t_file = st.file_uploader("Turnos", type=['xlsx', 'csv'])
    v_file = st.file_uploader("Ventas", type=['xlsx', 'csv'])
    d_ini = st.date_input("Inicio", date(2025, 12, 1))
    d_fin = st.date_input("Fin", date(2025, 12, 31))
    
    st.warning("""
    **Modo Interactivo:**
    1. Carga los archivos.
    2. Modifica los checkboxes en la pestaña "Simulación".
    3. Los totales y comisiones se actualizan al instante.
    """)
    
    load_btn = st.button("🔄 Cargar / Reiniciar Datos")

# 3. LÓGICA DE CARGA
if load_btn and t_file and v_file:
    with st.spinner("Cargando datos base..."):
        # Cargar raw data
        df_sales, turnos, names = load_data_once(v_file, t_file)
        
        if df_sales is not None:
            # Generar matriz de estado inicial (Vende=True, Loza=False, Ausente=None)
            initial_matrix = generate_initial_state_matrix(turnos, names, d_ini, d_fin)
            
            # Guardar en sesión
            st.session_state.sales_df = df_sales
            st.session_state.turnos_dict = turnos
            st.session_state.ordered_names = names
            st.session_state.state_matrix = initial_matrix
            st.session_state.data_loaded = True
            st.rerun()
        else:
            st.error("Error en archivos.")

# 4. INTERFAZ PRINCIPAL
if st.session_state.data_loaded:
    # Recuperar datos de sesión
    df_s = st.session_state.sales_df
    turnos = st.session_state.turnos_dict
    names = st.session_state.ordered_names
    
    # --- PESTAÑAS ---
    tab_sim, tab_res, tab_exp = st.tabs(["🎛️ Simulación Operativa", "📊 Resultados y Comisiones", "📥 Exportar"])
    
    with tab_sim:
        st.markdown("##### Tablero de Control Operativo")
        st.caption("Marca ✅ para 'Vendiendo' (Actividad 1) o desmarca ⬜ para 'Loza/Colación' (Actividad 2). Solo puedes editar coordinadores presentes.")
        
        # CONFIGURACIÓN DEL EDITOR DE DATOS
        # Ocultamos date_idx y hour_idx de la edición, pero las necesitamos para el índice
        column_config = {
            "date_idx": st.column_config.TextColumn("Fecha", disabled=True),
            "hour_idx": st.column_config.NumberColumn("Hora", disabled=True, format="%d:00")
        }
        # Configurar columnas de coordinadores como Checkboxes
        for name in names:
            column_config[name] = st.column_config.CheckboxColumn(
                name,
                help=f"Activar/Desactivar venta para {name}",
                default=False
            )

        # MOSTRAR EDITOR
        # El usuario edita 'edited_matrix', y esto actualiza st.session_state.state_matrix automáticamente si lo manejamos bien
        edited_matrix = st.data_editor(
            st.session_state.state_matrix,
            column_config=column_config,
            height=400,
            use_container_width=True,
            key="editor_key" # Clave para persistencia interna de Streamlit
        )
        
        # Actualizar la matriz de estado en la sesión con lo que editó el usuario
        st.session_state.state_matrix = edited_matrix

    # --- CÁLCULO DINÁMICO ---
    # Se ejecuta en cada interacción con el editor
    df_h, df_d, df_t, df_s = calculate_metrics_dynamic(df_s, turnos, names, st.session_state.state_matrix, d_ini, d_fin)

    with tab_res:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown("##### Matriz Horaria Resultante")
            st.dataframe(df_h, use_container_width=True, height=400)
        with c2:
            st.markdown("##### Liquidación Final")
            # Mostrar tabla de totales con formato de moneda si es posible, o simple
            st.dataframe(df_t.style.format({"Ventas Totales": "${:,.0f}", "Comisión (2%)": "${:,.0f}"}), use_container_width=True)
            
            st.divider()
            st.markdown("##### Franjas Compartidas")
            st.dataframe(df_s, use_container_width=True)

    with tab_exp:
        st.write("Descarga el reporte con las modificaciones manuales aplicadas.")
        excel_bytes = generate_styled_excel({
            'Matriz_Horaria': df_h,
            'Resumen_Diario': df_d,
            'Totales_Comisiones': df_t,
            'Franjas_Compartidas': df_s
        })
        st.download_button("📥 Descargar Excel Final", excel_bytes, "Simulacion_Airport.xlsx")

else:
    st.info("👆 Carga los archivos en el menú lateral para comenzar la simulación.")
