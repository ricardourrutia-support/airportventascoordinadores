import streamlit as st
import pandas as pd
import io
from processor import load_turnos, generar_matriz_operativa

st.set_page_config(page_title="Matriz de Cobertura Airport", layout="wide")

st.title("📊 Reporte de Cobertura Horaria por Coordinador")

with st.sidebar:
    st.header("Configuración")
    f_turnos = st.file_uploader("Subir Excel Turnos", type=["xlsx"])
    f_ventas = st.file_uploader("Subir Excel Ventas", type=["xlsx"]) # Reservado para KPIs futuros
    f_inicio = st.date_input("Fecha Inicio")
    f_fin = st.date_input("Fecha Fin")

if st.button("🚀 Generar Matriz Operativa"):
    if f_turnos:
        try:
            turnos = load_turnos(f_turnos)
            # Aunque no usemos ventas para la matriz, la cargamos si existe
            df_v = pd.read_excel(f_ventas) if f_ventas else None
            
            df_matriz = generar_matriz_operativa(df_v, turnos, f_inicio, f_fin)

            st.subheader(f"📅 Visualización del Periodo: {f_inicio} al {f_fin}")
            
            # Aplicar estilo para que las celdas vacías no distraigan
            st.dataframe(df_matriz, use_container_width=True, height=600)

            # Exportar a Excel
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df_matriz.to_excel(writer, sheet_name="Matriz_Operativa", index=False)
            
            st.download_button(
                label="📥 Descargar Matriz en Excel",
                data=buf.getvalue(),
                file_name="Matriz_Cobertura_Coordinadores.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"Error al procesar: {e}")
    else:
        st.error("Por favor, sube el archivo de turnos.")
