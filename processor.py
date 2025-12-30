import pandas as pd
from datetime import datetime, time
import numpy as np

def parse_turno(turno_raw):
    """Lógica robusta para limpiar y entender los horarios del Excel."""
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre", "nan"]: return None
    try:
        # Limpiar ruidos del Excel como "Diurno/Nocturno" o "/"
        t_clean = t_str.split("diurno")[0].split("nocturno")[0].replace("/", "").strip()
        partes = t_clean.split("-")
        
        def extract_time(txt):
            txt = txt.strip().split(" ")[0]
            # Si tiene segundos H:M:S, si no H:M
            p = txt.split(":")
            return time(int(p[0]), int(p[1]))

        return (extract_time(partes[0]), extract_time(partes[1]))
    except:
        return None

def load_turnos(file):
    """Carga el Excel y normaliza los nombres de coordinadores."""
    df_raw = pd.read_excel(file, header=None)
    # Las fechas están en la fila 1 (índice 1)
    fechas_raw = df_raw.iloc[1].tolist()
    fechas = [fechas_raw[0]] + list(pd.to_datetime(fechas_raw[1:], errors="coerce"))
    
    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    col_nombre = df.columns[0]
    
    turnos_dict = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip().upper()
        if nombre in ["NAN", ""]: continue
        
        # Guardar turnos por fecha. Usamos .date() para comparación fácil
        dias = {}
        for f in df.columns[1:]:
            if pd.isna(f): continue
            dias[f.date() if hasattr(f, 'date') else f] = parse_turno(row[f])
        turnos_dict[nombre] = dias
    return turnos_dict

def procesar_operacion_maestra(df_ventas, turnos, fecha_i, fecha_f):
    """Calcula todas las métricas solicitadas con asignación proporcional."""
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_periodo = df_ventas.loc[mask].copy()
    
    if df_periodo.empty:
        return {"error": "No se encontraron ventas en el rango seleccionado."}

    # 1. Crear mapa fijo de coordinadores (Mapeo Estático C1-C6)
    nombres_ordenados = sorted(list(turnos.keys()))
    # Mapear nombre -> ID de columna
    mapa_fijo = {nom: i+1 for i, nom in enumerate(nombres_ordenados)}
    
    registros_asignados = []
    
    for _, row in df_periodo.iterrows():
        f_hora = row['date']
        monto = row['qt_price_local']
        prod = str(row['ds_product_name']).lower()
        j_id = row['journey_id']
        
        # Buscar coordinadores activos en ese micro-momento
        activos = []
        for nom, d_turnos in turnos.items():
            r = d_turnos.get(f_hora.date())
            if r:
                hi, hf = r
                # Lógica de cruce de medianoche (Nocturnos)
                if (hi <= hf and hi <= f_hora.time() <= hf) or (hi > hf and (f_hora.time() >= hi or f_hora.time() <= hf)):
                    activos.append(nom)

        if activos:
            n = len(activos)
            for nom in activos:
                registros_asignados.append({
                    "fecha": f_hora.date(),
                    "coordinador": nom,
                    "id_col": mapa_fijo[nom],
                    "v": monto / n,
                    "j": 1 / n,
                    "p": 1 / n,
                    "pc": 1 / n if "compartida" in prod else 0,
                    "pe": 1 / n if "exclusive" in prod else 0
                })

    if not registros_asignados:
        return {"error": "Ventas encontradas, pero ningún coordinador tenía turno en esos horarios."}

    df_calculado = pd.DataFrame(registros_asignados)

    # --- GENERACIÓN DEL REPORTE DIARIO POR COLUMNAS (Mapeo C1..C6) ---
    reporte_diario = []
    for dia, g_dia in df_calculado.groupby("fecha"):
        fila = {"Día": dia}
        for nom, idx in mapa_fijo.items():
            g_c = g_dia[g_dia["coordinador"] == nom]
            # Columnas Nombre
            fila[f"Coordinador {idx}"] = nom if not g_c.empty else ""
            # Columnas Métricas
            fila[f"Ventas Coord {idx}"] = round(g_c["v"].sum(), 0)
            fila[f"Journeys Coord {idx}"] = round(g_c["j"].sum(), 1)
            fila[f"Pasajeros Coord {idx}"] = round(g_c["p"].sum(), 1)
            fila[f"Pasajeros Exclusivos C{idx}"] = round(g_c["pe"].sum(), 1)
            fila[f"Pasajeros Compartidos C{idx}"] = round(g_c["pc"].sum(), 1)
        reporte_diario.append(fila)

    # --- RESUMEN GENERAL PARA DASHBOARD ---
    resumen_gral = df_calculado.groupby("coordinador").agg({
        "v": "sum", "j": "sum", "p": "sum", "pe": "sum", "pc": "sum"
    }).round(1).reset_index()
    resumen_gral.columns = ["Coordinador", "Ventas", "Journeys", "Pasajeros", "P. Exclusivos", "P. Compartidos"]

    return {
        "reporte_diario": pd.DataFrame(reporte_diario),
        "resumen_gral": resumen_gral,
        "mapeo": pd.DataFrame(list(mapa_fijo.items()), columns=["Nombre", "Posición"])
    }
