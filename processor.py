import pandas as pd
from datetime import datetime, time
import numpy as np

def parse_turno(turno_raw):
    if pd.isna(turno_raw): return None
    turno_raw = str(turno_raw).strip().lower()
    if turno_raw in ["", "libre"]: return None
    try:
        clean_txt = turno_raw.replace("diurno", "").replace("nocturno", "").replace("/", "").strip()
        partes = clean_txt.split("-")
        def extract_time(t_str):
            t_str = t_str.strip().split(" ")[0]
            parts = t_str.split(":")
            return time(int(parts[0]), int(parts[1]))
        return (extract_time(partes[0]), extract_time(partes[1]))
    except: return None

def load_turnos(file):
    df_raw = pd.read_excel(file, header=None)
    fechas_fila = df_raw.iloc[1].tolist()
    fechas = [fechas_fila[0]] + list(pd.to_datetime(fechas_fila[1:], errors="coerce"))
    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    col_nombre = df.columns[0]
    turnos_dict = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip().upper()
        if nombre == "NAN" or not nombre: continue
        dias = {f.date() if isinstance(f, datetime) else f: parse_turno(row[f]) for f in df.columns[1:] if not pd.isna(f)}
        turnos_dict[nombre] = dias
    return turnos_dict

def asignar_ventas_completo(df_ventas, turnos, fecha_i, fecha_f):
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_f = df_ventas.loc[mask].copy()
    
    if df_f.empty: return {"error": "No hay datos en el periodo"}

    # Identificar todos los coordinadores únicos para asignarles una columna fija
    todos_coordinadores = sorted(list(turnos.keys()))
    coord_map = {nom: i+1 for i, nom in enumerate(todos_coordinadores)}
    
    registros = []
    for _, row in df_f.iterrows():
        f_hora = row['date']
        fecha_v = f_hora.date()
        hora_v = f_hora.time()
        monto = row['qt_price_local']
        prod = str(row['ds_product_name']).lower()
        
        activos = []
        for nombre, d_turnos in turnos.items():
            r = d_turnos.get(fecha_v)
            if r:
                h_i, h_f = r
                if (h_i <= h_f and h_i <= hora_v <= h_f) or (h_i > h_f and (hora_v >= h_i or hora_v <= h_f)):
                    activos.append(nombre)

        if activos:
            n = len(activos)
            for nombre in activos:
                registros.append({
                    "fecha": fecha_v,
                    "coordinador": nombre,
                    "id_col": coord_map[nombre],
                    "venta": monto / n,
                    "journey": 1 / n,
                    "pasajero": 1 / n,
                    "p_compartido": (1 / n) if "compartida" in prod else 0,
                    "p_exclusivo": (1 / n) if "exclusive" in prod else 0
                })

    df_det = pd.DataFrame(registros)
    if df_det.empty: return {"error": "No se pudieron asignar ventas a ningún turno"}

    # --- REPORTE DIARIO ADAPTATIVO ---
    reporte_diario = []
    for dia, g_dia in df_det.groupby("fecha"):
        fila = {"Día": dia}
        # Inicializar columnas para los coordinadores mapeados
        for nombre, idx in coord_map.items():
            g_coord = g_dia[g_dia["coordinador"] == nombre]
            fila[f"Coordinador {idx}"] = nombre if not g_coord.empty else ""
            fila[f"Ventas Coord {idx}"] = round(g_coord["venta"].sum(), 0)
            fila[f"Journeys Coord {idx}"] = round(g_coord["journey"].sum(), 1)
            fila[f"Pasajeros Coord {idx}"] = round(g_coord["pasajero"].sum(), 1)
            fila[f"Pasajeros Exclusivos C{idx}"] = round(g_coord["p_exclusivo"].sum(), 1)
            fila[f"Pasajeros Compartidos C{idx}"] = round(g_coord["p_compartido"].sum(), 1)
        reporte_diario.append(fila)

    # --- RESUMEN GENERAL ---
    resumen_gral = df_det.groupby("coordinador").agg({
        "venta": "sum",
        "journey": "sum",
        "pasajero": "sum",
        "p_exclusivo": "sum",
        "p_compartido": "sum"
    }).round(1).reset_index()

    return {
        "reporte_diario": pd.DataFrame(reporte_diario),
        "resumen_gral": resumen_gral,
        "coordinadores_fijos": pd.DataFrame(list(coord_map.items()), columns=["Nombre", "Columna Asignada"])
    }
