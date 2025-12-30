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

def asignar_ventas_avanzado(df_ventas, turnos, fecha_i, fecha_f):
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_f = df_ventas.loc[mask].copy()
    
    if df_f.empty: return {"error": "Sin datos en el rango."}

    # Definir orden fijo de coordinadores (para que C1 sea siempre el mismo)
    coords_lista = sorted(list(turnos.keys()))
    coord_pos = {nom: i+1 for i, nom in enumerate(coords_lista)}
    
    registros = []
    for _, row in df_f.iterrows():
        f_hora = row['date']
        monto = row['qt_price_local']
        prod = str(row['ds_product_name']).lower()
        
        activos = []
        for nombre, d_turnos in turnos.items():
            r = d_turnos.get(f_hora.date())
            if r:
                hi, hf = r
                if (hi <= hf and hi <= f_hora.time() <= hf) or (hi > hf and (f_hora.time() >= hi or f_hora.time() <= hf)):
                    activos.append(nombre)

        if activos:
            n = len(activos)
            for nom in activos:
                registros.append({
                    "fecha": f_hora.date(), "coord": nom, "pos": coord_pos[nom],
                    "v": monto/n, "j": 1/n, "p": 1/n,
                    "pc": 1/n if "compartida" in prod else 0,
                    "pe": 1/n if "exclusive" in prod else 0
                })

    if not registros: return {"error": "No hubo coordinadores en las horas de venta."}
    df_res = pd.DataFrame(registros)

    # REPORTE DIARIO POR COLUMNAS
    diario = []
    for dia, g_dia in df_res.groupby("fecha"):
        fila = {"Día": dia}
        for nom, i in coord_pos.items():
            g_c = g_dia[g_dia["coord"] == nom]
            fila[f"Coordinador {i}"] = nom if not g_c.empty else ""
            fila[f"Ventas Coord {i}"] = round(g_c["v"].sum(), 0)
            fila[f"Journeys Coord {i}"] = round(g_c["j"].sum(), 1)
            fila[f"Pasajeros Coord {i}"] = round(g_c["p"].sum(), 1)
            fila[f"P. Exclusivos C{i}"] = round(g_c["pe"].sum(), 1)
            fila[f"P. Compartidos C{i}"] = round(g_c["pc"].sum(), 1)
        diario.append(fila)

    # RESUMEN GENERAL (Métricas solicitadas)
    resumen = df_res.groupby("coord").agg({
        "v": "sum", "j": "sum", "p": "sum", "pe": "sum", "pc": "sum"
    }).round(1).reset_index()
    resumen.columns = ["Coordinador", "Ventas Totales", "Total Journeys", "Total Pasajeros", "Pasajeros Exclusivos", "Pasajeros Compartidos"]

    return {"diario": pd.DataFrame(diario), "resumen": resumen}
