import pandas as pd
from datetime import datetime, time
import numpy as np

def parse_turno(turno_raw):
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre", "nan"]: return None
    try:
        t_clean = t_str.split("diurno")[0].split("nocturno")[0].replace("/", "").strip()
        partes = t_clean.split("-")
        def extract_time(txt):
            txt = txt.strip().split(" ")[0]
            p = txt.split(":")
            return time(int(p[0]), int(p[1]))
        return (extract_time(partes[0]), extract_time(partes[1]))
    except:
        return None

def load_turnos(file):
    df_raw = pd.read_excel(file, header=None)
    fechas_raw = df_raw.iloc[1].tolist()
    fechas = [fechas_raw[0]] + list(pd.to_datetime(fechas_raw[1:], errors="coerce"))
    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    col_nombre = df.columns[0]
    turnos_dict = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip().upper()
        if nombre in ["NAN", ""]: continue
        dias = {f.date() if hasattr(f, 'date') else f: parse_turno(row[f]) for f in df.columns[1:] if not pd.isna(f)}
        turnos_dict[nombre] = dias
    return turnos_dict

def procesar_maestro_airport(df_ventas, turnos, fecha_i, fecha_f):
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_p = df_ventas.loc[mask].copy()
    
    if df_p.empty:
        return {"error": "Sin datos en el rango seleccionado."}

    # Mapa fijo de coordinadores (C1, C2...) ordenado alfabéticamente
    nombres_master = sorted(list(turnos.keys()))
    mapa_fijo = {nom: i+1 for i, nom in enumerate(nombres_master)}
    
    registros = []
    for _, row in df_p.iterrows():
        f_hora = row['date']
        monto = row['qt_price_local']
        prod = str(row['ds_product_name']).lower()
        
        activos = []
        for nom, d_turnos in turnos.items():
            r = d_turnos.get(f_hora.date())
            if r:
                hi, hf = r
                if (hi <= hf and hi <= f_hora.time() <= hf) or (hi > hf and (f_hora.time() >= hi or f_hora.time() <= hf)):
                    activos.append(nom)

        if activos:
            n = len(activos)
            for nom in activos:
                registros.append({
                    "fecha": f_hora.date(), "coord": nom, "v": monto/n, "j": 1/n, "p": 1/n,
                    "pc": 1/n if "compartida" in prod else 0, "pe": 1/n if "exclusive" in prod else 0
                })

    if not registros: return {"error": "No hay coordinadores en los horarios de venta."}
    df_calc = pd.DataFrame(registros)

    # Reporte Diario Adaptativo
    reporte_diario = []
    for dia, g_dia in df_calc.groupby("fecha"):
        fila = {"Día": dia}
        for nom, idx in mapa_fijo.items():
            g_c = g_dia[g_dia["coord"] == nom]
            fila[f"Coordinador {idx}"] = nom if not g_c.empty else ""
            fila[f"Ventas C{idx}"] = round(g_c["v"].sum(), 0)
            fila[f"Journeys C{idx}"] = round(g_c["j"].sum(), 1)
            fila[f"Pasajeros C{idx}"] = round(g_c["p"].sum(), 1)
            fila[f"P_Excl C{idx}"] = round(g_c["pe"].sum(), 1)
            fila[f"P_Comp C{idx}"] = round(g_c["pc"].sum(), 1)
        reporte_diario.append(fila)

    resumen_gral = df_calc.groupby("coord").agg({"v":"sum","j":"sum","p":"sum","pe":"sum","pc":"sum"}).round(1).reset_index()
    return {"diario": pd.DataFrame(reporte_diario), "resumen": resumen_gral, "mapeo": pd.DataFrame(list(mapa_fijo.items()), columns=["Nombre", "ID"])}
