import pandas as pd
from datetime import datetime, time
import numpy as np

def parse_turno(turno_raw):
    """Parsea horarios complejos eliminando texto extra como 'Diurno/Nocturno'."""
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre", "nan"]: return None
    try:
        # Limpiar texto y separar por guion
        t_clean = t_str.split("diurno")[0].split("nocturno")[0].replace("/", "").strip()
        partes = t_clean.split("-")
        
        def extract_time(txt):
            txt = txt.strip().split(" ")[0]
            p = txt.split(":")
            # Retorna objeto time (Hora, Minuto)
            return time(int(p[0]), int(p[1]))

        return (extract_time(partes[0]), extract_time(partes[1]))
    except:
        return None

def load_turnos(file):
    """Carga el Excel de turnos identificando fechas en la fila 1."""
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
        
        dias = {}
        for f in df.columns[1:]:
            if pd.isna(f): continue
            # Guardamos la fecha como objeto date para cruce exacto
            d_key = f.date() if hasattr(f, 'date') else f
            dias[d_key] = parse_turno(row[f])
        turnos_dict[nombre] = dias
    return turnos_dict

def procesar_reporte_productividad(df_ventas, turnos, fecha_i, fecha_f):
    """Lógica principal: Distribución proporcional y mapeo de columnas C1-C6."""
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_p = df_ventas.loc[mask].copy()
    
    if df_p.empty:
        return {"error": "No hay ventas en el rango de fechas seleccionado."}

    # Crear mapeo fijo de coordinadores alfabéticamente (C1, C2, C3...)
    nombres_master = sorted(list(turnos.keys()))
    mapa_fijo = {nom: i+1 for i, nom in enumerate(nombres_master)}
    
    registros = []
    for _, row in df_p.iterrows():
        f_hora = row['date']
        monto = row['qt_price_local']
        prod = str(row['ds_product_name']).lower()
        
        # Detectar quiénes estaban trabajando en ese instante
        activos = []
        for nom, d_turnos in turnos.items():
            r = d_turnos.get(f_hora.date())
            if r:
                hi, hf = r
                # Lógica para turnos normales y nocturnos (cruce medianoche)
                if (hi <= hf and hi <= f_hora.time() <= hf) or (hi > hf and (f_hora.time() >= hi or f_hora.time() <= hf)):
                    activos.append(nom)

        if activos:
            n = len(activos)
            for nom in activos:
                registros.append({
                    "fecha": f_hora.date(),
                    "coordinador": nom,
                    "id_col": mapa_fijo[nom],
                    "v": monto / n,
                    "j": 1 / n, # Journey proporcional
                    "p": 1 / n, # Pasajero proporcional
                    "pc": 1 / n if "compartida" in prod else 0,
                    "pe": 1 / n if "exclusive" in prod else 0
                })

    if not registros:
        return {"error": "Se encontraron ventas pero no había coordinadores en turno para esos horarios."}

    df_calc = pd.DataFrame(registros)

    # Construcción del Reporte Diario con Columnas C1...C6
    reporte_diario = []
    for dia, g_dia in df_calc.groupby("fecha"):
        fila = {"Día": dia}
        for nom, idx in mapa_fijo.items():
            g_c = g_dia[g_dia["coordinador"] == nom]
            # Si el coordinador estuvo ese día, ponemos su nombre, si no, vacío
            fila[f"Coordinador {idx}"] = nom if not g_c.empty else ""
            fila[f"Ventas Coord {idx}"] = round(g_c["v"].sum(), 0)
            fila[f"Journeys Coord {idx}"] = round(g_c["j"].sum(), 1)
            fila[f"Pasajeros Coord {idx}"] = round(g_c["p"].sum(), 1)
            fila[f"Pasajeros Exclusivos C{idx}"] = round(g_c["pe"].sum(), 1)
            fila[f"Pasajeros Compartidos C{idx}"] = round(g_c["pc"].sum(), 1)
        reporte_diario.append(fila)

    # Resumen General Consolidado
    resumen_gral = df_calc.groupby("coordinador").agg({
        "v": "sum", "j": "sum", "p": "sum", "pe": "sum", "pc": "sum"
    }).round(1).reset_index()
    resumen_gral.columns = ["Coordinador", "Ventas", "Journeys", "Pasajeros", "P. Exclusivos", "P. Compartidos"]

    return {
        "reporte_diario": pd.DataFrame(reporte_diario),
        "resumen_gral": resumen_gral,
        "mapeo": pd.DataFrame(list(mapa_fijo.items()), columns=["Nombre", "Columna Asignada"])
    }
