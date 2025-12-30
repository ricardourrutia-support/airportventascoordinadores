import pandas as pd
from datetime import datetime, time

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
            # Maneja H:M:S o H:M
            return time(int(parts[0]), int(parts[1]))
        return (extract_time(partes[0]), extract_time(partes[1]))
    except:
        return None

def load_turnos(file):
    df_raw = pd.read_excel(file, header=None)
    fechas_fila = df_raw.iloc[1].tolist()
    fechas = [fechas_fila[0]] + list(pd.to_datetime(fechas_fila[1:], errors="coerce"))
    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    col_nombre = df.columns[0]
    turnos_dict = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip().upper() # Normalizado
        if nombre == "NAN" or not nombre: continue
        dias = {}
        for fecha in df.columns[1:]:
            if pd.isna(fecha): continue
            dias[fecha.date() if isinstance(fecha, datetime) else fecha] = parse_turno(row[fecha])
        turnos_dict[nombre] = dias
    return turnos_dict

def asignar_ventas(df_ventas, turnos, fecha_i, fecha_f):
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    mask = (df_ventas['date'].dt.date >= fecha_i) & (df_ventas['date'].dt.date <= fecha_f)
    df_f = df_ventas.loc[mask].copy()
    
    if df_f.empty:
        return {"error": "No hay datos en el rango seleccionado"}

    registros = []
    for _, row in df_f.iterrows():
        f_hora = row['date']
        fecha_v = f_hora.date()
        hora_v = f_hora.time()
        monto = row['qt_price_local']
        
        activos = []
        for nombre, d_turnos in turnos.items():
            r = d_turnos.get(fecha_v)
            if r:
                h_i, h_f = r
                if (h_i <= h_f and h_i <= hora_v <= h_f) or (h_i > h_f and (hora_v >= h_i or hora_v <= h_f)):
                    activos.append({"nombre": nombre, "turno": f"{h_i.strftime('%H:%M')}-{h_f.strftime('%H:%M')}"})

        if activos:
            m_div = monto / len(activos)
            for a in activos:
                registros.append({
                    "fecha": fecha_v, "franja": f"{f_hora.hour:02d}:00", "hora_exacta": f_hora,
                    "coordinador": a['nombre'], "turno_ref": a['turno'], "venta_asignada": m_div, "venta_original": monto, "estado": "Asignado"
                })
        else:
            registros.append({
                "fecha": fecha_v, "franja": f"{f_hora.hour:02d}:00", "hora_exacta": f_hora,
                "coordinador": "SIN AGENTE", "turno_ref": "N/A", "venta_asignada": 0, "venta_original": monto, "estado": "No Asignado"
            })

    df_detallado = pd.DataFrame(registros)
    
    # Reportes específicos
    df_resumen_pago = df_detallado[df_detallado["estado"] == "Asignado"].groupby("coordinador")["venta_asignada"].sum().reset_index()
    
    df_sin_agente_det = df_detallado[df_detallado["estado"] == "No Asignado"].copy()
    df_sin_agente_res = df_sin_agente_det.groupby(["fecha", "franja"]).agg(
        ventas_totales_no_asignadas=("venta_original", "sum"),
        cantidad_viajes=("venta_original", "count")
    ).reset_index()

    # Vista Visual por columnas (Coordinador X | Turno X)
    vista_visual_list = []
    for (f, fr), group in df_detallado.groupby(["fecha", "franja"]):
        fila = {"Fecha": f, "Franja": fr}
        agentes = group[group["estado"] == "Asignado"][["coordinador", "turno_ref"]].drop_duplicates()
        for i, (_, r_ag) in enumerate(agentes.iterrows(), 1):
            fila[f"Coordinador {i}"] = r_ag["coordinador"]
            fila[f"Turno Coordinador {i}"] = r_ag["turno_ref"]
        fila["Venta Total Franja"] = group.drop_duplicates(subset=["hora_exacta"])["venta_original"].sum()
        vista_visual_list.append(fila)
    
    df_visual = pd.DataFrame(vista_visual_list).fillna("-")

    # RETORNO UNICO (DICCIONARIO)
    return {
        "detalle_completo": df_detallado,
        "resumen_pagos": df_resumen_pago,
        "sin_agente_resumen": df_sin_agente_res,
        "sin_agente_detalle": df_sin_agente_det,
        "vista_visual": df_visual
    }
