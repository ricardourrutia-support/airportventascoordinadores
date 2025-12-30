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
        nombre = str(row[col_nombre]).strip()
        if nombre == "nan" or not nombre: continue
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
    if df_f.empty: return None, None, None, None

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
    
    # 1. Resumen General por Coordinador
    df_resumen = df_detallado[df_detallado["estado"] == "Asignado"].groupby("coordinador")["venta_asignada"].sum().reset_index()

    # 2. Resumen Ventas SIN ASIGNAR (Para objeciones)
    df_sin_agente = df_detallado[df_detallado["estado"] == "No Asignado"].groupby(["fecha", "franja"]).agg(
        ventas_perdidas=("venta_original", "sum"),
        cantidad_ventas=("venta_original", "count")
    ).reset_index()

    # 3. Vista Visual para Supervisores (Columnas por Coordinador)
    # Agrupamos por fecha/franja y listamos los coordinadores presentes
    vista_sup = []
    for (f, fr), group in df_detallado.groupby(["fecha", "franja"]):
        fila = {"Fecha": f, "Franja": fr}
        agentes_en_franja = group[group["estado"] == "Asignado"][["coordinador", "turno_ref"]].drop_duplicates()
        
        for i, (_, row_ag) in enumerate(agentes_en_franja.iterrows(), 1):
            fila[f"Coordinador {i}"] = row_ag["coordinador"]
            fila[f"Turno {i}"] = row_ag["turno_ref"]
        
        fila["Venta Total Franja"] = group["venta_original"].unique().sum() if group["estado"].iloc[0] == "No Asignado" else group.drop_duplicates(subset=["hora_exacta"])["venta_original"].sum()
        vista_sup.append(fila)
    
    df_vista_visual = pd.DataFrame(vista_sup).fillna("-")

    return df_detallado, df_resumen, df_sin_agente, df_vista_visual
