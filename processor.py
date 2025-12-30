import pandas as pd
from datetime import datetime, time

# ----------------------------------------------------------
# PARSER DE TURNOS MEJORADO
# ----------------------------------------------------------
def parse_turno(turno_raw):
    if pd.isna(turno_raw):
        return None
    turno_raw = str(turno_raw).strip().lower()
    if turno_raw == "" or turno_raw == "libre":
        return None

    try:
        # Limpieza: quitamos "diurno", "nocturno", "/" y espacios extra
        clean_txt = turno_raw.replace("diurno", "").replace("nocturno", "").replace("/", "").strip()
        partes = clean_txt.split("-")
        
        def extract_time(t_str):
            t_str = t_str.strip().split(" ")[0]
            # Manejar H:M o H:M:S
            parts = t_str.split(":")
            h = int(parts[0])
            m = int(parts[1])
            return time(h, m)

        h_ini = extract_time(partes[0])
        h_fin = extract_time(partes[1])
        return (h_ini, h_fin)
    except:
        return None

# ----------------------------------------------------------
# CARGA DE TURNOS
# ----------------------------------------------------------
def load_turnos(file):
    # Leer el excel, asumiendo que la fila 1 (index 1) tiene las fechas
    df_raw = pd.read_excel(file, header=None)
    
    # Extraer fechas de la fila 1 y limpiar
    fechas_fila = df_raw.iloc[1].tolist()
    # La primera celda suele ser "Nombre" o vacía, las demás son fechas
    fechas = [fechas_fila[0]] + list(pd.to_datetime(fechas_fila[1:], errors="coerce"))

    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    col_nombre = df.columns[0]

    turnos_dict = {}
    for _, row in df.iterrows():
        nombre = str(row[col_nombre]).strip()
        if nombre == "nan" or not nombre: continue
        
        turnos_persona = {}
        for fecha in df.columns[1:]:
            if pd.isna(fecha): continue
            val_turno = row[fecha]
            rango = parse_turno(val_turno)
            turnos_persona[fecha.date() if isinstance(fecha, datetime) else fecha] = rango
        
        turnos_dict[nombre] = turnos_persona
    
    return turnos_dict

# ----------------------------------------------------------
# ASIGNACIÓN DE VENTAS Y RESUMEN POR FRANJA
# ----------------------------------------------------------
def asignar_ventas(df_ventas, turnos, fecha_inicio, fecha_fin):
    # Asegurar tipos de fecha
    df_ventas['date'] = pd.to_datetime(df_ventas['date'])
    # Filtrar rango
    mask = (df_ventas['date'].dt.date >= fecha_inicio) & (df_ventas['date'].dt.date <= fecha_fin)
    df_filtrado = df_ventas.loc[mask].copy()

    if df_filtrado.empty:
        return None, None, None

    registros = []
    for _, row in df_filtrado.iterrows():
        f_hora = row['date']
        fecha_v = f_hora.date()
        hora_v = f_hora.time()
        monto = row['qt_price_local']

        activos = []
        for nombre, fechas_turnos in turnos.items():
            rango = fechas_turnos.get(fecha_v)
            if rango:
                h_ini, h_fin = rango
                # Lógica de cruce de medianoche o rango normal
                if h_ini <= h_fin:
                    if h_ini <= hora_v <= h_fin:
                        activos.append((nombre, rango))
                else: # Turno nocturno (ej: 22:00 a 06:00)
                    if hora_v >= h_ini or hora_v <= h_fin:
                        activos.append((nombre, rango))

        if activos:
            monto_div = monto / len(activos)
            for nombre, rango in activos:
                registros.append({
                    "fecha": fecha_v,
                    "hora": f_hora,
                    "franja_horaria": f"{f_hora.hour:02d}:00 - {f_hora.hour+1:02d}:00",
                    "coordinador": nombre,
                    "venta_original": monto,
                    "venta_asignada": monto_div,
                    "num_coordinadores": len(activos)
                })
        else:
            registros.append({
                "fecha": fecha_v,
                "hora": f_hora,
                "franja_horaria": f"{f_hora.hour:02d}:00 - {f_hora.hour+1:02d}:00",
                "coordinador": "SIN ASIGNAR",
                "venta_original": monto,
                "venta_asignada": 0,
                "num_coordinadores": 0
            })

    df_detallado = pd.DataFrame(registros)

    # 1. Resumen General (Lo que ya tenías)
    df_resumen_general = df_detallado.groupby("coordinador")["venta_asignada"].sum().reset_index()

    # 2. Resumen por Franjas Horarias (Nueva Visualización)
    df_franjas = df_detallado.groupby(["fecha", "franja_horaria"]).agg({
        "num_coordinadores": "max",
        "venta_original": "sum",
        "coordinador": lambda x: ", ".join(set(x))
    }).reset_index()
    df_franjas.columns = ["Fecha", "Franja", "Coordinadores Activos", "Venta Total Franja", "Nombres"]

    return df_detallado, df_resumen_general, df_franjas
