import pandas as pd
from datetime import datetime, time

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
    except: return None

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

def generar_matriz_operativa(df_ventas, turnos, fecha_i, fecha_f):
    # Generar rango de fechas
    rango_dias = pd.date_range(fecha_i, fecha_f)
    matriz_final = []

    for dia in rango_dias:
        fecha_actual = dia.date()
        # Iterar por las 24 horas del día
        for hora in range(24):
            hora_inicio = time(hora, 0)
            hora_fin = time((hora + 1) % 24, 0)
            
            # Buscar coordinadores activos en esta hora específica
            activos = []
            for nombre, d_turnos in turnos.items():
                r = d_turnos.get(fecha_actual)
                if r:
                    hi, hf = r
                    # Lógica de cruce de medianoche
                    if (hi <= hf and hi <= hora_inicio < hf) or (hi > hf and (hora_inicio >= hi or hora_inicio < hf)):
                        activos.append(nombre)
            
            # Crear la fila del reporte
            fila = {
                "Día": fecha_actual,
                "Hora Inicio": f"{hora:02d}:00",
                "Hora Fin": f"{(hora+1)%24:02d}:00"
            }
            
            # Llenar Coordinador 1 al 6
            for i in range(6):
                fila[f"Coordinador {i+1}"] = activos[i] if i < len(activos) else ""
            
            matriz_final.append(fila)

    return pd.DataFrame(matriz_final)
