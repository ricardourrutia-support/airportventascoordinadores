import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

def parse_time(t_str):
    t_str = str(t_str).strip().split(' ')[0]
    try:
        # Intenta H:M:S o H:M
        if t_str.count(':') == 2:
            return datetime.strptime(t_str, "%H:%M:%S").time()
        return datetime.strptime(t_str, "%H:%M").time()
    except: return None

def parse_turno_range(turno_raw):
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre", "nan"]: return None
    try:
        partes = t_str.split('-')
        t_ini, t_fin = parse_time(partes[0]), parse_time(partes[1])
        return (t_ini, t_fin) if t_ini and t_fin else None
    except: return None

def load_turnos(file_path):
    # Detección robusta de formato y codificación
    if str(file_path.name).endswith('.xlsx'):
        df_raw = pd.read_excel(file_path, header=None)
    else:
        # Si es CSV, intentamos con latin1 que es común en Excels guardados como CSV
        df_raw = pd.read_csv(file_path, header=None, encoding='latin1', sep=None, engine='python')

    # Fila 1 tiene las fechas (según tu estructura)
    fechas_fila = df_raw.iloc[1].tolist()
    fechas = [fechas_fila[0]] + list(pd.to_datetime(fechas_fila[1:], errors='coerce'))
    
    df = df_raw.iloc[2:].copy()
    df.columns = fechas
    
    turnos_data = {}
    for _, row in df.iterrows():
        nombre = str(row.iloc[0]).strip().upper()
        if nombre in ["NAN", "", "NOMBRE"]: continue
        
        # Guardar turnos por fecha
        dias = {}
        for idx, f in enumerate(df.columns[1:]):
            if pd.isna(f): continue
            # Ajustar índice porque f empieza desde la col 1
            val = row.iloc[idx + 1]
            dias[f.date() if hasattr(f, 'date') else f] = parse_turno_range(val)
        turnos_data[nombre] = dias
    return turnos_data

def get_active_coordinators(sale_dt, turnos):
    s_date, s_time = sale_dt.date(), sale_dt.time()
    yesterday = s_date - timedelta(days=1)
    active = []
    for name, shifts in turnos.items():
        # Turno del mismo día
        if s_date in shifts and shifts[s_date]:
            start, end = shifts[s_date]
            if (start < end and start <= s_time < end) or (start > end and s_time >= start):
                active.append(name)
        # Turno del día anterior (cruce de medianoche)
        if yesterday in shifts and shifts[yesterday]:
            start, end = shifts[yesterday]
            if start > end and s_time < end:
                active.append(name)
    return list(set(active))

def procesar_v2_fijo(ventas_file, turnos_file, start_date, end_date):
    turnos = load_turnos(turnos_file)
    
    if str(ventas_file.name).endswith('.xlsx'):
        sales = pd.read_excel(ventas_file)
    else:
        sales = pd.read_csv(ventas_file, encoding='latin1', sep=None, engine='python')

    sales['date'] = pd.to_datetime(sales['date'])
    sales = sales[(sales['date'].dt.date >= start_date) & (sales['date'].dt.date <= end_date)].copy()

    # --- MAPEO ESTÁTICO DE COORDINADORES ---
    # Esto asegura que el Coordinador 1 sea SIEMPRE el mismo nombre
    lista_coordinadores = sorted(list(turnos.keys()))
    mapa_fijo = {nombre: i+1 for i, nombre in enumerate(lista_coordinadores)}
    
    records = []
    for _, row in sales.iterrows():
        activos = get_active_coordinators(row['date'], turnos)
        n = len(activos)
        if n > 0:
            for name in activos:
                records.append({
                    'fecha': row['date'].date(), 
                    'hora': row['date'].hour, 
                    'coordinador': name, 
                    'venta': row['qt_price_local'] / n
                })
    df_res = pd.DataFrame(records)

    # --- MATRIZ DE CASILLEROS FIJOS ---
    hourly_rows = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            activos_ahora = get_active_coordinators(check_dt, turnos)
            fila = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            
            for nombre, col_idx in mapa_fijo.items():
                if nombre in activos_ahora:
                    fila[f'Coordinador {col_idx}'] = nombre
                    v = df_res[(df_res['fecha'] == curr) & (df_res['hora'] == h) & (df_res['coordinador'] == nombre)]['venta'].sum()
                    fila[f'Venta C{col_idx}'] = round(v)
                else:
                    fila[f'Coordinador {col_idx}'] = "" # Vacío si no está
                    fila[f'Venta C{col_idx}'] = 0
            
            hourly_rows.append(fila)
        curr += timedelta(days=1)
        
    return pd.DataFrame(hourly_rows), lista_coordinadores
