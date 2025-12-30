import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

def parse_time(t_str):
    t_str = t_str.strip().split(' ')[0]
    try:
        return datetime.strptime(t_str, "%H:%M:%S").time() if t_str.count(':') == 2 else datetime.strptime(t_str, "%H:%M").time()
    except: return None

def parse_turno_range(turno_raw):
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre"]: return None
    try:
        partes = t_str.split('-')
        t_ini, t_fin = parse_time(partes[0]), parse_time(partes[1])
        return (t_ini, t_fin) if t_ini and t_fin else None
    except: return None

def load_turnos(file_path):
    df_raw = pd.read_excel(file_path) if str(file_path).endswith('.xlsx') else pd.read_csv(file_path)
    # Asumiendo que la fila 0 tiene las fechas reales
    actual_dates = pd.to_datetime(df_raw.iloc[0, 1:], errors='coerce').dt.date.tolist()
    turnos_data = {}
    for i in range(2, len(df_raw)):
        row = df_raw.iloc[i]
        name = str(row[0]).strip()
        if name in ["nan", ""]: continue
        person_shifts = {d: parse_turno_range(row[idx+1]) for idx, d in enumerate(actual_dates) if idx+1 < len(row)}
        turnos_data[name] = person_shifts
    return turnos_data

def get_active_coordinators(sale_dt, turnos):
    s_date, s_time = sale_dt.date(), sale_dt.time()
    yesterday = s_date - timedelta(days=1)
    active = []
    for name, shifts in turnos.items():
        if s_date in shifts:
            start, end = shifts[s_date]
            if (start < end and start <= s_time < end) or (start > end and s_time >= start):
                active.append(name)
        if yesterday in shifts:
            start, end = shifts[yesterday]
            if start > end and s_time < end: # Cruce de medianoche
                active.append(name)
    return list(set(active))

def process_all(sales_file, turnos_file, start_date, end_date):
    turnos = load_turnos(turnos_file)
    sales = pd.read_excel(sales_file) if str(sales_file).endswith('.xlsx') else pd.read_csv(sales_file)
    sales['date'] = pd.to_datetime(sales['date'])
    sales = sales[(sales['date'].dt.date >= start_date) & (sales['date'].dt.date <= end_date)].copy()

    # LISTA MAESTRA FIJA (Asegúrate de que los nombres coincidan exactamente con el Excel de turnos)
    coord_list = ['Jocsanna Lopez', 'Alexis Cornu', 'Massiel Muñoz', 'Luis Fuentes', 'Cristobal Encina', 'Gerardo Palma']
    
    records = []
    for _, row in sales.iterrows():
        active = get_active_coordinators(row['date'], turnos)
        n = len(active)
        for name in (active if n > 0 else ['SIN ASIGNAR']):
            records.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coordinador': name, 'venta': row['qt_price_local'] / (n if n > 0 else 1)})

    df_res = pd.DataFrame(records)
    
    # MATRIZ HORARIA CON COLUMNAS FIJAS
    hourly_rows = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            activos = get_active_coordinators(check_dt, turnos)
            row_dict = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            for i, name in enumerate(coord_list):
                row_dict[f'Coordinador {i+1}'] = name if name in activos else ""
                v = df_res[(df_res['fecha'] == curr) & (df_res['hora'] == h) & (df_res['coordinador'] == name)]['venta'].sum()
                row_dict[f'Venta C{i+1}'] = round(v)
            hourly_rows.append(row_dict)
        curr += timedelta(days=1)
        
    return pd.DataFrame(hourly_rows), df_res.groupby(['fecha', 'coordinador'])['venta'].sum().unstack().fillna(0)
