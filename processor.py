import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

def parse_time(t_str):
    t_str = t_str.strip().split(' ')[0]
    try:
        if t_str.count(':') == 2:
            return datetime.strptime(t_str, "%H:%M:%S").time()
        else:
            return datetime.strptime(t_str, "%H:%M").time()
    except:
        return None

def parse_turno_range(turno_raw):
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str == "" or t_str == "libre": return None
    try:
        partes = t_str.split('-')
        if len(partes) < 2: return None
        ini_str = partes[0].strip()
        fin_str = partes[1].strip().split(' ')[0]
        t_ini = parse_time(ini_str)
        t_fin = parse_time(fin_str)
        if t_ini and t_fin:
            return (t_ini, t_fin)
        return None
    except:
        return None

def load_turnos(file_path):
    if str(file_path).endswith('.xlsx'):
        df_raw = pd.read_excel(file_path)
    else:
        df_raw = pd.read_csv(file_path)
    
    # La fila 0 contiene las fechas reales en tu formato de Excel
    actual_dates = pd.to_datetime(df_raw.iloc[0, 1:], errors='coerce').dt.date.tolist()
    
    turnos_data = {}
    # Los nombres comienzan en la fila con índice 2
    for i in range(2, len(df_raw)):
        row = df_raw.iloc[i]
        name = str(row[0]).strip()
        if name == "nan" or name == "": continue
        
        person_shifts = {}
        for col_idx, d in enumerate(actual_dates):
            if col_idx + 1 < len(row):
                val = row[col_idx + 1]
                rng = parse_turno_range(val)
                if rng:
                    person_shifts[d] = rng
        turnos_data[name] = person_shifts
    return turnos_data

def get_active_coordinators(sale_dt, turnos):
    sale_date = sale_dt.date()
    sale_time = sale_dt.time()
    yesterday = sale_date - timedelta(days=1)
    
    active = []
    for name, shifts in turnos.items():
        if sale_date in shifts:
            start, end = shifts[sale_date]
            if start < end:
                if start <= sale_time < end:
                    active.append(name)
            else: # Turno que cruza la medianoche
                if sale_time >= start:
                    active.append(name)
        if yesterday in shifts:
            start, end = shifts[yesterday]
            if start > end:
                if sale_time < end:
                    active.append(name)
    return list(set(active))

def process_all(sales_file, turnos_file, start_date, end_date):
    turnos = load_turnos(turnos_file)
    if str(sales_file).endswith('.xlsx'):
        sales = pd.read_excel(sales_file)
    else:
        sales = pd.read_csv(sales_file)
        
    sales['date'] = pd.to_datetime(sales['date'])
    mask = (sales['date'].dt.date >= start_date) & (sales['date'].dt.date <= end_date)
    sales = sales.loc[mask].copy()
    
    # Lista maestra de coordinadores (C1-C6)
    coord_list = ['Jocsanna Lopez', 'Alexis Cornu', 'Massiel Muñoz', 'Luis Fuentes', 'Cristobal Encina', 'Gerardo Palma']
    
    records = []
    for _, row in sales.iterrows():
        s_dt = row['date']
        price = row['qt_price_local']
        active = get_active_coordinators(s_dt, turnos)
        n = len(active)
        if n > 0:
            for name in active:
                records.append({
                    'fecha': s_dt.date(), 'hora': s_dt.hour, 'coordinador': name, 'venta': price / n
                })
        else:
            records.append({
                'fecha': s_dt.date(), 'hora': s_dt.hour, 'coordinador': 'SIN ASIGNAR', 'venta': price
            })

    df_results = pd.DataFrame(records)
    
    # --- Generación de Matriz Horaria ---
    hourly_rows = []
    current = start_date
    while current <= end_date:
        for h in range(24):
            check_dt = datetime.combine(current, time(h, 0))
            active_names = get_active_coordinators(check_dt, turnos)
            row_dict = {'Día': current, 'Hora Inicio': f'{h:02d}:00', 'Hora Fin': f'{(h+1)%24:02d}:00'}
            
            # Asignar Nombres y Ventas en columnas fijas
            h_data = df_results[(df_results['fecha'] == current) & (df_results['hora'] == h)]
            for i, name in enumerate(coord_list):
                row_dict[f'Coordinador {i+1}'] = name if name in active_names else ""
                v = h_data[h_data['coordinador'] == name]['venta'].sum()
                row_dict[f'Ventas C{i+1}'] = round(v) if v > 0 else 0
            
            hourly_rows.append(row_dict)
        current += timedelta(days=1)
    
    df_hourly = pd.DataFrame(hourly_rows)
    
    # --- Resumen Diario ---
    daily_summaries = []
    current = start_date
    while current <= end_date:
        day_data = df_results[df_results['fecha'] == current]
        sum_dict = {'Día': current}
        for i, name in enumerate(coord_list):
            v = day_data[day_data['coordinador'] == name]['venta'].sum()
            sum_dict[f'Ventas Coord {i+1}'] = round(v)
        daily_summaries.append(sum_dict)
        current += timedelta(days=1)
    
    return df_hourly, pd.DataFrame(daily_summaries), df_results.groupby('coordinador')['venta'].sum().reset_index() df_hourly, df_daily, df_total
