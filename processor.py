import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

def parse_time(t_str):
    t_str = str(t_str).strip().lower().split(' ')[0]
    try:
        if t_str.count(':') == 2:
            return datetime.strptime(t_str, "%H:%M:%S").time()
        return datetime.strptime(t_str, "%H:%M").time()
    except: return None

def parse_turno_range(turno_raw):
    if pd.isna(turno_raw): return None
    t_str = str(turno_raw).strip().lower()
    if t_str in ["", "libre", "nan"]: return None
    try:
        t_clean = t_str.replace("diurno","").replace("nocturno","").replace("/","").strip()
        partes = t_clean.split('-')
        t_ini, t_fin = parse_time(partes[0]), parse_time(partes[1])
        return (t_ini, t_fin) if t_ini and t_fin else None
    except: return None

def load_turnos(file_path):
    df_raw = pd.read_excel(file_path, header=None) if str(file_path.name).endswith('.xlsx') else pd.read_csv(file_path, header=None, encoding='latin1', sep=None, engine='python')
    actual_dates = pd.to_datetime(df_raw.iloc[1, 1:], errors='coerce').dt.date.tolist()
    df = df_raw.iloc[2:].copy()
    df.columns = [df_raw.iloc[2,0]] + actual_dates
    
    turnos_data = {}
    for _, row in df.iterrows():
        nombre = str(row.iloc[0]).strip().upper()
        if nombre in ["NAN", "", "NOMBRE"]: continue
        dias = {f: parse_turno_range(row.iloc[i+1]) for i, f in enumerate(actual_dates) if f is not pd.NaT}
        turnos_data[nombre] = dias
    return turnos_data

def get_active_coordinators(sale_dt, turnos):
    s_date, s_time = sale_dt.date(), sale_dt.time()
    yesterday = s_date - timedelta(days=1)
    active = []
    for name, shifts in turnos.items():
        if s_date in shifts and shifts[s_date]:
            start, end = shifts[s_date]
            if (start < end and start <= s_time < end) or (start > end and (s_time >= start or s_time < end)):
                active.append(name)
        if yesterday in shifts and shifts[yesterday]:
            start, end = shifts[yesterday]
            if start > end and s_time < end: active.append(name)
    return list(set(active))

def procesar_final_airport(ventas_file, turnos_file, start_date, end_date):
    turnos = load_turnos(turnos_file)
    sales = pd.read_excel(ventas_file) if str(ventas_file.name).endswith('.xlsx') else pd.read_csv(ventas_file, encoding='latin1', sep=None, engine='python')
    sales['date'] = pd.to_datetime(sales['date'])
    sales = sales[(sales['date'].dt.date >= start_date) & (sales['date'].dt.date <= end_date)].copy()

    # MAPEO FIJO: El Coordinador 1 siempre será el mismo nombre
    nombres_fijos = sorted(list(turnos.keys()))
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(nombres_fijos)}
    
    ventas_calc = []
    for _, row in sales.iterrows():
        activos = get_active_coordinators(row['date'], turnos)
        n = len(activos)
        if n > 0:
            for name in activos:
                ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': name, 'v': row['qt_price_local']/n, 'asignado': True})
        else:
            ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': 'SIN ASIGNAR', 'v': row['qt_price_local'], 'asignado': False})
    
    df_v = pd.DataFrame(ventas_calc)

    matriz_data = []
    na_horario = []
    
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            activos_h = get_active_coordinators(datetime.combine(curr, time(h, 0)), turnos)
            fila_h = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            
            for nom, idx in mapa_cols.items():
                if nom in activos_h:
                    fila_h[f'Coordinador {idx}'] = nom
                    v_h = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['coord']==nom)]['v'].sum()
                    fila_h[f'Venta C{idx}'] = round(v_h)
                else:
                    fila_h[f'Coordinador {idx}'] = ""
                    fila_h[f'Venta C{idx}'] = 0
            matriz_data.append(fila_h)
            
            # Ventas No Asignadas
            v_na = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['asignado']==False)]['v'].sum()
            na_horario.append({'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00', 'Venta No Asignada': round(v_na)})
        curr += timedelta(days=1)

    df_na_h = pd.DataFrame(na_horario)
    df_na_d = df_na_h.groupby('Día')['Venta No Asignada'].sum().reset_index()
    resumen_p = df_v[df_v['asignado']==True].groupby('coord')['v'].sum().round(0).reset_index()
    resumen_p.columns = ['Coordinador', 'Venta Total']

    return pd.DataFrame(matriz_data), df_na_h, df_na_d, resumen_p
