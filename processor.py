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
        # Limpieza de ruidos típicos en tu Excel
        t_clean = t_str.replace("diurno","").replace("nocturno","").replace("/","").strip()
        partes = t_clean.split('-')
        t_ini, t_fin = parse_time(partes[0]), parse_time(partes[1])
        return (t_ini, t_fin) if t_ini and t_fin else None
    except: return None

def read_file_safely(file):
    """Lee archivos Excel o CSV manejando errores de codificación."""
    if hasattr(file, 'name') and file.name.endswith('.xlsx'):
        return pd.read_excel(file, header=None)
    else:
        try:
            # Intento 1: UTF-8
            return pd.read_csv(file, header=None, encoding='utf-8', sep=None, engine='python')
        except UnicodeDecodeError:
            # Intento 2: Latin1 (Soluciona tu error de byte 0x9d)
            file.seek(0)
            return pd.read_csv(file, header=None, encoding='latin-1', sep=None, engine='python')

def load_turnos(file):
    df_raw = read_file_safely(file)
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

def procesar_maestro_v3(ventas_file, turnos_file, start_date, end_date):
    turnos = load_turnos(turnos_file)
    
    # Lectura de ventas con manejo de encoding
    if hasattr(ventas_file, 'name') and ventas_file.name.endswith('.xlsx'):
        sales = pd.read_excel(ventas_file)
    else:
        try:
            sales = pd.read_csv(ventas_file, encoding='utf-8', sep=None, engine='python')
        except UnicodeDecodeError:
            ventas_file.seek(0)
            sales = pd.read_csv(ventas_file, encoding='latin-1', sep=None, engine='python')

    sales['date'] = pd.to_datetime(sales['date'])
    sales = sales[(sales['date'].dt.date >= start_date) & (sales['date'].dt.date <= end_date)].copy()

    nombres_fijos = sorted(list(turnos.keys()))
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(nombres_fijos)}
    
    ventas_calc = []
    for _, row in sales.iterrows():
        activos = get_active_coordinators(row['date'], turnos)
        n = len(activos)
        if n > 0:
            for name in activos:
                ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': name, 'v': row['qt_price_local']/n, 'con_cuantos': n-1})
        else:
            ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': 'SIN ASIGNAR', 'v': row['qt_price_local'], 'con_cuantos': 0})
    
    df_v = pd.DataFrame(ventas_calc)

    matriz_data = []
    shared_stats = []
    
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            activos_h = get_active_coordinators(check_dt, turnos)
            fila_h = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            for nom, idx in mapa_cols.items():
                if nom in activos_h:
                    fila_h[f'Coordinador {idx}'] = nom
                    v_h = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['coord']==nom)]['v'].sum()
                    fila_h[f'Venta C{idx}'] = round(v_h)
                else:
                    fila_h[f'Coordinador {idx}'] = ""; fila_h[f'Venta C{idx}'] = 0
            matriz_data.append(fila_h)
        curr += timedelta(days=1)

    # Cálculo de Franjas Compartidas y Turnos
    for nom in nombres_fijos:
        # Contar turnos (días que tuvo algo asignado)
        turnos_totales = sum(1 for d in turnos[nom].values() if d is not None)
        
        # Analizar franjas compartidas basándose en la matriz generada
        df_matriz = pd.DataFrame(matriz_data)
        franjas_nom = df_matriz[df_matriz.iloc[:, 2::2].apply(lambda x: nom in x.values, axis=1)]
        
        solo = 0; con1 = 0; con2 = 0
        for _, r in franjas_nom.iterrows():
            otros = [v for v in r.iloc[2::2].values if v != "" and v != nom]
            if len(otros) == 0: solo += 1
            elif len(otros) == 1: con1 += 1
            else: con2 += 1
            
        v_total = df_v[df_v['coord']==nom]['v'].sum()
        shared_stats.append({
            'Coordinador': nom,
            'Ventas Totales': round(v_total),
            'Turnos (Días)': turnos_totales,
            'Horas Solo': solo,
            'Horas con 1 más': con1,
            'Horas con 2 o más': con2
        })

    return pd.DataFrame(matriz_data), pd.DataFrame(shared_stats)
