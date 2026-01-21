import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# --- REGLAS DE GESTIÓN (LOZA/COLACIÓN) ---
def is_in_special_hours(shift_start_time, h_curr):
    """
    Retorna True si el coordinador está en horario de Loza/Colación.
    Optimizado para recibir enteros (horas).
    """
    if shift_start_time is None: return False
    h_start = shift_start_time.hour
    
    # Regla 1: Turno de 10:00 (Gestión OFF: 10-11 y 14-16)
    if h_start == 10:
        if h_curr == 10: return True
        if 14 <= h_curr < 16: return True
        
    # Regla 2: Turno de 05:00 (Gestión OFF: 11-14)
    elif h_start == 5:
        if 11 <= h_curr < 14: return True
        
    # Regla 3: Turno de 21:00 (Gestión OFF: 06-09)
    elif h_start == 21:
        if 6 <= h_curr < 9: return True
        
    return False

# --- FUNCIONES DE PARSEO ---
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

def read_file_generic(file, has_header=True):
    header_val = 0 if has_header else None
    if hasattr(file, 'name') and file.name.endswith('.xlsx'):
        return pd.read_excel(file, header=header_val)
    else:
        try:
            return pd.read_csv(file, header=header_val, encoding='utf-8', sep=None, engine='python')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, header=header_val, encoding='latin-1', sep=None, engine='python')

def load_turnos(file):
    df_raw = read_file_generic(file, has_header=False)
    actual_dates = pd.to_datetime(df_raw.iloc[1, 1:], errors='coerce').dt.date.tolist()
    
    turnos_data = {}
    ordered_names = []
    
    for i in range(2, len(df_raw)):
        row = df_raw.iloc[i]
        nombre = str(row.iloc[0]).strip()
        if nombre.upper() in ["NAN", "", "NOMBRE"]: continue
        
        if nombre not in ordered_names:
            ordered_names.append(nombre)
            
        dias = {f: parse_turno_range(row.iloc[j+1]) for j, f in enumerate(actual_dates) if f is not pd.NaT}
        turnos_data[nombre] = dias
        
    return turnos_data, ordered_names

# --- GENERADOR DE EXCEL ---
def generate_styled_excel(dfs_dict):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    header_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'white', 'bg_color': '#7145D6', 'border': 0, 'align': 'center', 'valign': 'vcenter'})
    cell_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#333333', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center', 'valign': 'vcenter'})
    name_col_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#7145D6', 'bg_color': '#F9F9F9', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'left'})
    special_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#9C4A00', 'bg_color': '#FFEB9C', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center'})

    for sheet_name, df in dfs_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
            ws.set_column(idx, idx, max_len, cell_fmt)
            if idx == 0: ws.set_column(idx, idx, max_len + 2, name_col_fmt)
            ws.write(0, idx, col, header_fmt)
            
            if sheet_name == 'Matriz_Horaria':
                ws.conditional_format(1, idx, len(df), idx, {'type': 'text', 'criteria': 'containing', 'value': '(*)', 'format': special_fmt})

        ws.hide_gridlines(2)
    writer.close()
    return output.getvalue()

def process_all(sales_file, turnos_file, start_date, end_date):
    # 1. Cargar Turnos
    turnos, ordered_names = load_turnos(turnos_file)
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(ordered_names)}
    
    # 2. Cargar y Limpiar Ventas
    df_sales = read_file_generic(sales_file, has_header=True)
    df_sales.columns = [c.strip() for c in df_sales.columns]
    
    # Renombrar fecha
    if 'createdAt_local' in df_sales.columns:
        df_sales.rename(columns={'createdAt_local': 'date'}, inplace=True)
    elif 'date' not in df_sales.columns:
        for col in df_sales.columns:
            if 'date' in col.lower() or 'created' in col.lower() or 'fecha' in col.lower():
                df_sales.rename(columns={col: 'date'}, inplace=True)
                break
    
    if 'date' not in df_sales.columns:
        return None, None, None, None

    df_sales['date'] = pd.to_datetime(df_sales['date'])
    # Filtro de fecha
    df_sales = df_sales[(df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)].copy()
    
    # PRE-PROCESAMIENTO: Agrupar ventas por (Fecha, Hora) para evitar iterar miles de filas
    # Esto reduce la complejidad de O(N_ventas) a O(N_horas)
    df_sales['date_only'] = df_sales['date'].dt.date
    df_sales['hour_only'] = df_sales['date'].dt.hour
    
    # Suma de ventas por hora
    sales_grouped = df_sales.groupby(['date_only', 'hour_only'])['qt_price_local'].sum().to_dict()

    # 3. Generar Matriz Horaria y Distribuir Ventas (Iterando Horas, no ventas)
    matriz_data = []
    daily_accumulators = {name: {} for name in ordered_names} # {name: {date: total}}
    
    # Contadores para estadísticas
    stats_franjas = {name: {'Solo': 0, 'Con 1': 0, 'Con 2+': 0} for name in ordered_names}
    
    curr = start_date
    while curr <= end_date:
        # Pre-chequear si hay ventas este día para optimizar
        # (Opcional, pero ayuda)
        
        for h in range(24):
            # Obtener ventas totales de esa hora (diccionario es instantáneo)
            total_venta_hora = sales_grouped.get((curr, h), 0)
            
            # Determinar Activos y Eligibles
            # Lógica de turnos se ejecuta 24 veces por día (rápido)
            check_dt_time = time(h, 0)
            
            fisicos = [] # Lista de nombres
            eligibles = [] # Lista de nombres
            
            # Check rápido de turnos
            # Optimización: iterar turnos una vez
            yesterday = curr - timedelta(days=1)
            
            for name in ordered_names:
                shifts = turnos.get(name, {})
                
                # Check Turno Hoy
                if curr in shifts and shifts[curr]:
                    start, end = shifts[curr]
                    # Lógica de cruce y pertenencia
                    if (start < end and start <= check_dt_time < end) or (start > end and (check_dt_time >= start or check_dt_time < end)):
                        fisicos.append(name)
                        if not is_in_special_hours(start, h):
                            eligibles.append(name)
                        continue # Ya lo encontramos hoy
                
                # Check Turno Ayer (Overflow)
                if yesterday in shifts and shifts[yesterday]:
                    start, end = shifts[yesterday]
                    if start > end and check_dt_time < end:
                        fisicos.append(name)
                        if not is_in_special_hours(start, h):
                            eligibles.append(name)

            # Distribuir ventas
            n_eligibles = len(eligibles)
            venta_asignada = total_venta_hora / n_eligibles if n_eligibles > 0 else 0
            
            # Construir fila
            fila = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            
            for name in ordered_names:
                idx = mapa_cols[name]
                if name in fisicos:
                    if name in eligibles:
                        # Está trabajando y vendiendo
                        fila[f'Coord {idx}'] = name
                        fila[f'Venta C{idx}'] = round(venta_asignada)
                        
                        # Acumular diario
                        daily_accumulators[name][curr] = daily_accumulators[name].get(curr, 0) + venta_asignada
                        
                        # Estadísticas de competencia
                        others = n_eligibles - 1
                        if others == 0: stats_franjas[name]['Solo'] += 1
                        elif others == 1: stats_franjas[name]['Con 1'] += 1
                        else: stats_franjas[name]['Con 2+'] += 1
                        
                    else:
                        # Está en Loza/Colación
                        fila[f'Coord {idx}'] = f"{name} (*)"
                        fila[f'Venta C{idx}'] = 0
                else:
                    # No está
                    fila[f'Coord {idx}'] = ""
                    fila[f'Venta C{idx}'] = 0
            
            matriz_data.append(fila)
        curr += timedelta(days=1)

    # 4. Construir DataFrames Finales
    df_hourly = pd.DataFrame(matriz_data)
    
    # Resumen Diario
    daily_rows = []
    curr = start_date
    while curr <= end_date:
        r = {'Día': curr}
        for name in ordered_names:
            r[name] = round(daily_accumulators[name].get(curr, 0))
        daily_rows.append(r)
        curr += timedelta(days=1)
    df_daily = pd.DataFrame(daily_rows)
    
    # Totales y Estadísticas
    total_metrics = []
    shared_stats_list = []
    
    for name in ordered_names:
        # Sumar todo lo acumulado
        total_v = sum(daily_accumulators[name].values())
        
        # Contar días trabajados
        dias_trabajados = 0
        if name in turnos:
            for d, rng in turnos[name].items():
                if start_date <= d <= end_date and rng is not None:
                    dias_trabajados += 1
        
        total_metrics.append({
            'Coordinador': name,
            'Ventas Totales': round(total_v),
            'Turnos Trabajados': dias_trabajados
        })
        
        shared_stats_list.append({
            'Coordinador': name,
            'Horas Solo (Venta)': stats_franjas[name]['Solo'],
            'Horas con 1 (Venta)': stats_franjas[name]['Con 1'],
            'Horas con 2+ (Venta)': stats_franjas[name]['Con 2+']
        })
        
    return df_hourly, df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats_list)
