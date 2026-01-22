import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# --- REGLAS DE GESTIÓN (LOZA/COLACIÓN) ---
def get_initial_status(shift_start, h_curr):
    """
    Define si está VENDIENDO (True) o en LOZA (False).
    """
    if shift_start is None: return False
    h_start = shift_start.hour
    
    # Regla 1: Turno de 10:00 (Gestión OFF: 10-11 y 14-16)
    if h_start == 10:
        if h_curr == 10: return False
        if 14 <= h_curr < 16: return False
        
    # Regla 2: Turno de 05:00 (Gestión OFF: 11-14)
    elif h_start == 5:
        if 11 <= h_curr < 14: return False
        
    # Regla 3: Turno de 21:00 (Gestión OFF: 05-08) <-- CORREGIDO
    elif h_start == 21:
        if 5 <= h_curr < 8: return False
        
    return True 

# --- PARSEO DE FECHAS ---
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

# --- CARGA DE DATOS ---
def load_data_once(sales_file, turnos_file):
    # 1. Turnos
    df_turnos = read_file_generic(turnos_file, has_header=False)
    actual_dates = pd.to_datetime(df_turnos.iloc[1, 1:], errors='coerce').dt.date.tolist()
    
    turnos_dict = {}
    ordered_names = [] 
    
    for i in range(2, len(df_turnos)):
        row = df_turnos.iloc[i]
        nombre = str(row.iloc[0]).strip()
        if nombre.upper() in ["NAN", "", "NOMBRE"]: continue
        
        if nombre not in ordered_names: ordered_names.append(nombre)
        dias = {f: parse_turno_range(row.iloc[j+1]) for j, f in enumerate(actual_dates) if f is not pd.NaT}
        turnos_dict[nombre] = dias

    # 2. Ventas
    df_sales = read_file_generic(sales_file, has_header=True)
    df_sales.columns = [c.strip() for c in df_sales.columns]
    
    if 'createdAt_local' in df_sales.columns:
        df_sales.rename(columns={'createdAt_local': 'date'}, inplace=True)
    elif 'date' not in df_sales.columns:
        for col in df_sales.columns:
            if 'date' in col.lower() or 'created' in col.lower() or 'fecha' in col.lower():
                df_sales.rename(columns={col: 'date'}, inplace=True)
                break
    
    if 'date' not in df_sales.columns: return None, None, None

    df_sales['date'] = pd.to_datetime(df_sales['date'])
    return df_sales, turnos_dict, ordered_names

# --- GENERACIÓN DE MATRIZ DE ESTADO (CORREGIDA) ---
def generate_initial_state_matrix(turnos, ordered_names, start_date, end_date):
    data = []
    curr = start_date
    while curr <= end_date:
        yesterday = curr - timedelta(days=1)
        for h in range(24):
            check_time = time(h, 0) # e.g. 05:00:00
            
            # Indices ocultos para la lógica interna
            row = {'_date_str': str(curr), '_hour': h}
            
            # Columnas visibles en el Editor
            row['Fecha'] = str(curr)
            row['Hora'] = f"{h:02d}:00"
            
            for name in ordered_names:
                shifts = turnos.get(name, {})
                status = None 
                
                # --- CORRECCIÓN CRÍTICA DE TURNOS NOCTURNOS ---
                
                # 1. Revisar turno asignado a HOY
                if curr in shifts and shifts[curr]:
                    s, e = shifts[curr]
                    if s > e: 
                        # Turno cruza medianoche (Ej: 21:00 a 08:00)
                        # HOY solo es válido desde la hora de inicio hasta las 23:59
                        if h >= s.hour:
                            status = get_initial_status(s, h)
                    else:
                        # Turno normal (Ej: 10:00 a 21:00)
                        if s <= check_time < e:
                            status = get_initial_status(s, h)
                
                # 2. Revisar turno asignado a AYER (Overflow)
                if status is None and yesterday in shifts and shifts[yesterday]:
                    s, e = shifts[yesterday]
                    if s > e: 
                        # Turno cruzaba medianoche (Ej: 21:00 ayer a 08:00 hoy)
                        # HOY es válido solo la madrugada (hasta la hora de fin)
                        if check_time < e:
                            status = get_initial_status(s, h)
                
                row[name] = status
            data.append(row)
        curr += timedelta(days=1)
    return pd.DataFrame(data)

# --- CÁLCULO DINÁMICO ---
def calculate_metrics_dynamic(df_sales, turnos, ordered_names, state_matrix, start_date, end_date):
    mask = (df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)
    sales_f = df_sales.loc[mask].copy()
    
    # Agrupación optimizada
    sales_f['d_str'] = sales_f['date'].dt.date.astype(str)
    sales_f['h'] = sales_f['date'].dt.hour
    sales_grouped = sales_f.groupby(['d_str', 'h'])['qt_price_local'].sum().to_dict()
    
    matriz_display = []
    daily_accum = {name: {} for name in ordered_names}
    stats_franjas = {name: {'Solo': 0, 'Con 1': 0, 'Con 2+': 0} for name in ordered_names}
    active_hours_count = {name: 0 for name in ordered_names}
    
    for _, row in state_matrix.iterrows():
        d_str = row['_date_str']
        h = row['_hour']
        
        # Leemos el estado desde la matriz editada por el usuario
        eligibles = [n for n in ordered_names if row[n] is True]
        fisicos = [n for n in ordered_names if pd.notna(row[n])]
        
        total_v = sales_grouped.get((d_str, h), 0)
        n = len(eligibles)
        monto = total_v / n if n > 0 else 0
        
        fila_vis = {'Día': d_str, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
        
        for i, name in enumerate(ordered_names):
            idx = i + 1
            if name in fisicos:
                if name in eligibles:
                    fila_vis[f'Coord {idx}'] = name
                    fila_vis[f'Venta C{idx}'] = round(monto)
                    
                    c_date = datetime.strptime(d_str, "%Y-%m-%d").date()
                    daily_accum[name][c_date] = daily_accum[name].get(c_date, 0) + monto
                    active_hours_count[name] += 1
                    
                    others = n - 1
                    if others == 0: stats_franjas[name]['Solo'] += 1
                    elif others == 1: stats_franjas[name]['Con 1'] += 1
                    else: stats_franjas[name]['Con 2+'] += 1
                else:
                    fila_vis[f'Coord {idx}'] = f"{name} (*)"
                    fila_vis[f'Venta C{idx}'] = 0
            else:
                fila_vis[f'Coord {idx}'] = ""
                fila_vis[f'Venta C{idx}'] = 0
        matriz_display.append(fila_vis)

    df_h_display = pd.DataFrame(matriz_display)
    
    # Resumen Diario
    daily_rows = []
    d_iter = start_date
    while d_iter <= end_date:
        r = {'Día': d_iter}
        for name in ordered_names:
            r[name] = round(daily_accum[name].get(d_iter, 0))
        daily_rows.append(r)
        d_iter += timedelta(days=1)
    df_daily = pd.DataFrame(daily_rows)
    
    # Totales
    total_metrics = []
    shared_stats_list = []
    
    for name in ordered_names:
        tot_v = sum(daily_accum[name].values())
        
        # Turnos Trabajados (Días distintos en el calendario)
        dias_trabajados = 0
        if name in turnos:
            for d, rng in turnos[name].items():
                if start_date <= d <= end_date and rng is not None:
                    dias_trabajados += 1
        
        horas_activas = active_hours_count[name]
        promedio = tot_v / horas_activas if horas_activas > 0 else 0
        
        total_metrics.append({
            'Coordinador': name,
            'Ventas Totales': round(tot_v),
            'Comisión (2%)': round(tot_v * 0.02),
            'Turnos Trabajados': dias_trabajados,
            'Horas Activas': horas_activas,
            'Promedio Venta/Hora': round(promedio)
        })
        
        shared_stats_list.append({
            'Coordinador': name,
            'Horas Solo': stats_franjas[name]['Solo'],
            'Horas con 1': stats_franjas[name]['Con 1'],
            'Horas con 2+': stats_franjas[name]['Con 2+']
        })
        
    return df_h_display, df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats_list)

# --- EXPORTAR EXCEL ---
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
