import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# --- UTILIDADES ---
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

# --- REGLAS BASE (SOLO PARA ESTADO INICIAL) ---
def get_initial_status(shift_start, h_curr):
    """Retorna True si está HABILITADO PARA VENTA por defecto, False si es hora de Loza."""
    if shift_start is None: return False
    h_start = shift_start.hour
    
    # Reglas de exclusión (Si cae aquí, retorna False = No Vende/Loza)
    if h_start == 10:
        if h_curr == 10 or (14 <= h_curr < 16): return False
    elif h_start == 5:
        if 11 <= h_curr < 14: return False
    elif h_start == 21:
        if 6 <= h_curr < 9: return False
        
    return True # Por defecto vende

# --- CARGA DE DATOS (SOLO UNA VEZ) ---
def load_data_once(sales_file, turnos_file):
    # 1. Cargar Turnos
    if hasattr(turnos_file, 'name') and turnos_file.name.endswith('.xlsx'):
        df_turnos = pd.read_excel(turnos_file, header=None)
    else:
        df_turnos = pd.read_csv(turnos_file, header=None, encoding='latin1', sep=None, engine='python')
    
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

    # 2. Cargar Ventas
    if hasattr(sales_file, 'name') and sales_file.name.endswith('.xlsx'):
        df_sales = pd.read_excel(sales_file)
    else:
        try:
            df_sales = pd.read_csv(sales_file, encoding='utf-8', sep=None, engine='python')
        except:
            sales_file.seek(0)
            df_sales = pd.read_csv(sales_file, encoding='latin-1', sep=None, engine='python')
    
    # Normalizar columnas
    df_sales.columns = [c.strip() for c in df_sales.columns]
    col_map = {c: 'date' for c in df_sales.columns if 'created' in c.lower() or 'fecha' in c.lower() or 'date' in c.lower()}
    if col_map: df_sales.rename(columns=col_map, inplace=True)
    
    if 'date' not in df_sales.columns: return None, None, None # Error

    df_sales['date'] = pd.to_datetime(df_sales['date'])
    
    return df_sales, turnos_dict, ordered_names

# --- GENERACIÓN DE MATRIZ DE ESTADO INICIAL ---
def generate_initial_state_matrix(turnos, ordered_names, start_date, end_date):
    """Crea un DataFrame que representa el estado 'VENDIENDO' (True/False) para cada slot."""
    data = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_time = time(h, 0)
            row = {'date_idx': str(curr), 'hour_idx': h} # Indices para cruce
            
            yesterday = curr - timedelta(days=1)
            
            for name in ordered_names:
                shifts = turnos.get(name, {})
                status = None # None significa NO ESTÁ DE TURNO
                
                # Check Turno Hoy
                if curr in shifts and shifts[curr]:
                    s, e = shifts[curr]
                    if (s < e and s <= check_time < e) or (s > e and (check_time >= s or check_time < e)):
                        status = get_initial_status(s, h) # True (Venta) o False (Loza)
                
                # Check Turno Ayer
                if status is None and yesterday in shifts and shifts[yesterday]:
                    s, e = shifts[yesterday]
                    if s > e and check_time < e:
                        status = get_initial_status(s, h)
                
                row[name] = status
            data.append(row)
        curr += timedelta(days=1)
    
    return pd.DataFrame(data)

# --- CÁLCULO PRINCIPAL (SIMULACIÓN) ---
def calculate_metrics_dynamic(df_sales, turnos, ordered_names, state_matrix, start_date, end_date):
    # Filtrar ventas por fecha
    mask = (df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)
    sales_filtered = df_sales.loc[mask].copy()
    
    # Agrupar ventas por hora para velocidad
    sales_filtered['d_str'] = sales_filtered['date'].dt.date.astype(str)
    sales_filtered['h'] = sales_filtered['date'].dt.hour
    sales_grouped = sales_filtered.groupby(['d_str', 'h'])['qt_price_local'].sum().to_dict()
    
    # Mapeo de columnas
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(ordered_names)}
    
    matriz_display = []
    daily_accum = {name: {} for name in ordered_names}
    stats_franjas = {name: {'Solo': 0, 'Con 1': 0, 'Con 2+': 0} for name in ordered_names}
    
    # Iterar sobre la matriz de estados (que ya tiene las modificaciones del usuario)
    for _, row_state in state_matrix.iterrows():
        d_str = row_state['date_idx']
        h = row_state['hour_idx']
        
        # Obtener quiénes están VENDIENDO (True) en este slot según la matriz editada
        eligibles = [name for name in ordered_names if row_state[name] is True]
        # Obtener quiénes están DE TURNO (True o False, no None)
        fisicos = [name for name in ordered_names if pd.notna(row_state[name])]
        
        # Distribuir Venta
        total_venta = sales_grouped.get((d_str, h), 0)
        n = len(eligibles)
        monto_per_capita = total_venta / n if n > 0 else 0
        
        # Construir Fila Visual
        fila_vis = {'Día': d_str, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
        
        for name in ordered_names:
            idx = mapa_cols[name]
            col_name_c = f'Coord {idx}'
            col_sales_c = f'Venta C{idx}'
            
            if name in fisicos:
                if name in eligibles:
                    # Actividad 1: Vendiendo
                    fila_vis[col_name_c] = name
                    fila_vis[col_sales_c] = round(monto_per_capita)
                    
                    # Acumular
                    curr_date = datetime.strptime(d_str, "%Y-%m-%d").date()
                    daily_accum[name][curr_date] = daily_accum[name].get(curr_date, 0) + monto_per_capita
                    
                    # Stats
                    others = n - 1
                    if others == 0: stats_franjas[name]['Solo'] += 1
                    elif others == 1: stats_franjas[name]['Con 1'] += 1
                    else: stats_franjas[name]['Con 2+'] += 1
                    
                else:
                    # Actividad 2: Loza/Colación
                    fila_vis[col_name_c] = f"{name} (*)"
                    fila_vis[col_sales_c] = 0
            else:
                fila_vis[col_name_c] = ""
                fila_vis[col_sales_c] = 0
                
        matriz_display.append(fila_vis)
        
    # --- CONSTRUIR RESULTADOS ---
    df_hourly_display = pd.DataFrame(matriz_display)
    
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
    
    # Totales + Comisión
    total_metrics = []
    shared_stats_list = []
    
    for name in ordered_names:
        tot_v = sum(daily_accum[name].values())
        
        # Turnos (Días trabajados) - Se calcula desde el diccionario original de turnos
        # (Esto es estático, no depende de la simulación de loza)
        dias_trabajados = 0
        if name in turnos:
            for d, rng in turnos[name].items():
                if start_date <= d <= end_date and rng is not None:
                    dias_trabajados += 1
                    
        total_metrics.append({
            'Coordinador': name,
            'Ventas Totales': round(tot_v),
            'Comisión (2%)': round(tot_v * 0.02),
            'Turnos Trabajados': dias_trabajados
        })
        
        shared_stats_list.append({
            'Coordinador': name,
            'Horas Solo': stats_franjas[name]['Solo'],
            'Horas con 1': stats_franjas[name]['Con 1'],
            'Horas con 2+': stats_franjas[name]['Con 2+']
        })
        
    return df_hourly_display, df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats_list)

# --- EXPORTAR EXCEL ---
def generate_styled_excel(dfs_dict):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    header_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'white', 'bg_color': '#7145D6', 'border': 0, 'align': 'center', 'valign': 'vcenter'})
    cell_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#333333', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center', 'valign': 'vcenter'})
    special_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#9C4A00', 'bg_color': '#FFEB9C', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center'})

    for sheet_name, df in dfs_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
            ws.set_column(idx, idx, max_len, cell_fmt)
            ws.write(0, idx, col, header_fmt)
            if sheet_name == 'Matriz_Horaria':
                ws.conditional_format(1, idx, len(df), idx, {'type': 'text', 'criteria': 'containing', 'value': '(*)', 'format': special_fmt})
        ws.hide_gridlines(2)
    writer.close()
    return output.getvalue()
