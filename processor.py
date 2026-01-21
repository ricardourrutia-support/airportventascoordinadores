import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# --- 1. CONFIGURACIÓN DE REGLAS DE GESTIÓN ---
def is_in_special_hours(shift_start_time, current_time):
    """
    Determina si una hora específica es 'Horario de Gestión/Colación' 
    basado en la hora de inicio del turno.
    Retorna True si NO debe recibir ventas (está en loza/colación).
    """
    h_start = shift_start_time.hour
    h_curr = current_time.hour
    
    # Regla 1: Turno de 10:00 (Gestión 10-11 y 14-16)
    if h_start == 10:
        if h_curr == 10: return True          # 10:00 a 10:59
        if 14 <= h_curr < 16: return True     # 14:00 a 15:59
        
    # Regla 2: Turno de 05:00 (Gestión 11-14)
    elif h_start == 5:
        if 11 <= h_curr < 14: return True     # 11:00 a 13:59
        
    # Regla 3: Turno de 21:00 (Gestión 06-09 am del día siguiente o actual)
    elif h_start == 21:
        if 6 <= h_curr < 9: return True       # 06:00 a 08:59
        
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

def get_active_coordinators(sale_dt, turnos):
    """
    Retorna lista de tuplas: (Nombre, HoraInicioTurno)
    """
    s_date, s_time = sale_dt.date(), sale_dt.time()
    yesterday = s_date - timedelta(days=1)
    
    active_details = [] # [(Nombre, StartTime), ...]
    
    for name, shifts in turnos.items():
        # Turno Hoy
        if s_date in shifts and shifts[s_date]:
            start, end = shifts[s_date]
            if (start < end and start <= s_time < end) or (start > end and (s_time >= start or s_time < end)):
                active_details.append((name, start))
                
        # Turno Ayer (Overflow)
        if yesterday in shifts and shifts[yesterday]:
            start, end = shifts[yesterday]
            if start > end and s_time < end:
                active_details.append((name, start))
                
    # Eliminamos duplicados si hubiese cruce raro (set de tuplas)
    return list(set(active_details))

# --- GENERADOR DE EXCEL CON FORMATO CONDICIONAL ---
def generate_styled_excel(dfs_dict):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    # Formatos
    header_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'white', 'bg_color': '#7145D6', 'border': 0, 'align': 'center', 'valign': 'vcenter'})
    cell_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#333333', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center', 'valign': 'vcenter'})
    name_col_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#7145D6', 'bg_color': '#F9F9F9', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'left'})
    
    # Formato especial para "En Loza" (Color Naranja suave)
    special_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'font_color': '#9C6500', 'bg_color': '#FFEB9C', 'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'center'})

    for sheet_name, df in dfs_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
            ws.set_column(idx, idx, max_len, cell_fmt)
            
            if idx == 0:
                ws.set_column(idx, idx, max_len + 2, name_col_fmt)
            
            ws.write(0, idx, col, header_fmt)
            
            # Aplicar formato condicional para celdas con (*)
            if sheet_name == 'Matriz_Horaria':
                ws.conditional_format(1, idx, len(df), idx, {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': '(*)',
                    'format': special_fmt
                })

        ws.hide_gridlines(2)

    writer.close()
    return output.getvalue()

def process_all(sales_file, turnos_file, start_date, end_date):
    turnos, ordered_names = load_turnos(turnos_file)
    df_sales = read_file_generic(sales_file, has_header=True)
    
    # Normalización Header
    df_sales.columns = [c.strip() for c in df_sales.columns]
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
    df_sales = df_sales[(df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)].copy()

    mapa_cols = {nombre: i+1 for i, nombre in enumerate(ordered_names)}
    
    ventas_calc = []
    
    # --- PROCESO DE ASIGNACIÓN CON REGLAS DE GESTIÓN ---
    for _, row in df_sales.iterrows():
        # 1. Obtener quiénes están físicamente
        activos_fisicos_info = get_active_coordinators(row['date'], turnos) # Lista de (Nombre, StartTime)
        
        # 2. Filtrar quiénes pueden vender (NO están en Loza/Colación)
        eligibles_para_venta = []
        for name, start_time in activos_fisicos_info:
            if not is_in_special_hours(start_time, row['date'].time()):
                eligibles_para_venta.append(name)
        
        n = len(eligibles_para_venta)
        
        if n > 0:
            for name in eligibles_para_venta:
                ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': name, 'v': row['qt_price_local']/n})
        else:
            ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': 'SIN ASIGNAR', 'v': row['qt_price_local']})
            
    df_v = pd.DataFrame(ventas_calc)
    if df_v.empty: df_v = pd.DataFrame(columns=['fecha', 'hora', 'coord', 'v'])

    # 1. Matriz Horaria
    matriz_data = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            
            # Info Física vs Info Venta
            activos_fisicos_tuples = get_active_coordinators(check_dt, turnos)
            nombres_fisicos = [x[0] for x in activos_fisicos_tuples]
            
            # Quiénes estaban habilitados para vender en esta hora (aprox, checkeamos con min 00)
            nombres_eligibles = []
            for name, start_t in activos_fisicos_tuples:
                if not is_in_special_hours(start_t, check_dt.time()):
                    nombres_eligibles.append(name)
            
            fila = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            
            # Auxiliares para estadística de competencia (Solo contamos si estaba ELIGIBLE)
            fila['_count_eligible'] = len(nombres_eligibles)
            fila['_eligibles'] = nombres_eligibles
            
            for nom in ordered_names:
                idx = mapa_cols[nom]
                if nom in nombres_fisicos:
                    # Si está físico pero NO eligible (está en Loza) -> Marcamos con (*)
                    if nom not in nombres_eligibles:
                        fila[f'Coord {idx}'] = f"{nom} (*)"
                        fila[f'Venta C{idx}'] = 0 # No recibe venta
                    else:
                        fila[f'Coord {idx}'] = nom
                        val = 0
                        if not df_v.empty:
                            val = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['coord']==nom)]['v'].sum()
                        fila[f'Venta C{idx}'] = round(val)
                else:
                    fila[f'Coord {idx}'] = ""
                    fila[f'Venta C{idx}'] = 0
            matriz_data.append(fila)
        curr += timedelta(days=1)
    
    df_hourly = pd.DataFrame(matriz_data)
    
    # 2. Resumen Diario
    daily_rows = []
    curr = start_date
    while curr <= end_date:
        r = {'Día': curr}
        for nom in ordered_names:
            val = df_v[(df_v['fecha']==curr) & (df_v['coord']==nom)]['v'].sum() if not df_v.empty else 0
            r[nom] = round(val)
        daily_rows.append(r)
    df_daily = pd.DataFrame(daily_rows)

    # 3. Totales y Franjas
    total_metrics = []
    shared_stats = []
    
    for nom in ordered_names:
        tot_v = df_v[df_v['coord']==nom]['v'].sum() if not df_v.empty else 0
        
        dias_trabajados = 0
        if nom in turnos:
            for d, rng in turnos[nom].items():
                if start_date <= d <= end_date and rng is not None:
                    dias_trabajados += 1
        
        total_metrics.append({'Coordinador': nom, 'Ventas Totales': round(tot_v), 'Turnos Trabajados': dias_trabajados})
        
        # Franjas (Basadas en competencia de venta)
        solo = 0; con1 = 0; con2 = 0
        for _, row_h in df_hourly.iterrows():
            if nom in row_h['_eligibles']: # Solo cuenta si estaba vendiendo
                others = row_h['_count_eligible'] - 1
                if others == 0: solo += 1
                elif others == 1: con1 += 1
                else: con2 += 1
        
        shared_stats.append({'Coordinador': nom, 'Solo (Horas Venta)': solo, 'Con 1 (Horas Venta)': con1, 'Con 2+ (Horas Venta)': con2})

    df_h_clean = df_hourly.drop(columns=['_count_eligible', '_eligibles'])
    return df_h_clean, df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats)
