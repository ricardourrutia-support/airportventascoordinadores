import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

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
    """Lee archivos Excel/CSV. has_header=True para Ventas, False para Turnos."""
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
    # Turnos se lee SIN header porque la estructura es compleja (fechas en fila 1)
    df_raw = read_file_generic(file, has_header=False)
    
    # Fila 1 tiene las fechas (índice 1 en dataframe 0-indexado)
    actual_dates = pd.to_datetime(df_raw.iloc[1, 1:], errors='coerce').dt.date.tolist()
    
    turnos_data = {}
    ordered_names = []
    
    # Los datos parten desde la fila 2
    for i in range(2, len(df_raw)):
        row = df_raw.iloc[i]
        nombre = str(row.iloc[0]).strip()
        if nombre.upper() in ["NAN", "", "NOMBRE"]: continue
        
        # Guardamos el nombre tal cual para respetar el orden del archivo
        # (Sin ordenar alfabéticamente)
        if nombre not in ordered_names:
            ordered_names.append(nombre)
            
        dias = {f: parse_turno_range(row.iloc[j+1]) for j, f in enumerate(actual_dates) if f is not pd.NaT}
        turnos_data[nombre] = dias
        
    return turnos_data, ordered_names

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

# --- GENERADOR DE EXCEL ESTILO CABIFY ---
def generate_styled_excel(dfs_dict):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    # Estilos Corporativos Minimalistas
    header_fmt = workbook.add_format({
        'bold': True, 'font_name': 'Arial', 'font_size': 10,
        'font_color': 'white', 'bg_color': '#7145D6', # Cabify Purple
        'border': 0, 'align': 'center', 'valign': 'vcenter'
    })
    
    cell_fmt = workbook.add_format({
        'font_name': 'Arial', 'font_size': 10, 'font_color': '#333333',
        'border': 0, 'bottom': 1, 'bottom_color': '#E0E0E0', # Línea sutil
        'align': 'center', 'valign': 'vcenter'
    })

    name_col_fmt = workbook.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True,
        'font_color': '#7145D6', 'bg_color': '#F9F9F9',
        'bottom': 1, 'bottom_color': '#E0E0E0', 'align': 'left'
    })

    for sheet_name, df in dfs_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        for idx, col in enumerate(df.columns):
            # Ancho automático basado en contenido
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
            ws.set_column(idx, idx, max_len, cell_fmt)
            
            # Primera columna destacada (nombres/fechas)
            if idx == 0:
                ws.set_column(idx, idx, max_len + 2, name_col_fmt)
            
            ws.write(0, idx, col, header_fmt)
        
        ws.hide_gridlines(2) # Ocultar líneas de cuadrícula

    writer.close()
    return output.getvalue()

def process_all(sales_file, turnos_file, start_date, end_date):
    turnos, ordered_names = load_turnos(turnos_file)
    
    # Leemos ventas CON header=0 (estándar)
    df_sales = read_file_generic(sales_file, has_header=True)
    
    # Normalización de columnas
    df_sales.columns = [c.strip() for c in df_sales.columns]
    
    # Renombrar createdAt_local -> date
    if 'createdAt_local' in df_sales.columns:
        df_sales.rename(columns={'createdAt_local': 'date'}, inplace=True)
    
    # Validación crítica
    if 'date' not in df_sales.columns:
        # Intento desesperado de encontrar columna de fecha
        for col in df_sales.columns:
            if 'date' in col.lower() or 'created' in col.lower() or 'fecha' in col.lower():
                df_sales.rename(columns={col: 'date'}, inplace=True)
                break
        if 'date' not in df_sales.columns:
            return None, None, None, None # Error controlado
            
    df_sales['date'] = pd.to_datetime(df_sales['date'])
    df_sales = df_sales[(df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)].copy()

    # Mapeo usando ORDEN DEL ARCHIVO
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(ordered_names)}
    
    # Cálculo de ventas
    ventas_calc = []
    # Pre-cálculo de mapeo de turnos para optimizar
    # (Aunque con pocos datos el loop directo funciona bien)
    
    for _, row in df_sales.iterrows():
        activos = get_active_coordinators(row['date'], turnos)
        n = len(activos)
        if n > 0:
            for name in activos:
                ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': name, 'v': row['qt_price_local']/n})
        else:
            ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': 'SIN ASIGNAR', 'v': row['qt_price_local']})
            
    df_v = pd.DataFrame(ventas_calc)
    if df_v.empty:
        # Crear DF vacío con columnas esperadas para no romper el código
        df_v = pd.DataFrame(columns=['fecha', 'hora', 'coord', 'v'])

    # 1. Matriz Horaria
    matriz_data = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            activos_h = get_active_coordinators(check_dt, turnos)
            
            fila = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00'}
            # Auxiliares para cálculo de franjas
            fila['_count'] = len(activos_h)
            fila['_activos'] = activos_h
            
            for nom in ordered_names:
                idx = mapa_cols[nom]
                if nom in activos_h:
                    fila[f'Coord {idx}'] = nom
                    if not df_v.empty:
                        val = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['coord']==nom)]['v'].sum()
                        fila[f'Venta C{idx}'] = round(val)
                    else:
                        fila[f'Venta C{idx}'] = 0
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
            if not df_v.empty:
                val = df_v[(df_v['fecha']==curr) & (df_v['coord']==nom)]['v'].sum()
            else: val = 0
            r[nom] = round(val)
        daily_rows.append(r)
        curr += timedelta(days=1)
    df_daily = pd.DataFrame(daily_rows)

    # 3. Métricas Totales y Franjas
    total_metrics = []
    shared_stats = []
    
    for nom in ordered_names:
        # Ventas
        if not df_v.empty:
            tot_v = df_v[df_v['coord']==nom]['v'].sum()
        else: tot_v = 0
        
        # Turnos trabajados
        dias_trabajados = 0
        if nom in turnos:
            for d, rng in turnos[nom].items():
                if start_date <= d <= end_date and rng is not None:
                    dias_trabajados += 1
        
        total_metrics.append({
            'Coordinador': nom, 
            'Ventas Totales': round(tot_v), 
            'Días Trabajados': dias_trabajados
        })
        
        # Franjas
        solo = 0; con1 = 0; con2 = 0
        # Usamos df_hourly que ya tiene la info calculada
        for _, row_h in df_hourly.iterrows():
            if nom in row_h['_activos']:
                others = row_h['_count'] - 1
                if others == 0: solo += 1
                elif others == 1: con1 += 1
                else: con2 += 1
        
        shared_stats.append({
            'Coordinador': nom,
            'Hrs Solo': solo,
            'Hrs con 1': con1,
            'Hrs con 2+': con2
        })

    # Limpiar columnas auxiliares antes de devolver
    df_hourly_clean = df_hourly.drop(columns=['_count', '_activos'])
    
    return df_hourly_clean, df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats)
