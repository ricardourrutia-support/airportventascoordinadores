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

def read_file_safely(file):
    """Lee el archivo intentando detectar el formato correcto."""
    if hasattr(file, 'name') and file.name.endswith('.xlsx'):
        return pd.read_excel(file, header=None)
    else:
        try:
            return pd.read_csv(file, header=None, encoding='utf-8', sep=None, engine='python')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, header=None, encoding='latin-1', sep=None, engine='python')

def load_turnos(file):
    df_raw = read_file_safely(file)
    # Fila 1 tiene las fechas (índice 1)
    actual_dates = pd.to_datetime(df_raw.iloc[1, 1:], errors='coerce').dt.date.tolist()
    
    turnos_data = {}
    # Los datos parten desde la fila 2 (índice 2)
    # IMPORTANTE: No ordenamos alfabéticamente después, respetamos este orden de inserción
    for i in range(2, len(df_raw)):
        row = df_raw.iloc[i]
        nombre = str(row.iloc[0]).strip() # Quitamos .upper() para mantener formato original si deseas
        if nombre.upper() in ["NAN", "", "NOMBRE"]: continue
        
        # Guardamos tal cual viene en el archivo (ej: Jocsana Lopez)
        dias = {f: parse_turno_range(row.iloc[j+1]) for j, f in enumerate(actual_dates) if f is not pd.NaT}
        turnos_data[nombre.upper()] = dias # Usamos llave mayúscula para consistencia interna
    return turnos_data, list(turnos_data.keys()) # Devolvemos la lista ordenada original

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

    # Formatos Estilo Cabify
    header_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 10,
        'font_color': 'white',
        'bg_color': '#7145D6', # Cabify Purple
        'border': 0,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    cell_fmt = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'font_color': '#333333',
        'border': 0, # Minimalista: sin bordes internos duros
        'bottom': 1, # Solo líneas horizontales suaves
        'bottom_color': '#E0E0E0',
        'align': 'center',
        'valign': 'vcenter'
    })

    coord_col_fmt = workbook.add_format({
        'font_name': 'Arial', 
        'font_size': 10,
        'bold': True,
        'font_color': '#7145D6', # Morado para nombres
        'bg_color': '#F8F8F8',
        'bottom': 1,
        'bottom_color': '#E0E0E0',
        'align': 'left'
    })

    for sheet_name, df in dfs_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        
        # Ajustar ancho de columnas y aplicar formato
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 4
            worksheet.set_column(idx, idx, max_len, cell_fmt)
            
            # Formato especial para primera columna si es texto
            if idx == 0:
                worksheet.set_column(idx, idx, max_len + 5, coord_col_fmt)
            
            # Escribir cabecera con formato
            worksheet.write(0, idx, col, header_fmt)
        
        # Ocultar líneas de cuadrícula para look limpio
        worksheet.hide_gridlines(2)

    writer.close()
    return output.getvalue()

def process_all(sales_file, turnos_file, start_date, end_date):
    turnos, ordered_names = load_turnos(turnos_file)
    
    # Manejo robusto de ventas y columna createdAt_local
    df_sales = read_file_safely(sales_file)
    if 'createdAt_local' in df_sales.columns:
        df_sales.rename(columns={'createdAt_local': 'date'}, inplace=True)
    
    df_sales['date'] = pd.to_datetime(df_sales['date'])
    df_sales = df_sales[(df_sales['date'].dt.date >= start_date) & (df_sales['date'].dt.date <= end_date)].copy()

    # Mapeo Fijo usando el ORDEN DEL ARCHIVO (no alfabético)
    mapa_cols = {nombre: i+1 for i, nombre in enumerate(ordered_names)}
    
    ventas_calc = []
    for _, row in df_sales.iterrows():
        activos = get_active_coordinators(row['date'], turnos)
        n = len(activos)
        if n > 0:
            for name in activos:
                ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': name, 'v': row['qt_price_local']/n, 'asignado': True})
        else:
            ventas_calc.append({'fecha': row['date'].date(), 'hora': row['date'].hour, 'coord': 'SIN ASIGNAR', 'v': row['qt_price_local'], 'asignado': False})
    
    df_v = pd.DataFrame(ventas_calc)

    # 1. Matriz Horaria
    matriz_data = []
    curr = start_date
    while curr <= end_date:
        for h in range(24):
            check_dt = datetime.combine(curr, time(h, 0))
            activos_h = get_active_coordinators(check_dt, turnos)
            fila_h = {'Día': curr, 'Tramo': f'{h:02d}:00 - {h+1:02d}:00', 'Count': len(activos_h), 'Activos': activos_h}
            
            for nom in ordered_names: # Usamos el orden del archivo
                idx = mapa_cols[nom]
                if nom in activos_h:
                    fila_h[f'Coordinador {idx}'] = nom
                    # Ventas
                    if not df_v.empty:
                        v_h = df_v[(df_v['fecha']==curr) & (df_v['hora']==h) & (df_v['coord']==nom)]['v'].sum()
                        fila_h[f'Venta C{idx}'] = round(v_h)
                    else:
                        fila_h[f'Venta C{idx}'] = 0
                else:
                    fila_h[f'Coordinador {idx}'] = ""
                    fila_h[f'Venta C{idx}'] = 0
            matriz_data.append(fila_h)
        curr += timedelta(days=1)
    
    df_hourly = pd.DataFrame(matriz_data)
    
    # 2. Resumen Diario
    daily_rows = []
    curr = start_date
    while curr <= end_date:
        row = {'Día': curr}
        for nom in ordered_names:
            idx = mapa_cols[nom]
            if not df_v.empty:
                v = df_v[(df_v['fecha']==curr) & (df_v['coord']==nom)]['v'].sum()
            else: v = 0
            row[f'{nom} (C{idx})'] = round(v)
        daily_rows.append(row)
        curr += timedelta(days=1)
    df_daily = pd.DataFrame(daily_rows)

    # 3. Totales y Franjas
    total_metrics = []
    shared_stats = []
    
    for nom in ordered_names:
        if not df_v.empty:
            v = df_v[df_v['coord']==nom]['v'].sum()
        else: v = 0
        
        # Contar días trabajados
        worked_days = 0
        if nom in turnos:
            for d in turnos[nom]:
                if start_date <= d <= end_date and turnos[nom][d] is not None:
                    worked_days += 1
        
        total_metrics.append({'Coordinador': nom, 'Ventas Totales': round(v), 'Turnos Trabajados': worked_days})
        
        # Franjas
        solo = 0; con1 = 0; con2plus = 0
        for _, r in df_hourly.iterrows():
            if nom in r['Activos']:
                others = r['Count'] - 1
                if others == 0: solo += 1
                elif others == 1: con1 += 1
                else: con2plus += 1
        shared_stats.append({'Coordinador': nom, 'Solo (Horas)': solo, 'Con 1 (Horas)': con1, 'Con 2+ (Horas)': con2plus})

    return df_hourly.drop(columns=['Count', 'Activos']), df_daily, pd.DataFrame(total_metrics), pd.DataFrame(shared_stats)
