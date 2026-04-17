import pandas as pd
import sys
import json
import os

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

def normalizar_columnas_unicas(columns):
    usadas = {}
    resultado = []
    for idx, col in enumerate(columns):
        nombre = str(col).strip() if pd.notna(col) and str(col).strip() and not str(col).startswith('Unnamed') else f"Col_{idx}"
        if nombre in usadas:
            usadas[nombre] += 1
            nombre = f"{nombre}_{usadas[nombre]}"
        else:
            usadas[nombre] = 0
        resultado.append(nombre)
    return resultado

def score_sheet_for_primary(name, df):
    name_l = str(name).lower()
    rows, cols = df.shape
    score = rows * cols
    if cols < 2 or rows < 2:
        return -1
    if any(k in name_l for k in ['muestra total', 'base de comisiones', 'consolidado total']):
        score += 500
    if any(k in name_l for k in ['ventas', 'inventario', 'stock', 'productos', 'resumen', 'dashboard', 'datos']):
        score += 180
    if any(k in name_l for k in ['hallazgo', 'oportunidad', 'mejora', 'coso', 'td', 'distribucion']):
        score -= 200
    headers = [str(c).strip().lower() for c in df.columns if pd.notna(c)]
    if any(h in headers for h in ['solicitante', 'valor total solicitado', 'ventas_totales', 'precio_venta', 'stock']):
        score += 120
    return score

def extract_real_sheet(df):
    try:
        raw_data = [df.columns.tolist()] + df.values.tolist()
        best_row_idx = 0
        max_valid = 0
        for i, row in enumerate(raw_data[:20]):
            valid_cols = [str(x) for x in row if pd.notna(x) and str(x).strip() and not str(x).startswith('Unnamed') and 'TITLE:' not in str(x) and 'TYPE:' not in str(x) and 'SUBTITLE:' not in str(x)]
            if len(valid_cols) > max_valid:
                max_valid = len(valid_cols)
                best_row_idx = i
                
        if max_valid >= 2:
            new_header = raw_data[best_row_idx]
            new_data = raw_data[best_row_idx+1:]
            df_new = pd.DataFrame(new_data, columns=new_header)
            df_new.columns = normalizar_columnas_unicas(df_new.columns)
            df_new = df_new.dropna(axis=1, how='all').dropna(axis=0, how='all')
            return df_new
    except: pass
    return df

def limpiar_df(df):
    df = df.copy()
    df.columns = normalizar_columnas_unicas(df.columns)
    for idx, col in enumerate(df.columns):
        serie = df.iloc[:, idx]
        col_str = str(col).lower()
        if 'fecha' in col_str or 'date' in col_str or 'mes' in col_str:
            try:
                if serie.dtype == 'float64' or serie.dtype == 'int64':
                    df.iloc[:, idx] = pd.to_datetime(serie, unit='D', origin='1899-12-30').dt.strftime('%d/%m/%Y')
            except:
                pass
        if serie.dtype == object or serie.dtype == 'string':
            df.iloc[:, idx] = serie.fillna('—')
        else:
            df.iloc[:, idx] = serie.fillna(0)
    return df

def remover_filas_basura(df):
    df = df.copy()
    df.columns = normalizar_columnas_unicas(df.columns)
    palabras_basura = ['diapositiva', 'power point', 'powerpoint', 
                       'agrega', 'información', 'columna', 'imagen adjunta', 'placeholder']
    mask = pd.Series([True] * len(df), index=df.index)
    for idx, col in enumerate(df.columns):
        serie = df.iloc[:, idx]
        if serie.dtype == object:
            for palabra in palabras_basura:
                mask &= ~serie.astype(str).str.lower().str.contains(palabra, na=False)
    df = df[mask]
    c_importantes = ['Id Comisión', 'Solicitante', 'Valor Total Solicitado']
    cols_check = [c for c in c_importantes if c in df.columns]
    if cols_check:
        df = df[df[cols_check].notna().any(axis=1)]
        if 'Id Comisión' in df.columns:
            df = df[df['Id Comisión'].notna()]
    return df

def leer_coso(excel_path):
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_name = None
        for sn in xl.sheet_names:
            if 'coso' in sn.lower():
                sheet_name = sn
                break
        if not sheet_name: return None
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=3)
        header_row_idx = 0
        for i in range(min(10, len(df))):
            fila_valores = df.iloc[i].tolist()
            fila_str = " ".join([str(v).lower() for v in fila_valores])
            if any(k in fila_str for k in ['componente', 'item', 'evalua', 'punto de control', 'control']):
                header_row_idx = i
                break
        df = df.iloc[header_row_idx:].copy()
        while df.shape[1] > 0 and df.iloc[0, 0] is None:
            df = df.iloc[:, 1:]
        if df.shape[1] >= 3:
            df.columns = ['Componente', 'Item', 'Estado'] + [f'Col_{i}' for i in range(df.shape[1]-3)]
            df['Componente'] = df['Componente'].ffill()
            df = df[df['Estado'].notna()]
            df = df[~df['Componente'].astype(str).str.lower().str.contains('componente|evalua', na=False)]
            df['Item'] = df['Item'].astype(str).str[:120]
            return {
                'encabezados': ['Componente', 'Ítems Evaluados', 'Estado'],
                'filas': df[['Componente', 'Item', 'Estado']].values.tolist()
            }
        return None
    except Exception as e: 
        print(f"Error leer_coso: {e}", file=sys.stderr)
        return None

def leer_distribucion_mes(excel_path):
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_name = None
        for sn in xl.sheet_names:
            if 'td' == sn.lower().strip() or 'distribucion' in sn.lower():
                sheet_name = sn
                break
        if not sheet_name: return None
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all').iloc[1:]
        if df.shape[1] > 0:
            df.columns = ['Centro de Costos'] + [f'C{i}' for i in range(df.shape[1]-1)]
            df_slide = df[['Centro de Costos']].dropna().head(12)
            return {
                'encabezados': ['Centro de Costos'],
                'filas': df_slide.values.tolist()
            }
        return None
    except: return None

def preparar_datos_para_slides(excel_path):
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        return {"error": str(e)}

    resultado = {}
    resultado['metadatos'] = {
        'hojas_encontradas': list(sheets.keys()),
        'archivo': os.path.basename(excel_path)
    }
    
    for key in sheets.keys():
        sheets[key] = extract_real_sheet(sheets[key])

    # === BUSCAR HOJA PRINCIPAL (Más robusto) ===
    target_sheet = None
    
    if sheets:
        best_sheet = None
        best_score = -1
        for name, df in sheets.items():
            current_score = score_sheet_for_primary(name, df)
            if current_score > best_score:
                best_score = current_score
                best_sheet = name
        target_sheet = best_sheet

    processed_sheets = set()
    if target_sheet:
        processed_sheets.add(target_sheet)
        df = sheets[target_sheet]
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        es_comisiones = all(col in df.columns for col in ['Solicitante', 'Valor Total Solicitado'])
        
        # === RESUMEN EJECUTIVO ===
        total_registros = len(df)
        if es_comisiones:
            valor_total = 0
            if 'Valor Total Solicitado' in df.columns:
                valor_total = float(pd.to_numeric(df['Valor Total Solicitado'], errors='coerce').sum())
            
            unique_solicitantes = 0
            if 'Solicitante' in df.columns:
                unique_solicitantes = int(df['Solicitante'].nunique())
            
            unique_ciudades = 0
            if 'Ciudad Destino' in df.columns:
                unique_ciudades = int(df['Ciudad Destino'].nunique())
            
            unique_centros = 0
            if 'Centro de Costos' in df.columns:
                unique_centros = int(df['Centro de Costos'].astype(str).str.strip().nunique())
                valid_cc = df[df['Centro de Costos'].astype(str).str.strip().str.len() > 1]
                unique_centros = int(valid_cc['Centro de Costos'].nunique())
            
            promedio_comision = valor_total / total_registros if total_registros > 0 else 0
            
            valor_max = 0
            if 'Valor Total Solicitado' in df.columns:
                valor_max = float(pd.to_numeric(df['Valor Total Solicitado'], errors='coerce').max())
            
            resultado['resumen_ejecutivo'] = {
                'total_comisiones': total_registros,
                'valor_total': valor_total,
                'unique_solicitantes': unique_solicitantes,
                'unique_ciudades': unique_ciudades,
                'unique_centros': unique_centros,
                'promedio_comision': promedio_comision,
                'valor_max_comision': valor_max
            }
        else:
            cols_numericas = []
            for col in df.columns:
                serie_num = pd.to_numeric(df[col], errors='coerce')
                if serie_num.notna().sum() >= max(2, int(len(df) * 0.6)):
                    cols_numericas.append(col)
            resultado['resumen_generico'] = {
                'hoja_principal': target_sheet,
                'total_filas': total_registros,
                'total_columnas': int(df.shape[1]),
                'columnas_numericas': cols_numericas[:8],
                'columnas': df.columns.tolist()[:12]
            }
        
        # === TABLA PRINCIPAL (más columnas) ===
        cols = ['Id Comisión','Solicitante','Ciudad Destino',
                'Valor Total Solicitado','Estado','Centro de Costos']
        cols_exist = [c for c in cols if c in df.columns]
        if not cols_exist:
             cols_exist = df.columns[:7].tolist()
        
        df_slide = df[cols_exist].copy()
        for col in ['Solicitante','Ciudad Destino']:
            if col in df_slide.columns:
                df_slide[col] = df_slide[col].astype(str).str[:50]
        
        resultado['muestra_tabla'] = {
            'encabezados': cols_exist,
            'filas': df_slide.values.tolist()
        }
        
        # === GRÁFICA ESTADOS ===
        if es_comisiones and 'Estado' in df.columns:
            estados = df['Estado'].value_counts().head(8)
            resultado['grafica_estados'] = {
                'tipo': 'pie',
                'titulo': 'Distribución por Estado',
                'labels': estados.index.tolist(),
                'valores': estados.values.tolist(),
                'colores': ['1E3A5F','4472C4','70AD47','ED7D31','FF0000','FFC000','9B59B6','3498DB']
            }
        
        # === GRÁFICA VALORES POR TIPO DE GASTO ===
        cols_valores = {
            'Tiquete': 'Valor Tiquete Solicitado',
            'Alimentación': 'Valor Alimentación Solicitado', 
            'Hospedaje': 'Valor Hospedaje Solicitado',
            'Transporte': 'Valor Transporte Solicitado'
        }
        vals = {}
        for nombre, col in cols_valores.items():
            if col in df.columns:
                vals[nombre] = float(pd.to_numeric(df[col], errors='coerce').sum())
        
        if es_comisiones and vals:
            resultado['grafica_valores'] = {
                'tipo': 'bar',
                'titulo': 'Total por Tipo de Gasto (COP)',
                'labels': list(vals.keys()),
                'valores': list(vals.values()),
                'colores': ['4472C4','ED7D31','A9D18E','FFC000']
            }
        
        # === TOP CIUDADES ===
        if es_comisiones and 'Ciudad Destino' in df.columns:
            ciudades = df['Ciudad Destino'].value_counts().head(8)
            resultado['grafica_ciudades'] = {
                'tipo': 'bar',
                'titulo': 'Top Ciudades de Destino',
                'labels': [str(c)[:25] for c in ciudades.index.tolist()],
                'valores': ciudades.values.tolist(),
                'colores': ['4472C4']
            }
        
        # === TOP SOLICITANTES POR VALOR ===
        if es_comisiones and 'Solicitante' in df.columns and 'Valor Total Solicitado' in df.columns:
            df_vals = df.copy()
            df_vals['Valor Total Solicitado'] = pd.to_numeric(df_vals['Valor Total Solicitado'], errors='coerce').fillna(0)
            top_sol = df_vals.groupby('Solicitante').agg(
                total_valor=('Valor Total Solicitado', 'sum'),
                num_comisiones=('Id Comisión', 'count')
            ).sort_values('total_valor', ascending=False).head(8)
            
            resultado['top_solicitantes'] = {
                'labels': [str(s)[:30] for s in top_sol.index.tolist()],
                'valores': [float(v) for v in top_sol['total_valor'].tolist()],
                'conteos': top_sol['num_comisiones'].tolist()
            }
        
        # === DISTRIBUCIÓN POR CENTRO DE COSTOS ===
        if es_comisiones and 'Centro de Costos' in df.columns and 'Valor Total Solicitado' in df.columns:
            df_cc = df.copy()
            df_cc['Centro de Costos'] = df_cc['Centro de Costos'].astype(str).str.strip()
            df_cc['Valor Total Solicitado'] = pd.to_numeric(df_cc['Valor Total Solicitado'], errors='coerce').fillna(0)
            cc_top = df_cc.groupby('Centro de Costos')['Valor Total Solicitado'].sum().sort_values(ascending=False).head(8)
            
            resultado['centros_costos'] = {
                'labels': cc_top.index.tolist(),
                'valores': [float(v) for v in cc_top.values.tolist()]
            }
    
    # === Hallazgos y Oportunidades ===
    otras_tablas = {}
    for name, df in sheets.items():
        n_lower = name.lower()
        if 'hallazgo' in n_lower or 'oportunidad' in n_lower or 'mejora' in n_lower:
            df = remover_filas_basura(df)
            df = limpiar_df(df)
            if not df.empty:
               df_res = df.head(50).copy()
               for col in df_res.columns:
                   if df_res[col].dtype == object:
                       df_res[col] = df_res[col].astype(str).str[:200]
                       
               # Extract progress data if % column exists
               progress_data = None
               for col in df_res.columns:
                   col_str = str(col).strip()
                   if col_str == '%' or 'porcentaje' in col_str.lower() or 'avance' in col_str.lower():
                       try:
                           progress_vals = pd.to_numeric(df_res[col], errors='coerce').fillna(0)
                           progress_data = progress_vals.tolist()
                       except:
                           pass
                       break
               
               tabla_info = {
                   'encabezados': df_res.columns.tolist()[:6],
                   'filas': df_res.values.tolist()
               }
               if progress_data:
                   tabla_info['progress'] = progress_data
                   
               otras_tablas[name] = tabla_info
    
    if otras_tablas: 
        resultado['otras_tablas'] = otras_tablas

    # === HOJAS RESTANTES (Cualquier tabla no procesada) ===
    genericas = {}
    for name, df in sheets.items():
        if name in processed_sheets or name == target_sheet:
            continue
        
        # Saltarse hojas que probablemente no sean tablas (muy pequeñas o vacías)
        if df.empty or df.shape[1] < 2 or df.shape[0] < 2:
            continue
            
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        if not df.empty:
            filled_ratio = df.notna().sum().sum() / max(1, df.shape[0] * df.shape[1])
            if filled_ratio < 0.25:
                continue
            df_res = df.head(30).copy()
            for col in df_res.columns:
                if df_res[col].dtype == object:
                    df_res[col] = df_res[col].astype(str).str[:150]
            
            genericas[name] = {
                'encabezados': df_res.columns.tolist()[:8],
                'filas': df_res.values.tolist()
            }
            
    if genericas:
        resultado['genericas'] = genericas

    # === COSO y TD ===
    coso = leer_coso(excel_path)
    if coso: resultado['coso'] = coso
    
    td = leer_distribucion_mes(excel_path)
    if td: resultado['distribucion_mes'] = td

    return resultado

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "No file path provided"}))
        sys.exit(1)
    
    import warnings
    warnings.filterwarnings('ignore')
    
    path_excel = sys.argv[1]
    if not os.path.exists(path_excel):
        print(json.dumps({"error": f"File not found: {path_excel}"}))
        sys.exit(1)
        
    try:
        data = preparar_datos_para_slides(path_excel)
        print(json.dumps(data, ensure_ascii=False, default=str))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False, default=str))
        sys.exit(1)
