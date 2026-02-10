import pandas as pd
import numpy as np
import re
import warnings

# Configuración
pd.set_option('future.no_silent_downcasting', True)
warnings.simplefilter("ignore")

# CONSTANTES (Tus colores se usarán en app.py, aquí dejamos las llaves)
LLAVE_DIAN_CONT_COL_NAME = 'LLAVE_DIAN'
LLAVE_SERIE_FOLIO_COL_NAME = 'LLAVE_SERIE_FOLIO'

# COLORES PARA EXCEL (Necesarios para xlsxwriter)
CABIFY_PURPLE = '#7145D6'
CABIFY_LIGHT  = '#F3F0FA'
CABIFY_ACCENT = '#B89EF7'
WHITE         = '#FFFFFF'

# =================================================================
# 1. UTILIDADES Y LECTURA ADAPTADA PARA WEB
# =================================================================

def normalize_col_name(col_name):
    return re.sub(r'[^\w]+', '_', str(col_name)).lower().strip('_')

def standardize_company_name(name_series):
    if name_series is None or name_series.empty:
        return pd.Series(['SIN NOMBRE'] * len(name_series), dtype=str)
    return (
        name_series.astype(str)
        .str.upper()
        .str.replace(r'[^A-Z0-9\s]+', '', regex=True)
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
        .str.replace(r'\b(S A S|SAS|S A|SA|LTDA|LTDA|BIC|B I C)\b', '', regex=True) 
        .str.strip()
    )

def clean_nit_numeric(nit_series):
    if nit_series is None or nit_series.empty:
        return pd.Series([''] * len(nit_series), dtype=str)
    return nit_series.astype(str).str.replace(r'[^0-9]+', '', regex=True)

def limpiar_moneda_colombia(valor):
    if pd.isna(valor) or str(valor).strip() == '': return 0.0
    s = str(valor).replace('$', '').replace(' ', '')
    try:
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
            else: s = s.replace(',', '')
        elif ',' in s: 
            parts = s.split(',')
            if len(parts[-1]) == 2: s = s.replace(',', '.') 
            else: s = s.replace(',', '') 
        return float(s)
    except: return 0.0

def leer_contabilidad_completa(file_obj):
    """
    Lee el archivo de contabilidad desde un objeto en memoria (Streamlit).
    """
    if file_obj is None: return None
    try:
        # Pre-lectura para encontrar encabezado
        # En web, debemos resetear el puntero del archivo si lo leemos dos veces
        file_obj.seek(0) 
        try: df_preview = pd.read_excel(file_obj, nrows=20, header=None, engine='openpyxl')
        except: df_preview = pd.read_excel(file_obj, nrows=20, header=None)
        
        header_row = 0
        for i, row in df_preview.iterrows():
            row_str = row.astype(str).values
            if 'Cuenta' in row_str and 'Fecha' in row_str:
                header_row = i; break
        
        # Leemos el archivo real
        file_obj.seek(0) # Resetear puntero
        df = pd.read_excel(file_obj, header=header_row, engine='openpyxl')
        
        # Limpieza básica (Tu lógica original intacta)
        df['Cuenta'] = df['Cuenta'].astype(str).replace(['nan', 'None', ''], np.nan)
        condicion_cabecera = df['Cuenta'].str.match(r'^\d', na=False)
        df['CUENTA_COMPLETA'] = df['Cuenta'].where(condicion_cabecera, other=np.nan).ffill()
        df = df[~df['Cuenta'].astype(str).str.startswith('Total', na=False)] 
        df = df[df['Fecha'].notna()] 
        df['Cuenta'] = df['CUENTA_COMPLETA']
        
        df['CODIGO_CUENTA'] = df['Cuenta'].str.strip().str.extract(r'^(\d+)')
        
        if 'Fecha' in df.columns:
            df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce').dt.strftime('%Y-%m-%d')

        # Cálculo de Saldo
        col_deb = next((c for c in df.columns if 'Déb' in c), None)
        col_cred = next((c for c in df.columns if 'Créd' in c), None)
        
        val_deb = df[col_deb].apply(limpiar_moneda_colombia) if col_deb else 0.0
        val_cred = df[col_cred].apply(limpiar_moneda_colombia) if col_cred else 0.0
        
        df['SALDO_NETO_CALCULADO'] = val_deb - val_cred
        
        col_ref_orig = next((c for c in df.columns if 'mero de doc' in c or 'Nro' in c), 'Número de documento')
        col_nit_orig = next((c for c in df.columns if 'Identifi' in c or 'Nit' in c), 'Número Identificación')
        col_nom_orig = next((c for c in df.columns if 'Nombre' in c), 'Nombre')
        col_nota_orig = next((c for c in df.columns if 'Nota' in c), 'Nota')

        df_renamed = df.rename(columns={
            col_ref_orig: 'u_ref', 
            col_nit_orig: 'u_infoco01',
            col_nom_orig: 'u_cardname', 
            col_nota_orig: 'u_memo', 
            'Cuenta': 'u_acctname'
        })
        df_renamed['u_saldo_f'] = df['SALDO_NETO_CALCULADO']
        
        if 'u_infoco01' in df_renamed.columns:
            df_renamed['u_infoco01'] = df_renamed['u_infoco01'].astype(str).str.replace(r'\.0$', '', regex=True)

        return df_renamed
    except Exception as e: 
        print(f"Error: {e}")
        return None

def leer_dian(file_obj):
    if file_obj is None: return None
    try:
        df = pd.read_excel(file_obj, dtype=str)
        col_map = {col: normalize_col_name(col) for col in df.columns}
        df.rename(columns=col_map, inplace=True)
        return df
    except: return None

def leer_gosocket(file_obj):
    if file_obj is None: return None
    try:
        df = pd.read_excel(file_obj, dtype=str)
        col_map = {col: normalize_col_name(col) for col in df.columns}
        df.rename(columns=col_map, inplace=True)
        return df
    except: return None

# =================================================================
# 2. FILTROS Y MOTORES (Tu lógica original intacta)
# =================================================================

def crear_llave_conciliacion(df):
    cols = df.columns
    prefijo = next((c for c in cols if 'prefijo' in c), None)
    folio = next((c for c in cols if 'folio' in c), None)
    if not prefijo or not folio: return df
    df[LLAVE_DIAN_CONT_COL_NAME] = (
        df[prefijo].astype(str).str.strip() + 
        df[folio].astype(str).str.strip()
    ).str.replace(r'[^\w]+', '', regex=True).str.upper()
    return df

def crear_llave_serie_folio(df):
    cols_normalized = df.columns
    series_col = next((c for c in cols_normalized if 'serie' in c or 'prefijo' in c), None)
    folio_col = next((c for c in cols_normalized if 'folio' in c or 'numero' in c), None)
    
    if not series_col or not folio_col: return df, False
    
    try:
        df[LLAVE_SERIE_FOLIO_COL_NAME] = (
            df[series_col].astype(str).str.strip() + 
            df[folio_col].astype(str).str.strip()
        ).str.replace(r'[^\w]+', '', regex=True).str.upper()
        return df, True
    except: return df, False

def filtrar_solo_gastos(df_completo):
    if df_completo is None or df_completo.empty: return pd.DataFrame()
    df = df_completo.copy()
    df = df[df['CODIGO_CUENTA'].str.startswith('5', na=False)]
    df = df[df['CODIGO_CUENTA'] != '51157001']
    df = df[~df['u_acctname'].str.contains('IVA', case=False, na=False)]
    df = df[~df['u_acctname'].str.contains('DIFERENCIA EN CAMBIO', case=False, na=False)]
    df = df[~df['u_acctname'].str.contains('DEPRECIACI', case=False, na=False)]
    return df

def filtrar_solo_ingresos(df_completo):
    if df_completo is None or df_completo.empty: return pd.DataFrame()
    df = df_completo.copy()
    df = df[df['CODIGO_CUENTA'].str.startswith('4', na=False)]
    df = df[~df['u_acctname'].str.contains('DIFERENCIA EN CAMBIO', case=False, na=False)]
    df['u_saldo_f'] = df['u_saldo_f'] * -1 
    return df

def filtrar_solo_iva_descontable(df_completo):
    if df_completo is None or df_completo.empty: return pd.DataFrame()
    df = df_completo.copy()
    mask_base = df['CODIGO_CUENTA'].str.startswith('24', na=False) | df['u_acctname'].str.contains('IVA', case=False, na=False)
    df = df[mask_base]
    mask_exclude = df['u_acctname'].str.upper().str.contains('GENERADO|VENTA|DEVOLUCION VENTA', regex=True, na=False)
    df = df[~mask_exclude]
    return df

def filtrar_solo_iva_generado(df_completo):
    if df_completo is None or df_completo.empty: return pd.DataFrame()
    df = df_completo.copy()
    mask_base = df['CODIGO_CUENTA'].str.startswith('24', na=False) | df['u_acctname'].str.contains('IVA', case=False, na=False)
    df = df[mask_base]
    mask_gen = df['u_acctname'].str.upper().str.contains('GENERADO', regex=False, na=False)
    df = df[mask_gen]
    df['u_saldo_f'] = df['u_saldo_f'] * -1
    return df

def filtrar_dian_gastos(df):
    if df is None or df.empty: return df
    col_grupo = next((c for c in df.columns if 'grupo' in normalize_col_name(c)), None)
    col_tipo = next((c for c in df.columns if 'tipo' in normalize_col_name(c) and 'documento' in normalize_col_name(c)), None)
    if not col_grupo: return df
    mask_recibidos = df[col_grupo].astype(str).str.strip().str.lower() == 'recibido'
    mask_doc_soporte = pd.Series([False] * len(df))
    if col_tipo:
        s_tipo = df[col_tipo].astype(str).str.strip().str.lower()
        mask_doc_soporte = s_tipo.str.contains('documento soporte', na=False) | s_tipo.str.contains('no obligado', na=False)
    return df[mask_recibidos | mask_doc_soporte].copy()

def filtrar_dian_ingresos(df):
    if df is None or df.empty: return df
    col_grupo = next((c for c in df.columns if 'grupo' in normalize_col_name(c)), None)
    col_tipo = next((c for c in df.columns if 'tipo' in normalize_col_name(c) and 'documento' in normalize_col_name(c)), None)
    if not col_grupo: return df
    mask_emitidos = df[col_grupo].astype(str).str.strip().str.lower() == 'emitido'
    mask_doc_soporte = pd.Series([False] * len(df))
    if col_tipo:
        s_tipo = df[col_tipo].astype(str).str.strip().str.lower()
        mask_doc_soporte = s_tipo.str.contains('documento soporte', na=False) | s_tipo.str.contains('no obligado', na=False)
    return df[mask_emitidos & (~mask_doc_soporte)].copy()

def ejecutar_conciliacion_universal(df_dian, df_cont):
    if 'u_ref' not in df_cont.columns: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_cont['LLAVE_CONT'] = df_cont['u_ref'].astype(str).str.strip().str.replace(r'[^\w]+', '', regex=True).str.upper()
    
    agg_dict = {'u_saldo_f': 'sum'}
    if 'u_infoco01' in df_cont.columns: agg_dict['u_infoco01'] = 'first'
    if 'u_cardname' in df_cont.columns: agg_dict['u_cardname'] = 'first'
    if 'u_acctname' in df_cont.columns: agg_dict['u_acctname'] = 'first' 
    
    df_cont_agg = df_cont.groupby('LLAVE_CONT').agg(agg_dict).reset_index()
    
    if LLAVE_DIAN_CONT_COL_NAME not in df_dian.columns: return pd.DataFrame(), df_dian, df_cont
    
    df_coinc = pd.merge(df_dian, df_cont_agg, left_on=LLAVE_DIAN_CONT_COL_NAME, right_on='LLAVE_CONT', how='inner', suffixes=('_DIAN', '_CONT'))
    
    df_left = pd.merge(df_dian, df_cont_agg[['LLAVE_CONT']], left_on=LLAVE_DIAN_CONT_COL_NAME, right_on='LLAVE_CONT', how='left', indicator=True)
    df_sob_dian = df_left[df_left['_merge'] == 'left_only'].drop(columns=['LLAVE_CONT', '_merge'])
    
    df_right = pd.merge(df_cont_agg, df_dian[[LLAVE_DIAN_CONT_COL_NAME]], left_on='LLAVE_CONT', right_on=LLAVE_DIAN_CONT_COL_NAME, how='left', indicator=True)
    df_sob_cont = pd.merge(df_cont, df_right[df_right['_merge'] == 'left_only'][['LLAVE_CONT']], on='LLAVE_CONT', how='inner')
    
    return df_coinc, df_sob_dian, df_sob_cont

def conciliar_ingresos_vs_gosocket(df_ingresos, df_gosocket):
    if df_ingresos is None or df_ingresos.empty:
        return pd.DataFrame(), df_ingresos if df_ingresos is not None else pd.DataFrame(), pd.DataFrame()
        
    df_ingresos['LLAVE_CONC'] = df_ingresos['u_ref'].astype(str).str.strip().str.replace(r'[^\w]+', '', regex=True).str.upper()
    
    if LLAVE_SERIE_FOLIO_COL_NAME not in df_gosocket.columns:
        df_gosocket, _ = crear_llave_serie_folio(df_gosocket)
        
    if LLAVE_SERIE_FOLIO_COL_NAME not in df_gosocket.columns:
        col_ref = next((c for c in df_gosocket.columns if 'referencia' in c), None)
        if col_ref:
            df_gosocket[LLAVE_SERIE_FOLIO_COL_NAME] = df_gosocket[col_ref].astype(str).str.strip().str.replace(r'[^\w]+', '', regex=True).str.upper()
        else:
            return pd.DataFrame(), df_ingresos, df_gosocket

    agg_dict = {'u_saldo_f': 'sum'}
    if 'u_infoco01' in df_ingresos.columns: agg_dict['u_infoco01'] = 'first'
    if 'u_cardname' in df_ingresos.columns: agg_dict['u_cardname'] = 'first'
    df_ing_agg = df_ingresos.groupby('LLAVE_CONC').agg(agg_dict).reset_index()

    df_coinc = pd.merge(df_ing_agg, df_gosocket, left_on='LLAVE_CONC', right_on=LLAVE_SERIE_FOLIO_COL_NAME, how='inner', suffixes=('_CONT', '_GO'))
    
    df_left = pd.merge(df_ing_agg, df_gosocket[[LLAVE_SERIE_FOLIO_COL_NAME]], left_on='LLAVE_CONC', right_on=LLAVE_SERIE_FOLIO_COL_NAME, how='left', indicator=True)
    df_sob_cont_agg = df_left[df_left['_merge'] == 'left_only']
    df_sob_cont = pd.merge(df_ingresos, df_sob_cont_agg[['LLAVE_CONC']], on='LLAVE_CONC', how='inner')

    df_right = pd.merge(df_gosocket, df_ing_agg[['LLAVE_CONC']], left_on=LLAVE_SERIE_FOLIO_COL_NAME, right_on='LLAVE_CONC', how='left', indicator=True)
    df_sob_go = df_right[df_right['_merge'] == 'left_only'].drop(columns=['LLAVE_CONC', '_merge'])

    return df_coinc, df_sob_cont, df_sob_go

# =================================================================
# 3. REPORT GENERATION (LOGIC ONLY)
# =================================================================

def formato_cabezote_cabify(workbook):
    return workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top',
        'fg_color': CABIFY_PURPLE, 'font_color': WHITE,
        'border': 1, 'align': 'center'
    })

def formatear_hoja_base(writer, sheet_name, df):
    if df.empty: return
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    worksheet.set_tab_color('green')
    fmt_header = formato_cabezote_cabify(workbook)
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, fmt_header)
    worksheet.set_column(0, len(df.columns) - 1, 15)

def procesar_reporte_cabify_generico(coin, sob_d, sob_c, writer, sheet_name, emisor_col, total_col, iva_col, is_iva_report=False):
    lista_dfs = []
    
    # 1. COINCIDENCIAS
    if not coin.empty:
        t = coin.copy()
        t['NIT'] = clean_nit_numeric(t['u_infoco01'])
        t['EMPRESA'] = t[emisor_col] if emisor_col in t.columns else t['u_cardname']
        t['EMPRESA_GRUPO'] = standardize_company_name(t['EMPRESA'])
        
        val_d = pd.to_numeric(t[total_col], errors='coerce').fillna(0)
        
        if not is_iva_report:
            if iva_col and iva_col in t.columns: 
                val_d -= pd.to_numeric(t[iva_col], errors='coerce').fillna(0)
        
        t['SUBTOTAL DIAN'] = val_d
        t['TOTAL CONTABILIDAD'] = t['u_saldo_f']
        t['DIFERENCIA'] = t['SUBTOTAL DIAN'] - t['TOTAL CONTABILIDAD']
        t['TIPO'] = 'COINCIDENCIA'
        t['LLAVE_DIAN'] = t[LLAVE_DIAN_CONT_COL_NAME]
        t['LLAVE_CONT'] = t['LLAVE_CONT']
        t['CUENTA_CONTABLE'] = t['u_acctname'] if 'u_acctname' in t.columns else ''
        lista_dfs.append(t)

    # 2. SOBRANTES DIAN
    if not sob_d.empty:
        t = sob_d.copy()
        col_nit = next((c for c in t.columns if ('emisor' in c or 'receptor' in c) and ('nit' in c or 'doc' in c)), None)
        t['NIT'] = clean_nit_numeric(t[col_nit]) if col_nit else ''
        t['EMPRESA'] = t[emisor_col] if emisor_col in t.columns else 'DESCONOCIDO'
        t['EMPRESA_GRUPO'] = standardize_company_name(t['EMPRESA'])
        
        val_d = pd.to_numeric(t[total_col], errors='coerce').fillna(0)
        if not is_iva_report:
            if iva_col and iva_col in t.columns: 
                val_d -= pd.to_numeric(t[iva_col], errors='coerce').fillna(0)
            
        t['SUBTOTAL DIAN'] = val_d
        t['TOTAL CONTABILIDAD'] = 0
        t['DIFERENCIA'] = val_d
        t['TIPO'] = 'SOBRANTE_DIAN'
        t['LLAVE_DIAN'] = t[LLAVE_DIAN_CONT_COL_NAME]
        t['LLAVE_CONT'] = ''
        t['CUENTA_CONTABLE'] = ''
        lista_dfs.append(t)

    # 3. SOBRANTES CONTABILIDAD
    if not sob_c.empty:
        t = sob_c.copy()
        t['NIT'] = clean_nit_numeric(t['u_infoco01'])
        t['EMPRESA'] = t['u_cardname']
        t['EMPRESA_GRUPO'] = standardize_company_name(t['u_cardname'])
        
        t['SUBTOTAL DIAN'] = 0
        t['TOTAL CONTABILIDAD'] = t['u_saldo_f']
        t['DIFERENCIA'] = -t['u_saldo_f']
        t['TIPO'] = 'SOBRANTE_CONT'
        t['LLAVE_DIAN'] = ''
        t['LLAVE_CONT'] = t['LLAVE_CONT'] if 'LLAVE_CONT' in t.columns else t['u_ref']
        t['CUENTA_CONTABLE'] = t['u_acctname'] if 'u_acctname' in t.columns else ''
        lista_dfs.append(t)

    cols_visibles = ['NIT', 'EMPRESA', 'LLAVE_DIAN', 'LLAVE_CONT', 'CUENTA_CONTABLE', 'SUBTOTAL DIAN', 'TOTAL CONTABILIDAD', 'DIFERENCIA', 'TIPO', 'GRUPO_TIPO']

    if not lista_dfs: 
        pd.DataFrame(columns=[c for c in cols_visibles if c != 'GRUPO_TIPO']).to_excel(writer, sheet_name=sheet_name, index=False)
        writer.sheets[sheet_name].set_tab_color(CABIFY_PURPLE)
        return

    df_full = pd.concat(lista_dfs, ignore_index=True)
    df_full = df_full[(df_full['SUBTOTAL DIAN'].abs() > 1) | (df_full['TOTAL CONTABILIDAD'].abs() > 1)].copy()
    
    if df_full.empty:
        pd.DataFrame(columns=[c for c in cols_visibles if c != 'GRUPO_TIPO']).to_excel(writer, sheet_name=sheet_name, index=False)
        writer.sheets[sheet_name].set_tab_color(CABIFY_PURPLE)
        return

    tipo_orden = {'COINCIDENCIA': 1, 'SOBRANTE_DIAN': 2, 'SOBRANTE_CONT': 3}
    df_full['ORDEN'] = df_full['TIPO'].map(tipo_orden)
    df_full.sort_values(by=['EMPRESA_GRUPO', 'NIT', 'ORDEN'], inplace=True)

    final_rows = []
    cols_sum = ['SUBTOTAL DIAN', 'TOTAL CONTABILIDAD', 'DIFERENCIA']
    
    for empresa, df_emp in df_full.groupby('EMPRESA_GRUPO', sort=False):
        nit_grupo = df_emp['NIT'].iloc[0] if not df_emp.empty else ''
        for tipo, df_tipo in df_emp.groupby('TIPO', sort=False):
            df_tipo['GRUPO_TIPO'] = 'DETALLE'
            final_rows.append(df_tipo)
            sums = df_tipo[cols_sum].sum()
            row_sub = pd.Series({'NIT': nit_grupo, 'EMPRESA': empresa, 'TIPO': f'SUBTOTAL {tipo}', 'GRUPO_TIPO': 'SUBTOTAL_TIPO', **sums})
            final_rows.append(row_sub.to_frame().T)
        sums_emp = df_emp[cols_sum].sum()
        row_emp = pd.Series({'NIT': nit_grupo, 'EMPRESA': f'TOTAL {empresa}', 'GRUPO_TIPO': 'SUBTOTAL_EMPRESA', **sums_emp})
        final_rows.append(row_emp.to_frame().T)

    if not final_rows:
        pd.DataFrame(columns=[c for c in cols_visibles if c != 'GRUPO_TIPO']).to_excel(writer, sheet_name=sheet_name, index=False)
        return

    df_out = pd.concat(final_rows, ignore_index=True).fillna('')
    sums_glob = df_out[df_out['GRUPO_TIPO'] == 'SUBTOTAL_EMPRESA'][cols_sum].sum()
    row_glob = pd.Series({'EMPRESA': 'GRAN TOTAL GLOBAL', 'GRUPO_TIPO': 'GRAN_TOTAL', **sums_glob})
    df_out = pd.concat([df_out, row_glob.to_frame().T], ignore_index=True).fillna('')

    cols_final = [c for c in cols_visibles if c in df_out.columns]
    df_write = df_out[cols_final].copy()
    
    df_write.drop(columns=['GRUPO_TIPO']).to_excel(writer, sheet_name=sheet_name, index=False)
    
    wb = writer.book
    ws = writer.sheets[sheet_name]
    
    fmt_header = formato_cabezote_cabify(wb)
    fmt_sub_tipo_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT})
    fmt_sub_tipo_num = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT, 'num_format': '#,##0.00'})
    fmt_sub_emp_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE})
    fmt_sub_emp_num = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 'num_format': '#,##0.00'})
    fmt_total_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE})
    fmt_total_num = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE, 'num_format': '#,##0.00'})

    try:
        idx_sub_dian = df_write.columns.get_loc('SUBTOTAL DIAN')
        idx_tot_cont = df_write.columns.get_loc('TOTAL CONTABILIDAD')
        idx_dif = df_write.columns.get_loc('DIFERENCIA')
        cols_moneda = [idx_sub_dian, idx_tot_cont, idx_dif]
    except:
        cols_moneda = []

    total_cols_count = len(df_write.columns) - 1
    for col_num, value in enumerate(df_write.drop(columns=['GRUPO_TIPO']).columns.values):
        ws.write(0, col_num, value, fmt_header)
    ws.set_column('B:B', 40)
    if cols_moneda:
        ws.set_column(cols_moneda[0], cols_moneda[-1], 18, wb.add_format({'num_format': '#,##0.00'}))

    for i, row in df_write.iterrows():
        tipo_fila = row['GRUPO_TIPO']
        excel_row = i + 1 
        if tipo_fila == 'DETALLE':
            ws.set_row(excel_row, None, None, {'level': 2, 'hidden': True})
        else:
            f_txt, f_num = None, None
            if tipo_fila == 'SUBTOTAL_TIPO':
                f_txt, f_num = fmt_sub_tipo_txt, fmt_sub_tipo_num
                ws.set_row(excel_row, None, None, {'level': 1, 'hidden': False})
            elif tipo_fila == 'SUBTOTAL_EMPRESA':
                f_txt, f_num = fmt_sub_emp_txt, fmt_sub_emp_num
                ws.set_row(excel_row, None, None, {'level': 0, 'collapsed': False})
            elif tipo_fila == 'GRAN_TOTAL':
                f_txt, f_num = fmt_total_txt, fmt_total_num

            if f_txt and f_num:
                for col_idx in range(total_cols_count):
                    valor = row.iloc[col_idx]
                    estilo = f_num if col_idx in cols_moneda else f_txt
                    ws.write(excel_row, col_idx, valor, estilo)
    ws.set_tab_color(CABIFY_PURPLE)