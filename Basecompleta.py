# -*- coding: utf-8 -*-
"""
Genera Sugerencias_Traslados_Proyecto.xlsx con las hojas:
- Traslados
- Stock_despues
- Cobertura_before_after
- Resumen_tienda
- Validaciones

Criterios clave:
- SKU = "Referencia-Talla"
- Min por SKU: 2 (tiendas físicas), 3 (e-commerce)
- Máximo por SKU destino: 5
- Curvas:
  * RANGO 5301/BEBES: tallas ['0M','3M','6M','9M','12M','18M'], mínimo 3 tallas con ≥2 u. c/u
  * RANGO 5302/NIÑOS: tallas ['2T','3T','4T','5T','6','8','10','12'], mínimo 5 tallas con ≥2 u. c/u
- Cobertura (ADU) desde Ventas: ADU = (unidades totales) / max(días distintos con venta, días entre min y max fecha)
- Origen debe quedar con cobertura mínima: 7 días (físicas) / 10 (e-com)
- Destino objetivo de cobertura: 7 días (físicas) / 10 (e-com)
- Bodega principal: solo envía, nunca recibe (puede quedar en 0)
- Selección de origen:
    1) MISMA REGIÓN primero
    2) PRIORIDAD (1–6) desde matriz dirección ORIGEN→DESTINO (menor = más cercano)
    3) Mayor cobertura del origen
    4) Menor ETA (días) ORIGEN→DESTINO
- La matriz de tiempos **NO es simétrica** (A→B != B→A)
"""

import pandas as pd
import numpy as np
import re
from collections import defaultdict

#When attempting to access a key that does not exist in a defaultdict, it automatically creates that key
# ---------------------
# Rutas (ajusta si es necesario)
# ---------------------

PATH_STOCK = "PruebaStockIA.xlsx"              # contiene hojas de ventas y stock (nombres detectados automáticamente)
PATH_TIEN  = "Clasificacion_Tiendas.xlsx"      # opcional: para Región/Ciudad (y/o Clasificación)
PATH_TENT  = "Tiempos de entrega.xlsx"         # opcional: matriz direccional con ORIGEN-DESTINO, DESTINO-ORIGEN, ETA, PRIORIDAD
OUT_PATH   = "Sugerencias_Traslados_Proyecto.xlsx"

# ---------------------
# Parámetros de negocio
# ---------------------
MIN_POR_SKU = 2
MIN_ECOM    = 3
MAX_STOCK_PER_SKU = 5

CURVAS_TALLAS = {
    'BEBES': ['0M','3M','6M','9M','12M','18M'],
    'NIÑOS': ['2T','3T','4T','5T','6','8','10','12']
}
CURVAS_MIN_TALLAS = {'BEBES':3, 'NIÑOS':5}

# Consideremos unificar las 2 lineas porque es igual para tiendas/ecommerce

ORIGIN_MIN_COV_DAYS, DEST_TARGET_COV_DAYS = 7, 7
ORIGIN_MIN_COV_ECOM, DEST_TARGET_COV_ECOM = 7, 7

COV_BUFFER_DAYS = 1  # evitar traslados si la cobertura del origen no es > cobertura del destino + 1 día

#Mejor incorporar filtro desde la extracción de datos

TIENDAS_CERRADAS = ["MEDELLIN FABRICATO", "CALI MALL PLAZA", "CARTAGENA MALLPLAZA"]
PALABRAS_EXCLUIR = ["FOTO", "FOTOG", "ARREG", "PRODUC", "IMPERF"]

# ---------------------
# Utilidades
# ---------------------

#Normaliza nombres de columnas: minúscula + sin saltos de línea.
def std_col(s):
    return str(s).strip().lower().replace("\n"," ")

# Limpia espacios y lo pasa a MAYÚSCULAS. Para joins y filtros con nombre tienda
# Nah, its better to use store and CO Ids
def normalize_store_name(x):
    return str(x).strip().upper() if pd.notna(x) else x

#Detecta e-commerce por palabras clave (ECO, ONLINE, WEB…).
def looks_like_ecom(name):
    s = normalize_store_name(name)
    return any(k in s for k in ["ECO","ECOM","ONLINE","VIRTUAL","WEB"])

#Marca tiendas a excluir por lista y palabras clave.
#Nah, otra cosa que se hace en la extracción

def is_excluded_store(name):
    s = normalize_store_name(name)
    return any(s == c for c in TIENDAS_CERRADAS) or any(k in s for k in PALABRAS_EXCLUIR)

# Sugerencia de mejora pues sku no esta separado por caracteres (Parametricemos eso)
def split_ref_talla_from_sku(sku):
    if pd.isna(sku): return (None, None)
    s = str(sku)
    if "-" in s:
        a,b = s.rsplit("-",1); return a.strip(), b.strip()
    if " " in s:
        a,b = s.rsplit(" ",1); return a.strip(), b.strip()
    return s, None

# Detectar columna de fecha del archivo
def detect_date_column(df):
    for c in df.columns:
        sc = std_col(c)
        if 'fecha' in sc or 'fec' in sc or 'f. doc' in sc or 'documento' in sc:
            try:
                if pd.to_datetime(df[c], errors='coerce').notna().any():
                    return c
            except:
                pass
    return None

# Ajusta las categorias en la columna RANGO
def canonical_rango(x):
    s = str(x).strip().upper()
    if ('5301' in s) or ('BEBE' in s): return 'BEBES'
    if ('5302' in s) or ('NIÑ' in s) or ('NIN' in s): return 'NIÑOS'
    return None

# Busca bodega principal por nombre y por mayor existencia
def pick_bodega_principal(stock_df):
    tmp = stock_df.groupby('Tienda', as_index=False)['Existencia'].sum()
    tmp['score'] = tmp['Tienda'].str.contains('BODEGA|CEDI|PRINCIPAL', case=False, regex=True).astype(int)
    tmp = tmp.sort_values(['score','Existencia'], ascending=[False, False])
    return tmp.iloc[0]['Tienda'] if len(tmp) else None

# Esto es para el archivo de tiempos de entrega. Usemos la prioridad que creamos 
def parse_lead_time_value(x):
    """Soporta valores como: 1, 2, 3, '3 a 4', 'A2', '1 A 2' → toma el máximo."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip().upper()
    nums = [int(n) for n in re.findall(r'\d+', s)]
    if not nums:
        return np.nan
    return max(nums)

# ---------------------
# Carga de archivos base
# ---------------------
xl = pd.ExcelFile(PATH_STOCK)
m = {std_col(n): n for n in xl.sheet_names}
ventas_sheet = next((m[k] for k in m if ('ventasref' in k or 'venta' in k or 'ventas' in k or 'sales' in k)), None)
stock_sheet  = next((m[k] for k in m if ('stockref' in k or 'stock' in k or 'invent' in k)), None)
if ventas_sheet is None or stock_sheet is None:
    raise ValueError("No se encontraron hojas de Ventas/Stock en el archivo base.")
ventas = xl.parse(ventas_sheet)
stock  = xl.parse(stock_sheet)

# Clasificación/Región/Ciudad (opcional)
try:
    ct = pd.read_excel(PATH_TIEN)
    ct_cols = {std_col(c): c for c in ct.columns}
    col_tienda_ct = ct_cols.get('tienda') or list(ct.columns)[0]
    col_region_ct = ct_cols.get('region') or ct_cols.get('región')
    col_ciudad_ct = ct_cols.get('ciudad')
    ct_df = ct[[col_tienda_ct] + ([col_region_ct] if col_region_ct else []) + ([col_ciudad_ct] if col_ciudad_ct else [])].copy()
    new_cols = ['Tienda']
    if col_region_ct: new_cols.append('Region')
    if col_ciudad_ct: new_cols.append('Ciudad')
    ct_df.columns = new_cols
    ct_df['Tienda'] = ct_df['Tienda'].apply(normalize_store_name)
    tiendas_map = ct_df.set_index('Tienda').to_dict(orient='index')
except Exception:
    tiendas_map = {}

# Tiempos de entrega (direccional, con PRIORIDAD)
try:
    tiempos_df = pd.read_excel(PATH_TENT)
    cols = {std_col(c): c for c in tiempos_df.columns}
    # Encabezados flexibles
    col_o   = cols.get('origen-destino') or cols.get('origen') or list(tiempos_df.columns)[0]
    col_d   = cols.get('destino-origen') or cols.get('destino') or list(tiempos_df.columns)[1]
    col_eta = cols.get('eta') or cols.get('dias') or cols.get('días') or cols.get('tiempo (dias)') or list(tiempos_df.columns)[-2]
    col_pri = next((cols[k] for k in cols if 'priorid' in k or 'priori' in k), None)

    tiempos_df = tiempos_df.copy()
    tiempos_df['_O']        = tiempos_df[col_o].astype(str).str.upper().str.strip()
    tiempos_df['_D']        = tiempos_df[col_d].astype(str).str.upper().str.strip()
    tiempos_df['_ETA_NUM']  = tiempos_df[col_eta].apply(parse_lead_time_value)
    tiempos_df['_PRI_NUM']  = pd.to_numeric(tiempos_df[col_pri], errors='coerce') if col_pri else np.nan
except Exception:
    tiempos_df = pd.DataFrame()

def delivery_priority(origin_store, dest_store, tiempos_df):
    """PRIORIDAD 1–6 (direccional A->B). Si no hay dato, NaN."""
    if tiempos_df is None or tiempos_df.empty:
        return np.nan
    oc = str(origin_store).strip().upper()
    dc = str(dest_store).strip().upper()
    m = tiempos_df[(tiempos_df['_O'] == oc) & (tiempos_df['_D'] == dc)]
    if len(m):
        val = m.iloc[0]['_PRI_NUM']
        return float(val) if pd.notna(val) else np.nan
    return np.nan

def delivery_days(origin_city, dest_city, tiempos_df):
    """ETA en días (direccional A->B). Si no hay dato, NaN."""
    if origin_city is None or dest_city is None or tiempos_df is None or tiempos_df.empty:
        return np.nan
    oc = str(origin_city).strip().upper()
    dc = str(dest_city).strip().upper()
    m = tiempos_df[(tiempos_df['_O'] == oc) & (tiempos_df['_D'] == dc)]
    if len(m):
        val = m.iloc[0]['_ETA_NUM']
        return float(val) if pd.notna(val) else np.nan
    return np.nan

# ---------------------
# Normalización de stock
# ---------------------
stc = {std_col(c): c for c in stock.columns}
c_sku   = stc.get('sku')
c_exist = stc.get('existencia') or stc.get('existencias')
c_tnd   = stc.get('desc. bodega') or stc.get('bodega') or stc.get('tienda')
c_ref   = stc.get('referencia')
c_talla = stc.get('talla')
c_rango = stc.get('rango')
if c_sku is None or c_exist is None or c_tnd is None:
    raise ValueError("Faltan columnas en Stock: SKU / Existencia / Desc. Bodega|Tienda")

stock_df = stock[[c_sku, c_exist, c_tnd] + ([c_ref] if c_ref else []) + ([c_talla] if c_talla else []) + ([c_rango] if c_rango else [])].copy()
stock_df.columns = ['SKU','Existencia','Tienda'] + (['Referencia'] if c_ref else []) + (['Talla'] if c_talla else []) + (['RANGO_RAW'] if c_rango else [])

# Si no vienen referencia/talla, derivar del SKU
if 'Referencia' not in stock_df.columns or 'Talla' not in stock_df.columns:
    ref_parsed, talla_parsed = zip(*stock_df['SKU'].apply(split_ref_talla_from_sku).tolist())
    if 'Referencia' not in stock_df.columns: stock_df['Referencia'] = ref_parsed
    if 'Talla' not in stock_df.columns:      stock_df['Talla']      = talla_parsed

stock_df['Tienda']     = stock_df['Tienda'].apply(normalize_store_name)
stock_df['Existencia'] = pd.to_numeric(stock_df['Existencia'], errors='coerce').fillna(0).astype(int)
stock_df               = stock_df[~stock_df['Tienda'].apply(is_excluded_store)].copy()
stock_df['RANGO_CAT']  = stock_df.get('RANGO_RAW', np.nan).apply(canonical_rango)

# Filtrar por tallas válidas (solo BEBES/NIÑOS con sus curvas)
valid_bebes = set(CURVAS_TALLAS['BEBES'])
valid_ninos = set(CURVAS_TALLAS['NIÑOS'])
mask_valid = ((stock_df['RANGO_CAT']=='BEBES') & (stock_df['Talla'].isin(valid_bebes))) | \
             ((stock_df['RANGO_CAT']=='NIÑOS') & (stock_df['Talla'].isin(valid_ninos)))
stock_df = stock_df[mask_valid].copy()

# ---------------------
# ADU (desde ventas) + cobertura informativa
# ---------------------
vtc   = {std_col(c): c for c in ventas.columns}
v_sku = vtc.get('sku')
v_qty = vtc.get('cantidad inv.') or vtc.get('cantidad') or vtc.get('unidades') or vtc.get('cantidad vendida')
v_tnd = vtc.get('desc. bodega') or vtc.get('desc. c.o.') or vtc.get('tienda') or vtc.get('bodega')
v_date = detect_date_column(ventas)

if v_sku and v_qty and v_tnd:
    v = ventas[[v_sku, v_qty, v_tnd] + ([v_date] if v_date else [])].copy()
    v.columns = ['SKU','Unidades','Tienda'] + (['Fecha'] if v_date else [])
    v['Tienda']   = v['Tienda'].apply(normalize_store_name)
    v['Unidades'] = pd.to_numeric(v['Unidades'], errors='coerce').fillna(0.0)

    # Filtrar ventas por tallas válidas (derivando talla del SKU)
    v[['Ref_tmp','Talla_tmp']] = v['SKU'].apply(lambda s: pd.Series(split_ref_talla_from_sku(s)))
    valid_union = valid_bebes.union(valid_ninos)
    v = v[v['Talla_tmp'].isin(valid_union)].copy()

    if 'Fecha' in v.columns:
        v['Fecha'] = pd.to_datetime(v['Fecha'], errors='coerce')
        v = v[v['Fecha'].notna()].copy()
        v['Dia'] = v['Fecha'].dt.date
        agg = v.groupby(['Tienda','SKU']).agg(
            total_units   = ('Unidades','sum'),
            fecha_min     = ('Dia','min'),
            fecha_max     = ('Dia','max'),
            dias_distintos= ('Dia','nunique')
        ).reset_index()
        agg['dias_rango'] = (pd.to_datetime(agg['fecha_max']) - pd.to_datetime(agg['fecha_min'])).dt.days.clip(lower=1)
        denom = agg[['dias_distintos','dias_rango']].max(axis=1).replace(0,1)
        agg['ADU'] = agg['total_units'] / denom
    else:
        agg = v.groupby(['Tienda','SKU']).agg(total_units=('Unidades','sum')).reset_index()
        agg['ADU'] = agg['total_units'] / 30.0
else:
    agg = pd.DataFrame(columns=['Tienda','SKU','ADU'])

adu = agg[['Tienda','SKU','ADU']] if 'ADU' in agg.columns else pd.DataFrame(columns=['Tienda','SKU','ADU'])
stock_df = stock_df.merge(adu, on=['Tienda','SKU'], how='left')
stock_df['ADU'] = stock_df['ADU'].fillna(0.0)
stock_df['Cobertura_dias'] = np.where(stock_df['ADU']>0, stock_df['Existencia']/stock_df['ADU'], np.inf)

# ---------------------
# Mínimos + bodega principal
# ---------------------
bodega_principal = pick_bodega_principal(stock_df)
stock_df['IsEcom']      = stock_df['Tienda'].apply(looks_like_ecom)
stock_df['MinObjetivo'] = np.where(stock_df['IsEcom'], MIN_ECOM, MIN_POR_SKU)
if bodega_principal:
    stock_df.loc[stock_df['Tienda']==bodega_principal, 'MinObjetivo'] = 0  # bodega puede quedar en 0

# Región/Ciudad
if tiendas_map:
    stock_df['Region'] = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Region'))
    stock_df['Ciudad'] = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Ciudad'))
else:
    stock_df['Region'] = None
    stock_df['Ciudad'] = None

# ---------------------
# Snapshot actual
# ---------------------
key_cols = ['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias']
current_stock = stock_df[key_cols + ['Existencia']].copy()

# Cobertura ANTES (para reporte)
cov_before = current_stock[['Tienda','SKU','Referencia','Talla','ADU','Existencia']].copy()
cov_before['Cobertura_antes'] = np.where(cov_before['ADU']>0, cov_before['Existencia']/cov_before['ADU'], np.inf)

# ---------------------
# Necesidades base (mínimos) con CAP de máximo
# ---------------------

if bodega_principal:
    mask_not_bodega = (current_stock['Tienda'] != bodega_principal)
else:
    # si no hay bodega detectada, no filtramos ninguna tienda
    mask_not_bodega = pd.Series(True, index=current_stock.index)

needs_df = current_stock[
    (current_stock['Existencia'] < current_stock['MinObjetivo']) &
    mask_not_bodega
].copy()
needs_df['Necesita'] = (needs_df['MinObjetivo'] - needs_df['Existencia']).astype(int)
needs_df['CapMax']   = MAX_STOCK_PER_SKU - needs_df['Existencia']
needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
needs_df = needs_df[needs_df['Necesita'] > 0].copy()

# ---------------------
# Necesidades extra por curva (N tallas cumpliendo mínimo)
# ---------------------
def build_curve_needs(current_stock):
    # Agrega por tienda-ref-talla para evitar duplicados
    base = current_stock.groupby(
        ['Tienda','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias'],
        as_index=False
    )['Existencia'].sum()

    groups = base[base['RANGO_CAT'].isin(CURVAS_TALLAS.keys())].groupby(
        ['Tienda','Referencia','RANGO_CAT'], as_index=False
    ).agg(
        Region=('Region','first'),
        Ciudad=('Ciudad','first'),
        IsEcom=('IsEcom','first'),
        MinObjetivo=('MinObjetivo','first'),
        ADU=('ADU','first'),
        Cobertura_dias=('Cobertura_dias','first')
    )
    if groups.empty:
        return pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])

    rows = []
    for _, g in groups.iterrows():
        tienda, ref, rango = g['Tienda'], g['Referencia'], g['RANGO_CAT']
        tallas = CURVAS_TALLAS.get(rango, [])
        exists = base[(base['Tienda']==tienda) & (base['Referencia']==ref) & (base['RANGO_CAT']==rango)][['Talla','Existencia']]
        exists = exists.set_index('Talla')['Existencia'].to_dict()

        cumple_count = sum(1 for t in tallas if exists.get(t, 0) >= MIN_POR_SKU)
        faltan = max(0, CURVAS_MIN_TALLAS.get(rango, 0) - cumple_count)
        if faltan <= 0:
            continue

        candidatos = [(t, exists.get(t, 0)) for t in tallas if exists.get(t, 0) < MIN_POR_SKU]
        candidatos.sort(key=lambda x: x[1])  # prioriza menor existencia

        for t, ex in candidatos[:faltan]:
            need_qty = min(MIN_POR_SKU - ex, max(0, MAX_STOCK_PER_SKU - ex))
            if need_qty > 0:
                rows.append({
                    'Tienda': tienda, 'SKU': f"{ref}-{t}", 'Referencia': ref, 'Talla': t, 'RANGO_CAT': rango,
                    'Region': g['Region'], 'Ciudad': g['Ciudad'], 'IsEcom': g['IsEcom'],
                    'MinObjetivo': g['MinObjetivo'], 'ADU': g['ADU'], 'Cobertura_dias': g['Cobertura_dias'],
                    'Necesita': int(need_qty)
                })
    if not rows:
        return pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])
    return pd.DataFrame(rows)

extra_needs = build_curve_needs(current_stock)
if not extra_needs.empty:
    extra = extra_needs[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']].copy()
    needs_df = pd.concat([needs_df[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']], extra], ignore_index=True)
    needs_df = needs_df.groupby(['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias'], as_index=False)['Necesita'].max()

    # Re-cap por máximo
    exist_now = current_stock.groupby(['Tienda','SKU'], as_index=False)['Existencia'].sum().rename(columns={'Existencia':'Exist_now'})
    needs_df = needs_df.merge(exist_now, on=['Tienda','SKU'], how='left')
    needs_df['CapMax'] = MAX_STOCK_PER_SKU - needs_df['Exist_now'].fillna(0)
    needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
    needs_df = needs_df[needs_df['Necesita'] > 0].drop(columns=['Exist_now','CapMax'])

# Priorización (ADU desc, gap desc)
needs_df = needs_df.sort_values(['ADU','Necesita'], ascending=[False, False]).reset_index(drop=True)

# ---------------------
# Motor de traslados (Cobertura + Máximo + Prioridad direccional)
# ---------------------
def get_stock(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    return int(r['Existencia'].sum()) if len(r) else 0

def get_cov_days(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    if not len(r): return 0.0
    adu = float(r['ADU'].iloc[0])
    ex  = int(r['Existencia'].sum())
    return (ex/adu) if adu>0 else np.inf

def allowed_to_send_from_origin(tienda, sku):
    # Bodega principal: puede enviar todo (puede quedar en 0)
    if bodega_principal and tienda == bodega_principal:
        return get_stock(tienda, sku)
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    if not len(r): return 0
    ex   = int(r['Existencia'].sum())
    mino = int(r['MinObjetivo'].iloc[0])
    adu  = float(r['ADU'].iloc[0])
    is_e = bool(r['IsEcom'].iloc[0])
    min_cov = ORIGIN_MIN_COV_ECOM if is_e else ORIGIN_MIN_COV_DAYS
    guard  = int(np.ceil(min_cov * adu)) if adu>0 else mino
    guard  = max(guard, mino)
    return max(0, ex - guard)

# Pre-candidatos por SKU
sku_to_origins = defaultdict(list)
for _, r in current_stock.iterrows():
    tienda, sku = r['Tienda'], r['SKU']
    if allowed_to_send_from_origin(tienda, sku) > 0 and not is_excluded_store(tienda):
        sku_to_origins[sku].append(tienda)

transfers = []
for _, need in needs_df.iterrows():
    dest, sku, ref, talla = need['Tienda'], need['SKU'], need['Referencia'], need['Talla']
    if bodega_principal and dest == bodega_principal:
        continue

    dest_row = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)]
    dest_exist = get_stock(dest, sku)
    dest_adu   = float(dest_row['ADU'].iloc[0]) if len(dest_row) else 0.0
    dest_ecom  = bool(dest_row['IsEcom'].iloc[0]) if len(dest_row) else False
    dest_min   = int(dest_row['MinObjetivo'].iloc[0]) if len(dest_row) else MIN_POR_SKU

    target_cov_days  = DEST_TARGET_COV_ECOM if dest_ecom else DEST_TARGET_COV_DAYS
    target_units_cov = int(np.ceil(target_cov_days * dest_adu)) if dest_adu>0 else dest_min
    target_units     = min(MAX_STOCK_PER_SKU, max(dest_min, target_units_cov))

    remaining = max(int(need['Necesita']), max(0, target_units - dest_exist))
    if remaining <= 0:
        continue

    stock_before_dest = dest_exist

    # orígenes: bodega (si existe) + tiendas con capacidad
    origins = list(dict.fromkeys(([bodega_principal] if bodega_principal else []) + sku_to_origins.get(sku, [])))
    origins = [o for o in origins if o and o != dest]

    # rankear: misma región → prioridad → mayor cobertura → menor ETA
    ranked = []
    for o in origins:
        can = allowed_to_send_from_origin(o, sku)
        if can <= 0:
            continue

        o_cov = get_cov_days(o, sku)
        d_cov = get_cov_days(dest, sku)
        if np.isfinite(o_cov) and np.isfinite(d_cov) and (o_cov <= d_cov + COV_BUFFER_DAYS):
            continue

        # Región/Ciudad
        try:
            o_row    = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].iloc[0]
            d_row    = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].iloc[0]
            o_region = o_row.get('Region'); d_region = d_row.get('Region')
            o_city   = o_row.get('Ciudad'); d_city    = d_row.get('Ciudad')
        except Exception:
            o_region, d_region, o_city, d_city = None, None, None, None

        same_region = (pd.notna(o_region) and pd.notna(d_region) and o_region == d_region)
        pri = delivery_priority(o, dest, tiempos_df) if same_region else np.nan  # 1..6 (menor = mejor)
        lt  = delivery_days(o_city, d_city, tiempos_df)

        ranked.append((
            0 if same_region else 1,                            # 1) misma región primero
            pri if same_region and pd.notna(pri) else 999,      # 2) prioridad asc (solo misma región)
            -(o_cov if np.isfinite(o_cov) else 1e9),            # 3) mayor cobertura origen
            (lt if pd.notna(lt) else 999),                      # 4) menor ETA
            o, can
        ))

    ranked.sort(key=lambda t: (t[0], t[1], t[2], t[3]))

    for reg_pen, pri, _, _, o, can in ranked:
        if remaining <= 0: break
        can_now = allowed_to_send_from_origin(o, sku)
        if can_now <= 0: continue

        dest_exist_now = get_stock(dest, sku)
        gap_to_target  = max(0, target_units - dest_exist_now)
        gap_to_cap     = max(0, MAX_STOCK_PER_SKU - dest_exist_now)

        qty = int(min(can_now, remaining, (gap_to_target if gap_to_target>0 else remaining), gap_to_cap))
        if qty <= 0: continue

        transfers.append({
            'Tienda Origen': o,
            'Tienda Destino': dest,
            'Stock antes traslado tienda destino': stock_before_dest,
            'Unidades a Trasladar': qty,
            'Referencia': ref,
            'Talla': talla
        })

        # Actualiza existencias (si hay filas duplicadas por tienda-SKU, reparte y luego se redondea)
        idx_o = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].index
        current_stock.loc[idx_o, 'Existencia'] = current_stock.loc[idx_o, 'Existencia'] - qty / max(1,len(idx_o))

        idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
        if len(idx_d)==0:
            ref_d, talla_d = split_ref_talla_from_sku(sku)
            new_row = {
                'Tienda': dest, 'SKU': sku, 'Referencia': ref_d, 'Talla': talla_d,
                'RANGO_CAT': None, 'Region': None, 'Ciudad': None, 'IsEcom': looks_like_ecom(dest),
                'MinObjetivo': MIN_ECOM if looks_like_ecom(dest) else MIN_POR_SKU,
                'ADU': dest_adu, 'Cobertura_dias': np.inf, 'Existencia': 0
            }
            current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
            idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
        current_stock.loc[idx_d, 'Existencia'] = current_stock.loc[idx_d, 'Existencia'] + qty / max(1,len(idx_d))

        remaining -= qty
        stock_before_dest = None

# Redondear existencias y evitar negativos por la división entre duplicados
current_stock['Existencia'] = current_stock['Existencia'].round().astype(int)
current_stock.loc[current_stock['Existencia']<0, 'Existencia'] = 0

# ---------------------
# Reportes y salida
# ---------------------
# Tabla de traslados
traslados_df = pd.DataFrame(
    transfers,
    columns=['Tienda Origen','Tienda Destino','Stock antes traslado tienda destino','Unidades a Trasladar','Referencia','Talla']
)

# Cobertura después
cov_after = current_stock[['Tienda','SKU','Referencia','Talla','ADU','Existencia']].copy()
cov_after['Cobertura_despues'] = np.where(cov_after['ADU']>0, cov_after['Existencia']/cov_after['ADU'], np.inf)

# Merge before/after y deltas
cov_merge = cov_before.merge(
    cov_after,
    on=['Tienda','SKU','Referencia','Talla','ADU'],
    how='outer',
    suffixes=('_antes','_despues')
)
for col in ['Existencia_antes','Existencia_despues','Cobertura_antes','Cobertura_despues']:
    if col not in cov_merge.columns:
        cov_merge[col] = np.nan

cov_merge['Delta_exist'] = cov_merge['Existencia_despues'].fillna(0) - cov_merge['Existencia_antes'].fillna(0)
cov_merge['Delta_cov']   = cov_merge['Cobertura_despues'].replace(np.inf, np.nan) - cov_merge['Cobertura_antes'].replace(np.inf, np.nan)

# Solo filas con cambios
cov_changes = cov_merge[(cov_merge['Delta_exist'] != 0) | (cov_merge['Delta_cov'].abs() > 1e-9)].copy()

# Stock después (agregado para evitar duplicados tienda-SKU)
stock_despues = (
    current_stock
      .groupby(['Tienda','SKU','Referencia','Talla'], as_index=False)['Existencia']
      .sum()
      .sort_values(['Tienda','Referencia','Talla'])
)

# Resumen por tienda: envíos/recepciones y cobertura promedio antes/después + % SKUs en target
mov_env = (traslados_df.groupby('Tienda Origen', as_index=False)['Unidades a Trasladar']
           .sum().rename(columns={'Tienda Origen':'Tienda','Unidades a Trasladar':'env_unidades'})) if len(traslados_df) else pd.DataFrame(columns=['Tienda','env_unidades'])
mov_rec = (traslados_df.groupby('Tienda Destino', as_index=False)['Unidades a Trasladar']
           .sum().rename(columns={'Tienda Destino':'Tienda','Unidades a Trasladar':'rec_unidades'})) if len(traslados_df) else pd.DataFrame(columns=['Tienda','rec_unidades'])
movs = pd.merge(mov_env, mov_rec, on='Tienda', how='outer').fillna(0)

cov_tienda_before = cov_before.groupby('Tienda', as_index=False)['Cobertura_antes'].mean()
cov_tienda_after  = cov_after.groupby('Tienda', as_index=False)['Cobertura_despues'].mean()
resumen_tienda = movs.merge(cov_tienda_before, on='Tienda', how='left').merge(cov_tienda_after, on='Tienda', how='left')

def _target_days_for_store(tienda):
    return DEST_TARGET_COV_ECOM if looks_like_ecom(tienda) else DEST_TARGET_COV_DAYS

cov_after_copy = cov_after.copy()
cov_after_copy['target_dias'] = cov_after_copy['Tienda'].apply(_target_days_for_store)
cov_after_copy['cumple_target'] = cov_after_copy.apply(
    lambda r: (r['Cobertura_despues'] >= r['target_dias']) if np.isfinite(r['Cobertura_despues']) else False,
    axis=1
)
pct_target = (cov_after_copy.groupby('Tienda', as_index=False)['cumple_target']
              .mean().rename(columns={'cumple_target':'%_SKUs_en_target'}))
resumen_tienda = resumen_tienda.merge(pct_target, on='Tienda', how='left')

# ---------------------
# Validaciones
# ---------------------
negativos = current_stock[current_stock['Existencia'] < 0].copy()
v1_ok = (len(negativos) == 0)

sobre_tope = stock_despues[stock_despues['Existencia'] > MAX_STOCK_PER_SKU].copy()
v2_ok = (len(sobre_tope) == 0)

to_bodega = pd.DataFrame()
v3_ok = True
if bodega_principal:
    to_bodega = traslados_df[traslados_df['Tienda Destino'] == bodega_principal].copy()
    v3_ok = (len(to_bodega) == 0)

if len(traslados_df):
    origins_pairs = traslados_df[['Tienda Origen','Referencia','Talla']].copy()
    origins_pairs['SKU'] = origins_pairs['Referencia'].astype(str) + "-" + origins_pairs['Talla'].astype(str)
    origins_pairs = origins_pairs[['Tienda Origen','SKU']].drop_duplicates()
else:
    origins_pairs = pd.DataFrame(columns=['Tienda Origen','SKU'])

cov_after_chk = cov_after.merge(
    current_stock[['Tienda','SKU','IsEcom']].drop_duplicates(), on=['Tienda','SKU'], how='left'
)
cov_after_chk['min_cov_req'] = np.where(cov_after_chk['IsEcom'], ORIGIN_MIN_COV_ECOM, ORIGIN_MIN_COV_DAYS)

if len(origins_pairs):
    cov_origins = cov_after_chk.merge(origins_pairs, left_on=['Tienda','SKU'], right_on=['Tienda Origen','SKU'], how='inner')
else:
    cov_origins = cov_after_chk.iloc[0:0].copy()

origen_bajo_cov = cov_origins[
    (np.isfinite(cov_origins['Cobertura_despues'])) & (cov_origins['Cobertura_despues'] < cov_origins['min_cov_req'])
][['Tienda','SKU','ADU','Cobertura_despues','min_cov_req']].copy()
v5_ok = (len(origen_bajo_cov) == 0)

pct = resumen_tienda.copy()
if '%_SKUs_en_target' in pct.columns:
    fuera_rango = pct[(pct['%_SKUs_en_target'] < 0) | (pct['%_SKUs_en_target'] > 1)].copy()
    v6_ok = (len(fuera_rango) == 0)
else:
    fuera_rango = pd.DataFrame()
    v6_ok = True

u_bad = traslados_df[(traslados_df['Unidades a Trasladar'] <= 0) | (traslados_df['Unidades a Trasladar'] % 1 != 0)].copy()
v7_ok = (len(u_bad) == 0)

validaciones_summary = pd.DataFrame([
    {'Validación':'Sin existencias negativas','Resultado':'OK' if v1_ok else 'FALLA','Detalles': len(negativos)},
    {'Validación':f'Tope máximo {MAX_STOCK_PER_SKU} por SKU','Resultado':'OK' if v2_ok else 'FALLA','Detalles': len(sobre_tope)},
    {'Validación':'Sin traslados hacia Bodega Principal','Resultado':'OK' if v3_ok else 'FALLA','Detalles': len(to_bodega)},
    {'Validación':'Cobertura mínima en orígenes (7/10 días)','Resultado':'OK' if v5_ok else 'FALLA','Detalles': len(origen_bajo_cov)},
    {'Validación':'% SKUs en target ∈ [0,1]','Resultado':'OK' if v6_ok else 'FALLA','Detalles': len(fuera_rango)},
    {'Validación':'Unidades a Trasladar > 0 y enteras','Resultado':'OK' if v7_ok else 'FALLA','Detalles': len(u_bad)},
])

# Armar hoja "Validaciones" en bloques
blocks = []
blocks.append(pd.DataFrame([['VALIDACIONES - RESUMEN','','']], columns=['Sección','Campo','Valor']))
blocks.append(validaciones_summary.rename(columns={'Validación':'Campo','Resultado':'Valor','Detalles':'Extras'}))

def add_block(title, df):
    blocks.append(pd.DataFrame([[title,'','']], columns=['Sección','Campo','Valor']))
    if len(df):
        if df.shape[1] > 25:
            df = df.iloc[:, :25]
        blocks.append(df.reset_index(drop=True).astype(object))
    else:
        blocks.append(pd.DataFrame([['(sin registros)','','']], columns=['Sección','Campo','Valor']))

add_block('DETALLES: Existencias negativas', negativos)
add_block(f'DETALLES: Sobre tope ({MAX_STOCK_PER_SKU})', sobre_tope)
add_block('DETALLES: Traslados hacia Bodega Principal', to_bodega)
add_block('DETALLES: Orígenes bajo cobertura mínima', origen_bajo_cov)
add_block('DETALLES: % SKUs en target fuera de [0,1]', fuera_rango)
add_block('DETALLES: Unidades a Trasladar inválidas', u_bad)

validaciones_sheet = pd.concat(blocks, ignore_index=True)

# ---------------------
# Escribir Excel
# ---------------------
with pd.ExcelWriter(OUT_PATH, engine='xlsxwriter') as writer:
    traslados_df.to_excel(writer, sheet_name="Traslados", index=False)
    stock_despues.to_excel(writer, sheet_name="Stock_despues", index=False)
    cov_changes.sort_values(['Tienda','Referencia','Talla']).to_excel(writer, sheet_name="Cobertura_before_after", index=False)
    resumen_tienda.sort_values('Tienda').to_excel(writer, sheet_name="Resumen_tienda", index=False)
    validaciones_sheet.to_excel(writer, sheet_name="Validaciones", index=False)

print(f"OK -> {OUT_PATH} | Traslados: {len(traslados_df)} | Filas Stock_despues: {len(stock_despues)}")