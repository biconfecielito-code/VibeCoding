import warnings
warnings.simplefilter("ignore", UserWarning)

import pandas as pd
import numpy as np
import re
from collections import defaultdict

# ---------------------
# Rutas (ajusta si es necesario)
# ---------------------
PATH_STOCK = "PruebaStockIA.xlsx"              # contiene hojas de ventas y stock (nombres detectados automáticamente)
PATH_TIEN  = "Clasificacion_Tiendas.xlsx"      # opcional: para Región/Tipo/RegionID
PATH_TENT  = "Tiempos de entrega.xlsx"         # direccional TIENDA: ORIGEN-DESTINO, DESTINO-ORIGEN, ETA, PRIORIDAD
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

ORIGIN_MIN_COV_DAYS, DEST_TARGET_COV_DAYS = 7, 7
ORIGIN_MIN_COV_ECOM, DEST_TARGET_COV_ECOM = 10, 10
COV_BUFFER_DAYS = 1  # evitar traslado si cobertura origen ≤ cobertura destino + 1 día

# ---------------------
# Utilidades
# ---------------------
def std_col(s):
    return str(s).strip().lower().replace("\n"," ")

def normalize_store_name(x):
    return str(x).strip().upper() if pd.notna(x) else x

def looks_like_ecom(name):
    s = normalize_store_name(name)
    return any(k in s for k in ["ECO","ECOM","ONLINE","VIRTUAL","WEB"])

# SKU sin separador → REF = primeros 7, TALLA = resto
def split_ref_talla_from_sku(sku):
    if pd.isna(sku):
        return (None, None)
    s = str(sku).strip()
    if len(s) < 8:
        return (s, None)
    ref = s[:7].strip()
    talla = s[7:].strip().upper()
    if not talla:
        talla = None
    return ref, talla

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

def canonical_rango(x):
    s = str(x).strip().upper()
    if ('5301' in s) or ('BEBE' in s) or s == "5301": return 'BEBES'
    if ('5302' in s) or ('NIÑ' in s) or ('NIN' in s) or s == "5302": return 'NIÑOS'
    return None

def pick_bodega_principal(stock_df):
    tmp = stock_df.groupby('Tienda', as_index=False)['Existencia'].sum()
    if not len(tmp): return None
    tmp['score'] = tmp['Tienda'].str.contains('BODEGA|CEDI|PRINCIPAL', case=False, regex=True).astype(int)
    tmp = tmp.sort_values(['score','Existencia'], ascending=[False, False])
    return tmp.iloc[0]['Tienda']

def parse_lead_time_value(x):
    """Soporta: 1, 2, 3, '3 a 4', '1 A 2' → toma el máximo."""
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

# ---------------------
# Clasificación / Región / Tipo / RegionID (opcional) — sin Ciudad
# ---------------------
try:
    ct = pd.read_excel(PATH_TIEN)
    ct_cols = {std_col(c): c for c in ct.columns}

    col_tienda    = ct_cols.get('tienda') or list(ct.columns)[0]
    col_tipo      = ct_cols.get('tipo')
    col_region    = ct_cols.get('region') or ct_cols.get('región')
    col_region_id = (ct_cols.get('region id') or ct_cols.get('region_id') or
                     ct_cols.get('regionid') or ct_cols.get('region code'))

    keep = [col_tienda] \
           + ([col_tipo] if col_tipo else []) \
           + ([col_region] if col_region else []) \
           + ([col_region_id] if col_region_id else [])
    ct_df = ct[keep].copy()

    new_cols = ['Tienda']
    if col_tipo:      new_cols.append('Tipo')
    if col_region:    new_cols.append('Region')
    if col_region_id: new_cols.append('RegionID')
    ct_df.columns = new_cols

    ct_df['Tienda'] = ct_df['Tienda'].apply(normalize_store_name)
    if 'Tipo' in ct_df.columns:
        ct_df['Tipo'] = ct_df['Tipo'].astype(str).str.strip().str.upper()
    if 'RegionID' in ct_df.columns:
        ct_df['RegionID'] = pd.to_numeric(ct_df['RegionID'], errors='coerce').astype('Int64')

    tiendas_map = ct_df.set_index('Tienda').to_dict(orient='index')
except Exception:
    tiendas_map = {}
    ct_df = None  # por si queremos fallback más abajo

# ---------------------
# Tiempos de entrega por TIENDA (direccional) con PRIORIDAD
# ---------------------
try:
    tiempos_df = pd.read_excel(PATH_TENT)
    cols = {std_col(c): c for c in tiempos_df.columns}
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
    """PRIORIDAD 1..n. Si no hay dato, NaN."""
    if tiempos_df is None or tiempos_df.empty:
        return np.nan
    oc = str(origin_store).strip().upper()
    dc = str(dest_store).strip().upper()
    m = tiempos_df[(tiempos_df['_O'] == oc) & (tiempos_df['_D'] == dc)]
    if len(m):
        val = m.iloc[0]['_PRI_NUM']
        return float(val) if pd.notna(val) else np.nan
    return np.nan

def delivery_days(origin_store, dest_store, tiempos_df):
    """ETA en días. Si no hay dato, NaN."""
    if tiempos_df is None or tiempos_df.empty:
        return np.nan
    oc = str(origin_store).strip().upper()
    dc = str(dest_store).strip().upper()
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

# Si no vienen referencia/talla, derivar del SKU (7 + resto)
if 'Referencia' not in stock_df.columns or 'Talla' not in stock_df.columns:
    ref_parsed, talla_parsed = zip(*stock_df['SKU'].apply(split_ref_talla_from_sku).tolist())
    if 'Referencia' not in stock_df.columns: stock_df['Referencia'] = ref_parsed
    if 'Talla' not in stock_df.columns:      stock_df['Talla']      = talla_parsed

stock_df['Tienda']     = stock_df['Tienda'].apply(normalize_store_name)
stock_df['Existencia'] = pd.to_numeric(stock_df['Existencia'], errors='coerce').fillna(0).astype(int)
stock_df['RANGO_CAT']  = stock_df.get('RANGO_RAW', np.nan).apply(canonical_rango)

# Añadir Tipo / Región / RegiónID desde clasificación (sin Ciudad)
if tiendas_map:
    stock_df['Tipo']     = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Tipo'))
    stock_df['Region']   = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Region'))
    stock_df['RegionID'] = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('RegionID'))
else:
    stock_df['Tipo'] = None
    stock_df['Region'] = None
    stock_df['RegionID'] = None

# Filtrar por tallas válidas (solo BEBES/NIÑOS con sus curvas)
valid_bebes = set(CURVAS_TALLAS['BEBES'])
valid_ninos = set(CURVAS_TALLAS['NIÑOS'])
mask_valid = ((stock_df['RANGO_CAT']=='BEBES') & (stock_df['Talla'].isin(valid_bebes))) | \
             ((stock_df['RANGO_CAT']=='NIÑOS') & (stock_df['Talla'].isin(valid_ninos)))
stock_df = stock_df[mask_valid].copy()

# ---------------------
# ADU (desde ventas) con DÍAS GLOBALES + Cobertura
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

    # Derivar Ref/Talla y filtrar por tallas válidas
    v[['Ref_tmp','Talla_tmp']] = v['SKU'].apply(lambda s: pd.Series(split_ref_talla_from_sku(s)))
    valid_union = valid_bebes.union(valid_ninos)
    v = v[v['Talla_tmp'].isin(valid_union)].copy()

    # DÍAS BASE GLOBAL (idénticos para todos los SKUs)
    if 'Fecha' in v.columns:
        v['Fecha'] = pd.to_datetime(v['Fecha'], errors='coerce')
        v = v[v['Fecha'].notna()].copy()
        v['Dia'] = v['Fecha'].dt.date
        dias_base_global = int(pd.Series(v['Dia'].unique()).nunique())
        if dias_base_global <= 0:
            dias_base_global = 1
    else:
        dias_base_global = 30

    agg = v.groupby(['Tienda','SKU']).agg(total_units=('Unidades','sum')).reset_index()
    agg['ADU'] = agg['total_units'] / dias_base_global
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

# ---------------------
# Snapshot actual
# ---------------------
key_cols = ['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias']
current_stock = stock_df[key_cols + ['Existencia']].copy()

# ---------------------
# Necesidades base (mínimos) con CAP de máximo
# ---------------------
if bodega_principal:
    mask_not_bodega = (current_stock['Tienda'] != bodega_principal)
else:
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
    base = current_stock.groupby(
        ['Tienda','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias'],
        as_index=False
    )['Existencia'].sum()

    groups = base[base['RANGO_CAT'].isin(CURVAS_TALLAS.keys())].groupby(
        ['Tienda','Referencia','RANGO_CAT'], as_index=False
    ).agg(
        Region=('Region','first'),
        RegionID=('RegionID','first'),
        Tipo=('Tipo','first'),
        IsEcom=('IsEcom','first'),
        MinObjetivo=('MinObjetivo','first'),
        ADU=('ADU','first'),
        Cobertura_dias=('Cobertura_dias','first')
    )
    if groups.empty:
        return pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])

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
                    'Tienda': tienda, 'SKU': f"{ref}{t}", 'Referencia': ref, 'Talla': t, 'RANGO_CAT': rango,
                    'Region': g['Region'], 'RegionID': g['RegionID'], 'Tipo': g['Tipo'],
                    'IsEcom': g['IsEcom'], 'MinObjetivo': g['MinObjetivo'], 'ADU': g['ADU'], 'Cobertura_dias': g['Cobertura_dias'],
                    'Necesita': int(need_qty)
                })
    if not rows:
        return pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])
    return pd.DataFrame(rows)

extra_needs = build_curve_needs(current_stock)
if not extra_needs.empty:
    extra = extra_needs[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']].copy()
    needs_df = pd.concat([needs_df[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']], extra], ignore_index=True)
    needs_df = needs_df.groupby(['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias'], as_index=False)['Necesita'].max()

    exist_now = current_stock.groupby(['Tienda','SKU'], as_index=False)['Existencia'].sum().rename(columns={'Existencia':'Exist_now'})
    needs_df = needs_df.merge(exist_now, on=['Tienda','SKU'], how='left')
    needs_df['CapMax'] = MAX_STOCK_PER_SKU - needs_df['Exist_now'].fillna(0)
    needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
    needs_df = needs_df[needs_df['Necesita'] > 0].drop(columns=['Exist_now','CapMax'])

# Priorización de necesidades (ADU desc, gap desc)
needs_df = needs_df.sort_values(['ADU','Necesita'], ascending=[False, False]).reset_index(drop=True)

# ---------------------
# Mapa de región a nivel tienda (para evitar depender del SKU)
# ---------------------
def _norm(s): return str(s).strip().upper() if pd.notna(s) else None

if 'ct_df' in globals() and ct_df is not None:
    store_meta_df = ct_df[['Tienda','RegionID','Region']].drop_duplicates('Tienda')
else:
    store_meta_df = current_stock[['Tienda','RegionID','Region']].drop_duplicates('Tienda')

_store_meta = {}
for _, r in store_meta_df.iterrows():
    t = _norm(r['Tienda'])
    rid = str(int(r['RegionID'])) if pd.notna(r.get('RegionID')) else None
    reg = _norm(r.get('Region'))
    _store_meta[t] = (rid, reg)

def same_region_storelevel(origin_store, dest_store):
    o = _store_meta.get(_norm(origin_store), (None, None))
    d = _store_meta.get(_norm(dest_store), (None, None))
    o_rid, o_reg = o
    d_rid, d_reg = d
    if o_rid and d_rid:
        return o_rid == d_rid
    if o_reg and d_reg:
        return o_reg == d_reg
    return False

# ---------------------
# Motor de traslados (usa prioridad SIEMPRE) + capturas de stock antes/después
# ---------------------
def get_stock(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    return int(r['Existencia'].sum()) if len(r) else 0

def get_cov_days(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    if not len(r): return np.inf
    adu = float(r['ADU'].iloc[0])
    ex  = int(r['Existencia'].sum())
    return (ex/adu) if adu>0 else np.inf

def allowed_to_send_from_origin(tienda, sku):
    # Bodega principal: puede enviar todo lo que tenga (sin mínimo)
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
    if allowed_to_send_from_origin(tienda, sku) > 0:
        sku_to_origins[sku].append(tienda)

transfers = []
for _, need in needs_df.iterrows():
    dest, sku, ref, talla = need['Tienda'], need['SKU'], need['Referencia'], need['Talla']
    # No enviar a bodega como destino
    if bodega_principal and dest == bodega_principal:
        continue

    # Datos del destino
    dest_row = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)]
    dest_adu   = float(dest_row['ADU'].iloc[0]) if len(dest_row) else 0.0
    dest_ecom  = bool(dest_row['IsEcom'].iloc[0]) if len(dest_row) else False
    dest_min   = int(dest_row['MinObjetivo'].iloc[0]) if len(dest_row) else (MIN_ECOM if looks_like_ecom(dest) else MIN_POR_SKU)

    target_cov_days  = DEST_TARGET_COV_ECOM if dest_ecom else DEST_TARGET_COV_DAYS
    target_units_cov = int(np.ceil(target_cov_days * dest_adu)) if dest_adu>0 else dest_min
    target_units     = min(MAX_STOCK_PER_SKU, max(dest_min, target_units_cov))

    # Candidatos: bodega (si existe) + tiendas con capacidad
    origins = list(dict.fromkeys(([bodega_principal] if bodega_principal else []) + sku_to_origins.get(sku, [])))
    origins = [o for o in origins if o and o != dest]

    ranked = []
    for o in origins:
        can = allowed_to_send_from_origin(o, sku)
        if can <= 0:
            continue

        o_cov = get_cov_days(o, sku)
        d_cov = get_cov_days(dest, sku)
        # filtro por cobertura relativa
        if np.isfinite(o_cov) and np.isfinite(d_cov) and (o_cov <= d_cov + COV_BUFFER_DAYS):
            continue

        same_region = same_region_storelevel(o, dest)
        pri = delivery_priority(o, dest, tiempos_df)       # <-- AHORA SIEMPRE
        pri_sort = int(pri) if pd.notna(pri) else 999
        lt  = delivery_days(o, dest, tiempos_df)
        lt_sort = float(lt) if pd.notna(lt) else 999

        ranked.append((
            0 if same_region else 1,                        # 1) MISMA región primero
            pri_sort,                                       # 2) prioridad (para todos)
            -(o_cov if np.isfinite(o_cov) else 1e9),        # 3) mayor cobertura de origen
            lt_sort,                                        # 4) menor ETA
            o
        ))

    ranked.sort(key=lambda t: (t[0], t[1], t[2], t[3]))

    # Asignaciones por origen según ranking hasta cumplir target/cap
    while True:
        dest_exist_now = get_stock(dest, sku)
        gap_to_target  = max(0, target_units - dest_exist_now)
        if gap_to_target <= 0:
            break

        moved = False
        for _, _, _, _, o in ranked:
            can_now = allowed_to_send_from_origin(o, sku)
            if can_now <= 0:
                continue

            origin_before = get_stock(o, sku)
            dest_before   = dest_exist_now

            gap_to_cap = max(0, MAX_STOCK_PER_SKU - dest_before)
            qty = int(min(can_now, gap_to_target, gap_to_cap))
            if qty <= 0:
                continue

            origin_after = origin_before - qty
            dest_after   = dest_before + qty

            transfers.append({
                'Tienda origen': o,
                'Tienda destino': dest,
                'Stock tienda origen antes traslado': origin_before,
                'Stock tienda origen despues traslado': origin_after,
                'Stock tienda destino antes traslado': dest_before,
                'Stock tienda destino despues del traslado': dest_after,
                'Unidades a trasladar': qty,
                'Referencia': ref,
                'Talla': talla
            })

            # Actualizar existencias en snapshot
            idx_o = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].index
            current_stock.loc[idx_o, 'Existencia'] = current_stock.loc[idx_o, 'Existencia'] - qty / max(1,len(idx_o))

            idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
            if len(idx_d)==0:
                ref_d, talla_d = split_ref_talla_from_sku(sku)
                new_row = {
                    'Tienda': dest, 'SKU': sku, 'Referencia': ref_d, 'Talla': talla_d,
                    'RANGO_CAT': None, 'Region': None, 'RegionID': None, 'Tipo': None,
                    'IsEcom': looks_like_ecom(dest), 'MinObjetivo': MIN_ECOM if looks_like_ecom(dest) else MIN_POR_SKU,
                    'ADU': dest_adu, 'Cobertura_dias': np.inf, 'Existencia': 0
                }
                current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
                idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
            current_stock.loc[idx_d, 'Existencia'] = current_stock.loc[idx_d, 'Existencia'] + qty / max(1,len(idx_d))

            moved = True
            break  # volver a evaluar gap_to_target con stocks actualizados

        if not moved:
            break  # no hay más candidatos viables

# Redondear existencias y evitar negativos
current_stock['Existencia'] = current_stock['Existencia'].round().astype(int)
current_stock.loc[current_stock['Existencia']<0, 'Existencia'] = 0

# ---------------------
# Salida
# ---------------------
traslados_df = pd.DataFrame(transfers, columns=[
    'Tienda origen','Tienda destino',
    'Stock tienda origen antes traslado','Stock tienda origen despues traslado',
    'Stock tienda destino antes traslado','Stock tienda destino despues del traslado',
    'Unidades a trasladar',
    'Referencia','Talla'
])

stock_despues = (
    current_stock
      .groupby(['Tienda','SKU','Referencia','Talla'], as_index=False)['Existencia']
      .sum()
      .sort_values(['Tienda','Referencia','Talla'])
)

with pd.ExcelWriter(OUT_PATH, engine='xlsxwriter') as writer:
    traslados_df.to_excel(writer, sheet_name="Traslados", index=False)
    stock_despues.to_excel(writer, sheet_name="Stock_despues", index=False)
