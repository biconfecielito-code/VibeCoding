# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import re
from collections import defaultdict

# Rutas
PATH_STOCK = "PruebaStockIA.xlsx"
PATH_TIEN  = "Clasificacion_Tiendas.xlsx"
PATH_TENT  = "Tiempos de entrega.xlsx"
OUT_PATH   = "Sugerencias_Traslados_Proyecto.xlsx"

# Parámetros
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
COV_BUFFER_DAYS = 1

TIENDAS_CERRADAS = ["MEDELLIN FABRICATO", "CALI MALL PLAZA", "CARTAGENA MALLPLAZA"]
PALABRAS_EXCLUIR = ["FOTO", "FOTOG", "ARREG", "PRODUC", "IMPERF"]

def std_col(s): return str(s).strip().lower().replace("\n"," ")
def normalize_store_name(x): return str(x).strip().upper() if pd.notna(x) else x
def looks_like_ecom(name):
    s = normalize_store_name(name)
    return any(k in s for k in ["ECO","ECOM","ONLINE","VIRTUAL","WEB"])
def is_excluded_store(name):
    s = normalize_store_name(name)
    return any(s == c for c in TIENDAS_CERRADAS) or any(k in s for k in PALABRAS_EXCLUIR)
def split_ref_talla_from_sku(sku):
    if pd.isna(sku): return (None, None)
    s = str(sku)
    if "-" in s: a,b = s.rsplit("-",1); return a.strip(), b.strip()
    if " " in s: a,b = s.rsplit(" ",1); return a.strip(), b.strip()
    return s, None
def detect_date_column(df):
    for c in df.columns:
        sc = std_col(c)
        if 'fecha' in sc or 'fec' in sc or 'f. doc' in sc or 'documento' in sc:
            try:
                if pd.to_datetime(df[c], errors='coerce').notna().any():
                    return c
            except: pass
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
    if pd.isna(x): return np.nan
    s = str(x).strip().upper()
    nums = [int(n) for n in re.findall(r'\d+', s)]
    return max(nums) if nums else np.nan

# Carga base
xl = pd.ExcelFile(PATH_STOCK)
m = {std_col(n): n for n in xl.sheet_names}
ventas_sheet = next((m[k] for k in m if ('ventasref' in k or 'venta' in k or 'ventas' in k or 'sales' in k)), None)
stock_sheet  = next((m[k] for k in m if ('stockref' in k or 'stock' in k or 'invent' in k)), None)
if ventas_sheet is None or stock_sheet is None:
    raise ValueError("No se encontraron hojas de Ventas/Stock.")
ventas = xl.parse(ventas_sheet)
stock  = xl.parse(stock_sheet)

# Clasificación/Región/Ciudad
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

# Tiempos entrega direccional con PRIORIDAD
try:
    tiempos_df = pd.read_excel(PATH_TENT)
    cols = {std_col(c): c for c in tiempos_df.columns}
    col_o   = cols.get('origen-destino') or cols.get('origen') or list(tiempos_df.columns)[0]
    col_d   = cols.get('destino-origen') or cols.get('destino') or list(tiempos_df.columns)[1]
    col_eta = cols.get('eta') or cols.get('dias') or cols.get('días') or cols.get('tiempo (dias)') or list(tiempos_df.columns)[-2]
    col_pri = next((cols[k] for k in cols if 'priorid' in k or 'priori' in k), None)

    tiempos_df = tiempos_df.copy()
    tiempos_df['_O']       = tiempos_df[col_o].astype(str).str.upper().str.strip()
    tiempos_df['_D']       = tiempos_df[col_d].astype(str).str.upper().str.strip()
    tiempos_df['_ETA_NUM'] = tiempos_df[col_eta].apply(parse_lead_time_value)
    tiempos_df['_PRI_NUM'] = pd.to_numeric(tiempos_df[col_pri], errors='coerce') if col_pri else np.nan
except Exception:
    tiempos_df = pd.DataFrame()

# Normalización stock
stc = {std_col(c): c for c in stock.columns}
c_sku   = stc.get('sku')
c_exist = stc.get('existencia') or stc.get('existencias')
c_tnd   = stc.get('desc. bodega') or stc.get('bodega') or stc.get('tienda')
c_ref   = stc.get('referencia')
c_talla = stc.get('talla')
c_rango = stc.get('rango')
if c_sku is None or c_exist is None or c_tnd is None:
    raise ValueError("Faltan columnas en Stock: SKU / Existencia / Tienda")
stock_df = stock[[c_sku, c_exist, c_tnd] + ([c_ref] if c_ref else []) + ([c_talla] if c_talla else []) + ([c_rango] if c_rango else [])].copy()
stock_df.columns = ['SKU','Existencia','Tienda'] + (['Referencia'] if c_ref else []) + (['Talla'] if c_talla else []) + (['RANGO_RAW'] if c_rango else [])
if 'Referencia' not in stock_df.columns or 'Talla' not in stock_df.columns:
    ref_parsed, talla_parsed = zip(*stock_df['SKU'].apply(split_ref_talla_from_sku).tolist())
    if 'Referencia' not in stock_df.columns: stock_df['Referencia'] = ref_parsed
    if 'Talla' not in stock_df.columns:      stock_df['Talla']      = talla_parsed
stock_df['Tienda']     = stock_df['Tienda'].apply(normalize_store_name)
stock_df['Existencia'] = pd.to_numeric(stock_df['Existencia'], errors='coerce').fillna(0).astype(int)
stock_df               = stock_df[~stock_df['Tienda'].apply(is_excluded_store)].copy()
stock_df['RANGO_CAT']  = stock_df.get('RANGO_RAW', np.nan).apply(canonical_rango)
valid_bebes = set(CURVAS_TALLAS['BEBES']); valid_ninos = set(CURVAS_TALLAS['NIÑOS'])
mask_valid = ((stock_df['RANGO_CAT']=='BEBES') & (stock_df['Talla'].isin(valid_bebes))) | ((stock_df['RANGO_CAT']=='NIÑOS') & (stock_df['Talla'].isin(valid_ninos)))
stock_df = stock_df[mask_valid].copy()

# ADU
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
    v[['Ref_tmp','Talla_tmp']] = v['SKU'].apply(lambda s: pd.Series(split_ref_talla_from_sku(s)))
    valid_union = valid_bebes.union(valid_ninos)
    v = v[v['Talla_tmp'].isin(valid_union)].copy()
    if 'Fecha' in v.columns:
        v['Fecha'] = pd.to_datetime(v['Fecha'], errors='coerce'); v = v[v['Fecha'].notna()].copy()
        v['Dia'] = v['Fecha'].dt.date
        agg = v.groupby(['Tienda','SKU']).agg(total_units=('Unidades','sum'),
                                              fecha_min=('Dia','min'), fecha_max=('Dia','max'),
                                              dias_distintos=('Dia','nunique')).reset_index()
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

# Mínimos y bodega
bodega_principal = pick_bodega_principal(stock_df)
stock_df['IsEcom']      = stock_df['Tienda'].apply(looks_like_ecom)
stock_df['MinObjetivo'] = np.where(stock_df['IsEcom'], MIN_ECOM, MIN_POR_SKU)
if bodega_principal:
    stock_df.loc[stock_df['Tienda']==bodega_principal, 'MinObjetivo'] = 0
if 'Region' not in stock_df.columns or 'Ciudad' not in stock_df.columns:
    stock_df['Region'] = None; stock_df['Ciudad'] = None
try:
    stock_df['Region'] = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Region'))
    stock_df['Ciudad'] = stock_df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Ciudad'))
except: pass

# Snapshot
key_cols = ['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias']
current_stock = stock_df[key_cols + ['Existencia']].copy()

# Necesidades base
if bodega_principal:
    mask_not_bodega = (current_stock['Tienda'] != bodega_principal)
else:
    mask_not_bodega = pd.Series(True, index=current_stock.index)
needs_df = current_stock[(current_stock['Existencia'] < current_stock['MinObjetivo']) & mask_not_bodega].copy()
needs_df['Necesita'] = (needs_df['MinObjetivo'] - needs_df['Existencia']).astype(int)
needs_df['CapMax']   = MAX_STOCK_PER_SKU - needs_df['Existencia']
needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
needs_df = needs_df[needs_df['Necesita'] > 0].copy()

# Curvas (necesidades extra)
def build_curve_needs(current_stock):
    base = current_stock.groupby(['Tienda','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias'], as_index=False)['Existencia'].sum()
    groups = base[base['RANGO_CAT'].isin(CURVAS_TALLAS.keys())].groupby(['Tienda','Referencia','RANGO_CAT'], as_index=False).agg(
        Region=('Region','first'), Ciudad=('Ciudad','first'), IsEcom=('IsEcom','first'),
        MinObjetivo=('MinObjetivo','first'), ADU=('ADU','first'), Cobertura_dias=('Cobertura_dias','first')
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
        if faltan <= 0: continue
        candidatos = [(t, exists.get(t, 0)) for t in tallas if exists.get(t, 0) < MIN_POR_SKU]
        candidatos.sort(key=lambda x: x[1])
        for t, ex in candidatos[:faltan]:
            need_qty = min(MIN_POR_SKU - ex, max(0, MAX_STOCK_PER_SKU - ex))
            if need_qty > 0:
                rows.append({'Tienda': tienda, 'SKU': f"{ref}-{t}", 'Referencia': ref, 'Talla': t, 'RANGO_CAT': rango,
                             'Region': g['Region'], 'Ciudad': g['Ciudad'], 'IsEcom': g['IsEcom'],
                             'MinObjetivo': g['MinObjetivo'], 'ADU': g['ADU'], 'Cobertura_dias': g['Cobertura_dias'],
                             'Necesita': int(need_qty)})
    if not rows:
        return pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])
    return pd.DataFrame(rows)

extra_needs = build_curve_needs(current_stock)
if not extra_needs.empty:
    extra = extra_needs[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']].copy()
    needs_df = pd.concat([needs_df[['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita']], extra], ignore_index=True)
    needs_df = needs_df.groupby(['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','Ciudad','IsEcom','MinObjetivo','ADU','Cobertura_dias'], as_index=False)['Necesita'].max()
    exist_now = current_stock.groupby(['Tienda','SKU'], as_index=False)['Existencia'].sum().rename(columns={'Existencia':'Exist_now'})
    needs_df = needs_df.merge(exist_now, on=['Tienda','SKU'], how='left')
    needs_df['CapMax'] = MAX_STOCK_PER_SKU - needs_df['Exist_now'].fillna(0)
    needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
    needs_df = needs_df[needs_df['Necesita'] > 0].drop(columns=['Exist_now','CapMax'])
needs_df = needs_df.sort_values(['ADU','Necesita'], ascending=[False, False]).reset_index(drop=True)

def get_stock(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    return int(r['Existencia'].sum()) if len(r) else 0
def get_cov_days(tienda, sku):
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    if not len(r): return 0.0
    adu = float(r['ADU'].iloc[0]); ex = int(r['Existencia'].sum())
    return (ex/adu) if adu>0 else np.inf
def allowed_to_send_from_origin(tienda, sku):
    if bodega_principal and tienda == bodega_principal:
        return get_stock(tienda, sku)
    r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
    if not len(r): return 0
    ex   = int(r['Existencia'].sum()); mino = int(r['MinObjetivo'].iloc[0])
    adu  = float(r['ADU'].iloc[0]); is_e = bool(r['IsEcom'].iloc[0])
    min_cov = ORIGIN_MIN_COV_ECOM if is_e else ORIGIN_MIN_COV_DAYS
    guard  = int(np.ceil(min_cov * adu)) if adu>0 else mino
    guard  = max(guard, mino)
    return max(0, ex - guard)

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
    if remaining <= 0: continue

    stock_before_dest = dest_exist
    origins = list(dict.fromkeys(([bodega_principal] if bodega_principal else []) + sku_to_origins.get(sku, [])))
    origins = [o for o in origins if o and o != dest]

    ranked = []
    for o in origins:
        can = allowed_to_send_from_origin(o, sku)
        if can <= 0: continue
        o_cov = get_cov_days(o, sku); d_cov = get_cov_days(dest, sku)
        if np.isfinite(o_cov) and np.isfinite(d_cov) and (o_cov <= d_cov + COV_BUFFER_DAYS): continue

        try:
            o_row    = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].iloc[0]
            d_row    = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].iloc[0]
            o_region = o_row.get('Region'); d_region = d_row.get('Region')
            o_city   = o_row.get('Ciudad'); d_city    = d_row.get('Ciudad')
        except Exception:
            o_region, d_region, o_city, d_city = None, None, None, None

        same_region = (pd.notna(o_region) and pd.notna(d_region) and o_region == d_region)

        if not tiempos_df.empty:
            pri_series = tiempos_df[(tiempos_df['_O'] == str(o).upper()) & (tiempos_df['_D'] == str(dest).upper())]['_PRI_NUM']
            pri = float(pri_series.iloc[0]) if len(pri_series) and pd.notna(pri_series.iloc[0]) else np.nan
            if pd.notna(o_city) and pd.notna(d_city):
                eta_series = tiempos_df[(tiempos_df['_O'] == str(o_city).upper()) & (tiempos_df['_D'] == str(d_city).upper())]['_ETA_NUM']
                lt = float(eta_series.iloc[0]) if len(eta_series) and pd.notna(eta_series.iloc[0]) else np.nan
            else:
                lt = np.nan
        else:
            pri, lt = np.nan, np.nan

        ranked.append((0 if same_region else 1,
                       pri if same_region and pd.notna(pri) else 999,
                       -(o_cov if np.isfinite(o_cov) else 1e9),
                       (lt if pd.notna(lt) else 999),
                       o, can))

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

        idx_o = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].index
        current_stock.loc[idx_o, 'Existencia'] = current_stock.loc[idx_o, 'Existencia'] - qty / max(1,len(idx_o))

        idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
        if len(idx_d)==0:
            ref_d, talla_d = split_ref_talla_from_sku(sku)
            new_row = {'Tienda': dest, 'SKU': sku, 'Referencia': ref_d, 'Talla': talla_d,
                       'RANGO_CAT': None, 'Region': None, 'Ciudad': None, 'IsEcom': looks_like_ecom(dest),
                       'MinObjetivo': MIN_ECOM if looks_like_ecom(dest) else MIN_POR_SKU,
                       'ADU': dest_adu, 'Cobertura_dias': np.inf, 'Existencia': 0}
            current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
            idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
        current_stock.loc[idx_d, 'Existencia'] = current_stock.loc[idx_d, 'Existencia'] + qty / max(1,len(idx_d))

        remaining -= qty
        stock_before_dest = None

current_stock['Existencia'] = current_stock['Existencia'].round().astype(int)
current_stock.loc[current_stock['Existencia']<0, 'Existencia'] = 0

traslados_df = pd.DataFrame(
    transfers,
    columns=['Tienda Origen','Tienda Destino','Stock antes traslado tienda destino','Unidades a Trasladar','Referencia','Talla']
)

stock_despues = (
    current_stock
      .groupby(['Tienda','SKU','Referencia','Talla'], as_index=False)['Existencia']
      .sum()
      .sort_values(['Tienda','Referencia','Talla'])
)

with pd.ExcelWriter(OUT_PATH, engine='openpyxl') as writer:
    traslados_df.to_excel(writer, sheet_name="Traslados", index=False)
    stock_despues.to_excel(writer, sheet_name="Stock_despues", index=False)

print(f"OK -> {OUT_PATH} | Traslados: {len(traslados_df)} | Filas Stock_despues: {len(stock_despues)}")
