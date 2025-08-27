# Basecompleta.py (reforzado)
import warnings
warnings.simplefilter("ignore", UserWarning)

import argparse
from pathlib import Path
import pandas as pd
import numpy as np
import re
from collections import defaultdict
import sys

# ===== Parámetros de negocio =====
MIN_POR_SKU = 2
MIN_ECOM    = 3
MAX_STOCK_PER_SKU = 5

CURVAS_TALLAS = {
    'BEBES': ['0M','3M','6M','9M','12M','18M'],
    'NIÑOS': ['2T','3T','4T','5T','6','8','10','12']
}
CURVAS_MIN_TALLAS = {'BEBES':3, 'NIÑOS':5}

ORIGIN_MIN_COV_DAYS, DEST_TARGET_COV_DAYS = 7, 7
ORIGIN_MIN_COV_ECOM, DEST_TARGET_COV_ECOM = 7, 7
COV_BUFFER_DAYS = 1

# ===== Utilidades =====
def std_col(s): return str(s).strip().lower().replace("\n"," ")
def normalize_store_name(x): return str(x).strip().upper() if pd.notna(x) else x
def looks_like_ecom(name):
    s = normalize_store_name(name)
    return any(k in s for k in ["ECO","ECOM","ONLINE","VIRTUAL","WEB"])
def split_ref_talla_from_sku(sku):
    if pd.isna(sku): return (None, None)
    s = str(sku).strip()
    if len(s) < 8: return (s, None)
    return s[:7].strip(), (s[7:].strip().upper() or None)
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
    if s in ('BEBES','NIÑOS'): return s
    return None
def parse_lead_time_value(x):
    if pd.isna(x): return np.nan
    s = str(x).strip().upper()
    nums = [int(n) for n in re.findall(r'\d+', s)]
    return max(nums) if nums else np.nan
def delivery_priority(origin_store, dest_store, tiempos_df):
    if tiempos_df is None or tiempos_df.empty: return np.nan
    oc = str(origin_store).strip().upper(); dc = str(dest_store).strip().upper()
    m = tiempos_df[(tiempos_df['_O'] == oc) & (tiempos_df['_D'] == dc)]
    if len(m):
        val = m.iloc[0]['_PRI_NUM']; return float(val) if pd.notna(val) else np.nan
    return np.nan
def delivery_days(origin_store, dest_store, tiempos_df):
    if tiempos_df is None or tiempos_df.empty: return np.nan
    oc = str(origin_store).strip().upper(); dc = str(dest_store).strip().upper()
    m = tiempos_df[(tiempos_df['_O'] == oc) & (tiempos_df['_D'] == dc)]
    if len(m):
        val = m.iloc[0]['_ETA_NUM']; return float(val) if pd.notna(val) else np.nan
    return np.nan
def _clean_ref_col(series: pd.Series) -> pd.Series:
    return (series.astype("string").str.strip().str.replace(r"[\x00-\x1F\x7F]", "", regex=True))
def pick_bodega_principal(stock_df):
    tmp = stock_df.groupby('Tienda', as_index=False)['Existencia'].sum()
    if not len(tmp): return None
    tmp['score'] = tmp['Tienda'].str.contains('BODEGA|CEDI|PRINCIPAL', case=False, regex=True).astype(int)
    tmp = tmp.sort_values(['score','Existencia'], ascending=[False, False])
    return tmp.iloc[0]['Tienda']

# ===== Carga =====
def load_ventas_y_stock(ventas_path, stock_path, ventas_sheet="Datos", stock_sheet="Sheet1"):
    ventas = pd.read_excel(ventas_path, sheet_name=ventas_sheet, engine="openpyxl")
    stock  = pd.read_excel(stock_path,  sheet_name=stock_sheet,  engine="openpyxl")
    print(f"[carga] ventas={len(ventas):,} (hoja='{ventas_sheet}')  stock={len(stock):,} (hoja='{stock_sheet}')")
    return ventas, stock

def load_tiendas(path_tien):
    if not path_tien: return {}, None
    try:
        ct = pd.read_excel(path_tien)
        ct_cols = {std_col(c): c for c in ct.columns}
        col_tienda    = ct_cols.get('tienda') or list(ct.columns)[0]
        col_tipo      = ct_cols.get('tipo')
        col_region    = ct_cols.get('region') or ct_cols.get('región')
        col_region_id = (ct_cols.get('region id') or ct_cols.get('region_id') or
                         ct_cols.get('regionid') or ct_cols.get('region code'))
        keep = [col_tienda] + ([col_tipo] if col_tipo else []) + ([col_region] if col_region else []) + ([col_region_id] if col_region_id else [])
        ct_df = ct[keep].copy()

        new_cols = ['Tienda']; 
        if col_tipo: new_cols.append('Tipo')
        if col_region: new_cols.append('Region')
        if col_region_id: new_cols.append('RegionID')
        ct_df.columns = new_cols

        ct_df['Tienda'] = ct_df['Tienda'].apply(normalize_store_name)
        if 'Tipo' in ct_df.columns: ct_df['Tipo'] = ct_df['Tipo'].astype(str).str.strip().str.upper()
        if 'RegionID' in ct_df.columns: ct_df['RegionID'] = pd.to_numeric(ct_df['RegionID'], errors='coerce').astype('Int64')
        tiendas_map = ct_df.set_index('Tienda').to_dict(orient='index')
        return tiendas_map, ct_df
    except Exception:
        return {}, None

def load_tiempos(path_tent):
    if not path_tent: return pd.DataFrame()
    try:
        tiempos_df = pd.read_excel(path_tent)
        if tiempos_df.empty: return pd.DataFrame()
        cols = {std_col(c): c for c in tiempos_df.columns}
        col_o   = cols.get('origen-destino') or cols.get('origen') or list(tiempos_df.columns)[0]
        col_d   = cols.get('destino-origen') or cols.get('destino') or list(tiempos_df.columns)[1]
        col_eta = (cols.get('eta') or cols.get('dias') or cols.get('días') or cols.get('tiempo (dias)') or list(tiempos_df.columns)[-2])
        col_pri = next((cols[k] for k in cols if 'priorid' in k or 'priori' in k), None)
        tiempos_df = tiempos_df.copy()
        tiempos_df['_O']        = tiempos_df[col_o].astype(str).str.upper().str.strip()
        tiempos_df['_D']        = tiempos_df[col_d].astype(str).str.upper().str.strip()
        tiempos_df['_ETA_NUM']  = tiempos_df[col_eta].apply(parse_lead_time_value)
        tiempos_df['_PRI_NUM']  = pd.to_numeric(tiempos_df[col_pri], errors='coerce') if col_pri else np.nan
        return tiempos_df
    except Exception:
        return pd.DataFrame()

# ===== Normalización / ADU / Curvas =====
def build_adu_from_ventas(ventas: pd.DataFrame) -> pd.DataFrame:
    vtc   = {std_col(c): c for c in ventas.columns}
    v_sku = vtc.get('sku')
    v_qty = (vtc.get('cantidad inv.') or vtc.get('cantidad') or vtc.get('unidades') or vtc.get('cantidad vendida'))
    v_tnd = (vtc.get('desc. bodega') or vtc.get('desc. c.o.') or vtc.get('tienda') or vtc.get('bodega'))
    v_date = detect_date_column(ventas)
    if not (v_sku and v_qty and v_tnd):
        return pd.DataFrame(columns=['Tienda','SKU','ADU'])

    v = ventas[[v_sku, v_qty, v_tnd] + ([v_date] if v_date else [])].copy()
    v.columns = ['SKU','Unidades','Tienda'] + (['Fecha'] if v_date else [])
    v['Tienda']   = v['Tienda'].apply(normalize_store_name)
    v['Unidades'] = pd.to_numeric(v['Unidades'], errors='coerce').fillna(0.0)

    if 'Fecha' in v.columns:
        v['Fecha'] = pd.to_datetime(v['Fecha'], errors='coerce')
        v = v[v['Fecha'].notna()].copy()
        v['Dia'] = v['Fecha'].dt.date
        dias_base_global = int(pd.Series(v['Dia'].unique()).nunique()) or 1
    else:
        dias_base_global = 30

    agg = v.groupby(['Tienda','SKU']).agg(total_units=('Unidades','sum')).reset_index()
    agg['ADU'] = agg['total_units'] / dias_base_global
    return agg[['Tienda','SKU','ADU']]

def normalize_stock(stock: pd.DataFrame, tiendas_map: dict | None = None) -> pd.DataFrame:
    stc = {std_col(c): c for c in stock.columns}
    c_sku   = stc.get('sku')
    c_exist = stc.get('existencia') or stc.get('existencias')
    c_tnd   = stc.get('desc. bodega') or stc.get('bodega') or stc.get('tienda')
    c_ref   = stc.get('referencia')
    c_talla = stc.get('talla') or stc.get('desc. detalle ext. 2')
    c_rango = stc.get('rango') or stc.get('rango_cat') or stc.get('rango raw') or stc.get('rango_raw')

    if c_sku is None or c_exist is None or c_tnd is None:
        raise ValueError("Stock: faltan columnas clave (SKU / Existencia / Tienda).")

    df = stock[[c_sku, c_exist, c_tnd] + ([c_ref] if c_ref else []) + ([c_talla] if c_talla else []) + ([c_rango] if c_rango else [])].copy()
    cols = ['SKU','Existencia','Tienda']
    if c_ref:   cols.append('Referencia')
    if c_talla: cols.append('Talla')
    if c_rango: cols.append('RANGO_RAW')
    df.columns = cols

    if 'Referencia' not in df.columns or 'Talla' not in df.columns:
        ref_parsed, talla_parsed = zip(*df['SKU'].apply(split_ref_talla_from_sku).tolist())
        if 'Referencia' not in df.columns: df['Referencia'] = ref_parsed
        if 'Talla' not in df.columns:      df['Talla']      = talla_parsed

    df['Tienda']     = df['Tienda'].apply(normalize_store_name)
    df['Existencia'] = pd.to_numeric(df['Existencia'], errors='coerce').fillna(0).round(0).astype(int)

    if 'RANGO_RAW' in df.columns:
        df['RANGO_CAT'] = df['RANGO_RAW'].apply(canonical_rango)
    else:
        df['RANGO_CAT'] = np.nan

    if tiendas_map:
        df['Tipo']     = df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Tipo'))
        df['Region']   = df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('Region'))
        df['RegionID'] = df['Tienda'].map(lambda t: tiendas_map.get(t, {}).get('RegionID'))
    else:
        df['Tipo'] = None; df['Region'] = None; df['RegionID'] = None

    if 'Referencia' in df.columns:
        df['Referencia'] = _clean_ref_col(df['Referencia'])

    return df

def filter_curves_safely(df: pd.DataFrame, disable: bool, debug: bool) -> pd.DataFrame:
    if disable:
        if debug: print("[curvas] filtro desactivado por --no-curve-filter")
        return df
    if 'RANGO_CAT' not in df.columns or 'Talla' not in df.columns:
        if debug: print("[curvas] columnas faltantes (RANGO_CAT/Talla); no se filtra.")
        return df
    valid_bebes = set(CURVAS_TALLAS['BEBES']); valid_ninos = set(CURVAS_TALLAS['NIÑOS'])
    mask_valid = ((df['RANGO_CAT']=='BEBES') & (df['Talla'].isin(valid_bebes))) | \
                 ((df['RANGO_CAT']=='NIÑOS') & (df['Talla'].isin(valid_ninos)))
    filtered = df[mask_valid].copy()
    if len(filtered)==0:
        if debug: print("[curvas] WARNING: filtro dejó 0 filas; se revierte a stock sin filtrar.")
        return df
    return filtered

# === ADICIÓN: Completar curva desde Bodega Principal (no intrusivo al flujo base) ===
# === ADICIÓN: Completar curva desde Bodega Principal priorizando por venta del SKU en la tienda ===
def completar_curva_desde_bodega(current_stock: pd.DataFrame,
                                 bodega_principal: str | None,
                                 transfers: list,
                                 debug: bool=False):
    """
    Completa curva (2 por talla) SOLO desde Bodega Principal hacia tiendas.
    Reglas:
      - Tiendas: categoría A/B/C embebida; desempate por ADU total (tienda).
      - Referencias: solo las que YA maneja la tienda; ordenadas por ADU de la referencia en esa tienda.
      - Tallas: SOLO las tallas (SKU exactos) con ADU > 0 en esa tienda (o sea, que la tienda SÍ las vende).
      - Respeta MAX_STOCK_PER_SKU. Bodega puede quedar en 0.
      - No introduce referencias nuevas en la tienda (pero puede crear la fila SKU destino si la talla no existía).
    """
    if not bodega_principal:
        if debug: print("[curva-bodega] no hay bodega principal; se omite.")
        return current_stock, transfers

    # Clasificación A/B/C embebida para priorizar tiendas (no toca tu lógica base)
    STORE_CATEGORY = {
        'BOGOTA PLAZA CENTRAL': 'C',
        'MEDELLIN FABRICATO': 'C',
        'NEIVA SAN PEDRO': 'B',
        'SABANETA MAYORCA': 'C',
        'BARRANQUILLA BUENAVISTA': 'B',
        'BARRANQUILLA PORTAL DEL PRADO': 'C',
        'BARRANQUILLA UNICO': 'A',
        'BARRANQUILLA VIVA': 'C',
        'CARTAGENA CARIBE PLAZA': 'B',
        'CARTAGENA MALL PLAZA': 'C',
        'CUCUTA UNICENTRO': 'C',
        'MONTERIA ALAMEDAS': 'C',
        'BUGA PLAZA': 'C',
        'CALI CHIPICHAPE': 'B',
        'CALI JARDIN PLAZA': 'A',
        'CALI MALL PLAZA': 'B',
        'CALI UNICENTRO': 'B',
        'CALI UNICO': 'A',
        'ECOMMERCE': 'A',
        'PALMIRA LLANOGRANDE': 'B',
        'POPAYAN CAMPANARIO': 'C',
        'TULUA LA HERRADURA': 'C',
    }
    def _rank_cat(name: str) -> int:
        return {'A':0,'B':1,'C':2}.get(STORE_CATEGORY.get(str(name).strip().upper(), 'C'), 3)

    # Para “solo tallas que venden”: umbral mínimo de ADU del SKU en tienda
    ADU_MIN_SKU = 1e-9

    # ADU total por tienda (para desempatar dentro de la categoría)
    adu_store = (current_stock[['Tienda','SKU','ADU']].drop_duplicates()
                 .groupby('Tienda')['ADU'].sum().to_dict())

    # Orden de tiendas (excluye bodega)
    tiendas = [t for t in current_stock['Tienda'].dropna().unique().tolist()
               if t != bodega_principal]
    tiendas.sort(key=lambda t: (_rank_cat(t), -float(adu_store.get(t, 0.0)), str(t)))

    # ADU por referencia en tienda (para priorizar referencias dentro de cada tienda)
    adu_ref_tienda = (current_stock[['Tienda','Referencia','SKU','ADU']].drop_duplicates()
                      .groupby(['Tienda','Referencia'])['ADU'].sum().reset_index())

    # Helper: ADU de un SKU (ref+talla) en tienda
    def _adu_sku(tienda, sku) -> float:
        r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
        return float(r['ADU'].iloc[0]) if len(r) else 0.0

    # Helper: ADU de referencia en tienda
    def _ref_adu(tienda, ref) -> float:
        v = adu_ref_tienda.loc[(adu_ref_tienda['Tienda']==tienda) & (adu_ref_tienda['Referencia']==ref),'ADU']
        return float(v.iloc[0]) if len(v) else 0.0

    BEBES = set(CURVAS_TALLAS.get('BEBES', []))
    NINOS = set(CURVAS_TALLAS.get('NIÑOS', []))

    def _rango_ref_en_tienda(tienda, ref) -> str | None:
        tallas_presentes = set(
            current_stock[(current_stock['Tienda']==tienda) & (current_stock['Referencia']==ref)]
            ['Talla'].dropna().astype(str).str.upper().tolist()
        )
        if tallas_presentes & BEBES: return 'BEBES'
        if tallas_presentes & NINOS: return 'NIÑOS'
        return None

    def _bodega_total():
        return int(current_stock[current_stock['Tienda']==bodega_principal]['Existencia'].sum())

    if debug:
        print(f"[curva-bodega] tiendas (top5): {[(t, _rank_cat(t), round(adu_store.get(t,0),2)) for t in tiendas[:5]]}")

    for tienda in tiendas:
        if _bodega_total() <= 0:
            if debug: print("[curva-bodega] bodega=0; fin global.")
            break

        # Solo referencias que YA maneja la tienda
        refs_tienda = (current_stock[current_stock['Tienda']==tienda]['Referencia']
                       .dropna().unique().tolist())
        # Orden por ADU de la referencia (desc)
        refs_tienda.sort(key=lambda r: -_ref_adu(tienda, r))

        for ref in refs_tienda:
            if _bodega_total() <= 0:
                if debug: print("[curva-bodega] bodega=0; fin global.")
                break

            rango = _rango_ref_en_tienda(tienda, ref)
            if rango is None:
                continue  # ceñirse a curvas establecidas

            curva = CURVAS_TALLAS[rango]
            # Ordenar tallas por ADU del SKU en esa tienda (desc) y filtrar tallas con ADU>0
            tallas_ord = sorted(
                [t for t in curva if _adu_sku(tienda, f"{ref}{t}") > ADU_MIN_SKU],
                key=lambda t: -_adu_sku(tienda, f"{ref}{t}")
            )
            if not tallas_ord:
                continue  # la tienda no tiene venta de ninguna talla de esta ref

            for talla in tallas_ord:
                if _bodega_total() <= 0:
                    if debug: print("[curva-bodega] bodega=0; fin global.")
                    break

                sku = f"{ref}{talla}"

                # Existencia actual en destino
                idx_d = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)].index
                dest_before = int(current_stock.loc[idx_d, 'Existencia'].sum()) if len(idx_d) else 0

                # Objetivo: 2 por talla, cap MAX_STOCK_PER_SKU
                need = max(0, min(MIN_POR_SKU - dest_before, MAX_STOCK_PER_SKU - dest_before))
                if need <= 0:
                    continue

                # Existencia en bodega para ese SKU
                idx_o = current_stock[(current_stock['Tienda']==bodega_principal) & (current_stock['SKU']==sku)].index
                bodega_before = int(current_stock.loc[idx_o, 'Existencia'].sum()) if len(idx_o) else 0
                if bodega_before <= 0:
                    continue

                qty = int(min(need, bodega_before))
                if qty <= 0:
                    continue

                # Si no existe fila destino para esta talla, créala (ref ya existe; talla puede ser nueva)
                if len(idx_d)==0:
                    new_row = {
                        'Tienda': tienda, 'SKU': sku, 'Referencia': ref, 'Talla': talla,
                        'RANGO_CAT': rango, 'Region': None, 'RegionID': None, 'Tipo': None,
                        'IsEcom': looks_like_ecom(tienda),
                        'MinObjetivo': MIN_ECOM if looks_like_ecom(tienda) else MIN_POR_SKU,
                        'ADU': _adu_sku(tienda, sku),
                        'Cobertura_dias': np.inf,
                        'Existencia': 0
                    }
                    current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
                    idx_d = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)].index

                # Actualizar existencias
                origin_before = bodega_before
                current_stock.loc[idx_o, 'Existencia'] -= qty / max(1, len(idx_o))
                current_stock.loc[idx_d, 'Existencia'] += qty / max(1, len(idx_d))
                origin_after = int(current_stock.loc[idx_o, 'Existencia'].sum())
                dest_after   = int(current_stock.loc[idx_d, 'Existencia'].sum())

                # Registrar traslado
                transfers.append({
                    'Tienda origen': bodega_principal, 'Tienda destino': tienda,
                    'Stock tienda origen antes traslado': origin_before,
                    'Stock tienda origen despues traslado': origin_after,
                    'Stock tienda destino antes traslado': dest_before,
                    'Stock tienda destino despues del traslado': dest_after,
                    'Unidades a trasladar': qty, 'Referencia': ref, 'Talla': talla
                })

    return current_stock, transfers

# === ADICIÓN: Drenaje inteligente desde Bodega por SKU (prioriza ADU del SKU en la tienda) ===
def drenar_residual_bodega_por_sku(current_stock: pd.DataFrame,
                                   bodega_principal: str | None,
                                   transfers: list,
                                   debug: bool=False):
    """
    Toma cada SKU con existencia en Bodega Principal y lo distribuye a tiendas
    que TENGAN ADU>0 DE ESE SKU (rotación real), priorizando:
      1) Categoría de tienda A/B/C (embebida abajo)
      2) ADU del SKU en esa tienda (desc)
    Respeta MAX_STOCK_PER_SKU en destino. No toca otras tiendas como origen.
    """
    if not bodega_principal:
        if debug: print("[drenaje] no hay bodega principal; omito drenaje.")
        return current_stock, transfers

    # Mapa A/B/C (solo para priorizar)
    STORE_CATEGORY = {
        'BOGOTA PLAZA CENTRAL': 'C',
        'MEDELLIN FABRICATO': 'C',
        'NEIVA SAN PEDRO': 'B',
        'SABANETA MAYORCA': 'C',
        'BARRANQUILLA BUENAVISTA': 'B',
        'BARRANQUILLA PORTAL DEL PRADO': 'C',
        'BARRANQUILLA UNICO': 'A',
        'BARRANQUILLA VIVA': 'C',
        'CARTAGENA CARIBE PLAZA': 'B',
        'CARTAGENA MALL PLAZA': 'C',
        'CUCUTA UNICENTRO': 'C',
        'MONTERIA ALAMEDAS': 'C',
        'BUGA PLAZA': 'C',
        'CALI CHIPICHAPE': 'B',
        'CALI JARDIN PLAZA': 'A',
        'CALI MALL PLAZA': 'B',
        'CALI UNICENTRO': 'B',
        'CALI UNICO': 'A',
        'ECOMMERCE': 'A',
        'PALMIRA LLANOGRANDE': 'B',
        'POPAYAN CAMPANARIO': 'C',
        'TULUA LA HERRADURA': 'C',
    }
    def _rank_cat(name: str) -> int:
        cat = STORE_CATEGORY.get(str(name).strip().upper(), 'C')
        return {'A':0,'B':1,'C':2}.get(cat, 3)

    # SKUs con stock en bodega (agregado)
    bod = (current_stock[current_stock['Tienda']==bodega_principal]
           .groupby(['SKU','Referencia','Talla'], as_index=False)['Existencia'].sum())
    bod = bod[bod['Existencia'] > 0].copy()
    if bod.empty:
        if debug: print("[drenaje] bodega ya está en 0; nada por hacer.")
        return current_stock, transfers

    # Priorizar SKUs más calientes (suma de ADU en todas las tiendas)
    sku_adu_sum = (current_stock[current_stock['Tienda']!=bodega_principal]
                   .groupby('SKU')['ADU'].sum().to_dict())
    bod['ADU_TOTAL'] = bod['SKU'].map(sku_adu_sum).fillna(0.0)
    bod = bod.sort_values('ADU_TOTAL', ascending=False)

    # Helpers
    def _stock(tienda, sku):
        r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
        return int(r['Existencia'].sum()) if len(r) else 0
    def _idx(tienda, sku):
        return current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)].index

    for _, row in bod.iterrows():
        sku, ref, talla = row['SKU'], row['Referencia'], row['Talla']
        rem = int(row['Existencia'])
        if rem <= 0:
            continue

        # Candidatos: tiendas (≠ bodega) que ya tienen fila de ese SKU y ADU>0
        cands = current_stock[
            (current_stock['SKU']==sku) &
            (current_stock['Tienda']!=bodega_principal)
        ][['Tienda','ADU']].drop_duplicates().copy()
        cands['ADU_SKU'] = cands['ADU'].fillna(0.0)
        cands = cands[cands['ADU_SKU'] > 0.0]
        if cands.empty:
            if debug: print(f"[drenaje] sin destinos con ADU>0 para SKU {sku}; se omite.")
            continue

        cands['CAT_RANK'] = cands['Tienda'].apply(_rank_cat)
        cands = cands.sort_values(['CAT_RANK','ADU_SKU','Tienda'], ascending=[True, False, True])

        # Greedy: llenar primero mejores candidatos hasta cap (=MAX_STOCK_PER_SKU)
        for _, c in cands.iterrows():
            if rem <= 0:
                break
            tienda = c['Tienda']
            dest_exist = _stock(tienda, sku)
            cap = max(0, MAX_STOCK_PER_SKU - dest_exist)
            if cap <= 0:
                continue

            qty = min(cap, rem)
            if qty <= 0:
                continue

            idx_o = _idx(bodega_principal, sku)
            idx_d = _idx(tienda, sku)

            origin_before = _stock(bodega_principal, sku)
            dest_before   = dest_exist

            # Actualiza existencias
            current_stock.loc[idx_o, 'Existencia'] -= qty / max(1, len(idx_o))
            if len(idx_d)==0:
                # si por alguna razón no hay fila (no debería ocurrir aquí), créala
                new_row = {
                    'Tienda': tienda, 'SKU': sku, 'Referencia': ref, 'Talla': talla,
                    'RANGO_CAT': None, 'Region': None, 'RegionID': None, 'Tipo': None,
                    'IsEcom': looks_like_ecom(tienda),
                    'MinObjetivo': MIN_ECOM if looks_like_ecom(tienda) else MIN_POR_SKU,
                    'ADU': float(c['ADU_SKU']), 'Cobertura_dias': np.inf, 'Existencia': 0
                }
                current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
                idx_d = _idx(tienda, sku)
            current_stock.loc[idx_d, 'Existencia'] += qty / max(1, len(idx_d))

            origin_after = _stock(bodega_principal, sku)
            dest_after   = _stock(tienda, sku)

            transfers.append({
                'Tienda origen': bodega_principal, 'Tienda destino': tienda,
                'Stock tienda origen antes traslado': origin_before,
                'Stock tienda origen despues traslado': origin_after,
                'Stock tienda destino antes traslado': dest_before,
                'Stock tienda destino despues del traslado': dest_after,
                'Unidades a trasladar': int(qty), 'Referencia': ref, 'Talla': talla
            })

            rem -= int(qty)

        if debug and rem > 0:
            print(f"[drenaje] SKU {sku}: remanente {rem} (sin capacidad/rotación en destinos).")

    return current_stock, transfers


# ===== Programa principal =====
def main():
    ap = argparse.ArgumentParser(description="Basecompleta (lee archivos preprocesados)")
    ap.add_argument("--ventas", required=True, help="Ruta a Ventas_procesadas_fmt.xlsx")
    ap.add_argument("--ventas-sheet", default="Datos", help="Hoja de ventas (default: Datos)")
    ap.add_argument("--stock", required=True, help="Ruta a Stock_procesado.xlsx")
    ap.add_argument("--stock-sheet", default="Sheet1", help="Hoja de stock (default: Sheet1)")
    ap.add_argument("--tiendas", default="", help="Clasificacion_Tiendas.xlsx (opcional)")
    ap.add_argument("--tiempos", default="", help="Tiempos de entrega.xlsx (opcional)")
    ap.add_argument("--out", default="Sugerencias_Traslados_Proyecto.xlsx", help="Salida final XLSX")
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--strict", action="store_true", help="Si se queda sin datos en pasos clave, sale con código 2.")
    ap.add_argument("--no-curve-filter", action="store_true", help="Desactiva el filtro por curvas (BEBES/NIÑOS).")
    args, _ = ap.parse_known_args()

    ventas_path = Path(args.ventas).resolve()
    stock_path  = Path(args.stock).resolve()
    if not ventas_path.is_file(): 
        print(f"[ERROR] No existe ventas: {ventas_path}"); sys.exit(2 if args.strict else 0)
    if not stock_path.is_file():  
        print(f"[ERROR] No existe stock:  {stock_path}"); sys.exit(2 if args.strict else 0)

    ventas, stock = load_ventas_y_stock(ventas_path, stock_path, args.ventas_sheet, args.stock_sheet)

    # sanity básico
    if len(ventas)==0 or len(stock)==0:
        print(f"[ERROR] ventas o stock vienen vacíos: ventas={len(ventas)}, stock={len(stock)}")
        sys.exit(2 if args.strict else 0)

    if args.debug:
        print("[debug] columnas ventas:", list(ventas.columns)[:25])
        print("[debug] columnas stock :", list(stock.columns)[:25])
        if "Referencia" in ventas.columns and "Referencia" in stock.columns:
            vset = set(_clean_ref_col(ventas["Referencia"]).dropna().unique())
            sset = set(_clean_ref_col(stock["Referencia"]).dropna().unique())
            print("[debug] intersección de refs:", len(vset & sset))

    tiendas_map, ct_df = load_tiendas(args.tiendas)
    tiempos_df = load_tiempos(args.tiempos)

    stock_df_raw = normalize_stock(stock, tiendas_map=tiendas_map)
    stock_df = filter_curves_safely(stock_df_raw, disable=args.no_curve_filter, debug=args.debug)

    if len(stock_df)==0:
        print("[ERROR] stock normalizado quedó vacío (aun sin filtro).")
        sys.exit(2 if args.strict else 0)

    # ADU & cobertura
    adu = build_adu_from_ventas(ventas)
    stock_df = stock_df.merge(adu, on=['Tienda','SKU'], how='left')
    stock_df['ADU'] = stock_df['ADU'].fillna(0.0)
    stock_df['Cobertura_dias'] = np.where(stock_df['ADU']>0, stock_df['Existencia']/stock_df['ADU'], np.inf)

    bodega_principal = pick_bodega_principal(stock_df)
    stock_df['IsEcom']      = stock_df['Tienda'].apply(looks_like_ecom)
    stock_df['MinObjetivo'] = np.where(stock_df['IsEcom'], MIN_ECOM, MIN_POR_SKU)
    if bodega_principal:
        stock_df.loc[stock_df['Tienda']==bodega_principal, 'MinObjetivo'] = 0

    # Snapshot
    key_cols = ['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias']
    for missing in ['RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias']:
        if missing not in stock_df.columns:
            stock_df[missing] = np.nan
    current_stock = stock_df[key_cols + ['Existencia']].copy()

    if len(current_stock)==0:
        print("[ERROR] current_stock vacío tras normalización.")
        sys.exit(2 if args.strict else 0)

    # Necesidades base
    if bodega_principal:
        mask_not_bodega = (current_stock['Tienda'] != bodega_principal)
    else:
        mask_not_bodega = pd.Series(True, index=current_stock.index)

    needs_df = current_stock[
        (current_stock['Existencia'] < current_stock['MinObjetivo']) &
        mask_not_bodega
    ].copy()
    if not needs_df.empty:
        needs_df['Necesita'] = (needs_df['MinObjetivo'] - needs_df['Existencia']).astype(int)
        needs_df['CapMax']   = MAX_STOCK_PER_SKU - needs_df['Existencia']
        needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
        needs_df = needs_df[needs_df['Necesita'] > 0].copy()
    else:
        needs_df = pd.DataFrame(columns=['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias','Necesita'])

    # Necesidades extra por curvas
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
            if faltan <= 0: continue
            candidatos = [(t, exists.get(t, 0)) for t in tallas if exists.get(t, 0) < MIN_POR_SKU]
            candidatos.sort(key=lambda x: x[1])
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
        if needs_df.empty:
            needs_df = extra.copy()
        else:
            needs_df = pd.concat([needs_df, extra], ignore_index=True)
        needs_df = needs_df.groupby(['Tienda','SKU','Referencia','Talla','RANGO_CAT','Region','RegionID','Tipo','IsEcom','MinObjetivo','ADU','Cobertura_dias'], as_index=False)['Necesita'].max()
        exist_now = current_stock.groupby(['Tienda','SKU'], as_index=False)['Existencia'].sum().rename(columns={'Existencia':'Exist_now'})
        needs_df = needs_df.merge(exist_now, on=['Tienda','SKU'], how='left')
        needs_df['CapMax'] = MAX_STOCK_PER_SKU - needs_df['Exist_now'].fillna(0)
        needs_df['Necesita'] = needs_df[['Necesita','CapMax']].min(axis=1).clip(lower=0).astype(int)
        needs_df = needs_df[needs_df['Necesita'] > 0].drop(columns=['Exist_now','CapMax'])

    if not needs_df.empty:
        needs_df = needs_df.sort_values(['ADU','Necesita'], ascending=[False, False]).reset_index(drop=True)

    # Mapa región por tienda
    def _norm(s): return str(s).strip().upper() if pd.notna(s) else None
    if 'RegionID' in current_stock.columns and 'Region' in current_stock.columns:
        store_meta_df = current_stock[['Tienda','RegionID','Region']].drop_duplicates('Tienda')
    else:
        store_meta_df = pd.DataFrame(columns=['Tienda','RegionID','Region'])
    _store_meta = { _norm(r['Tienda']): (str(int(r['RegionID'])) if pd.notna(r.get('RegionID')) else None,
                                          _norm(r.get('Region'))) for _, r in store_meta_df.iterrows() }

    def same_region_storelevel(origin_store, dest_store):
        o = _store_meta.get(_norm(origin_store), (None, None))
        d = _store_meta.get(_norm(dest_store), (None, None))
        o_rid, o_reg = o; d_rid, d_reg = d
        if o_rid and d_rid: return o_rid == d_rid
        if o_reg and d_reg: return o_reg == d_reg
        return False

    # Motor de traslados (puede resultar en 0 traslados y eso está bien)
    def get_stock(tienda, sku):
        r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
        return int(r['Existencia'].sum()) if len(r) else 0
    def get_cov_days(tienda, sku):
        r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
        if not len(r): return np.inf
        adu_val = float(r['ADU'].iloc[0]); ex = int(r['Existencia'].sum())
        return (ex/adu_val) if adu_val>0 else np.inf
    def allowed_to_send_from_origin(tienda, sku):
        if bodega_principal and tienda == bodega_principal:
            return get_stock(tienda, sku)
        r = current_stock[(current_stock['Tienda']==tienda) & (current_stock['SKU']==sku)]
        if not len(r): return 0
        ex   = int(r['Existencia'].sum())
        mino = int(r['MinObjetivo'].iloc[0]) if pd.notna(r['MinObjetivo'].iloc[0]) else MIN_POR_SKU
        adu_val  = float(r['ADU'].iloc[0]) if pd.notna(r['ADU'].iloc[0]) else 0.0
        is_e = bool(r['IsEcom'].iloc[0]) if pd.notna(r['IsEcom'].iloc[0]) else looks_like_ecom(tienda)
        min_cov = ORIGIN_MIN_COV_ECOM if is_e else ORIGIN_MIN_COV_DAYS
        guard  = int(np.ceil(min_cov * adu_val)) if adu_val>0 else mino
        guard  = max(guard, mino)
        return max(0, ex - guard)

    sku_to_origins = defaultdict(list)
    for _, r in current_stock.iterrows():
        tienda, sku = r['Tienda'], r['SKU']
        if allowed_to_send_from_origin(tienda, sku) > 0:
            sku_to_origins[sku].append(tienda)

    transfers = []
    if not needs_df.empty:
        for _, need in needs_df.iterrows():
            dest, sku, ref, talla = need['Tienda'], need['SKU'], need['Referencia'], need['Talla']
            if bodega_principal and dest == bodega_principal: continue
            dest_row  = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)]
            dest_adu  = float(dest_row['ADU'].iloc[0]) if len(dest_row) else 0.0
            dest_ecom = bool(dest_row['IsEcom'].iloc[0]) if len(dest_row) else looks_like_ecom(dest)
            dest_min  = int(dest_row['MinObjetivo'].iloc[0]) if len(dest_row) and pd.notna(dest_row['MinObjetivo'].iloc[0]) else (MIN_ECOM if dest_ecom else MIN_POR_SKU)

            target_cov_days  = DEST_TARGET_COV_ECOM if dest_ecom else DEST_TARGET_COV_DAYS
            target_units_cov = int(np.ceil(target_cov_days * dest_adu)) if dest_adu>0 else dest_min
            target_units     = min(MAX_STOCK_PER_SKU, max(dest_min, target_units_cov))

            origins = list(dict.fromkeys(([bodega_principal] if bodega_principal else []) + sku_to_origins.get(sku, [])))
            origins = [o for o in origins if o and o != dest]

            ranked = []
            for o in origins:
                can = allowed_to_send_from_origin(o, sku)
                if can <= 0: continue
                o_cov = get_cov_days(o, sku); d_cov = get_cov_days(dest, sku)
                if np.isfinite(o_cov) and np.isfinite(d_cov) and (o_cov <= d_cov + COV_BUFFER_DAYS): continue
                same_reg = same_region_storelevel(o, dest)
                pri = delivery_priority(o, dest, tiempos_df); pri_sort = int(pri) if pd.notna(pri) else 999
                lt  = delivery_days(o, dest, tiempos_df);     lt_sort  = float(lt) if pd.notna(lt) else 999
                ranked.append((0 if same_reg else 1, pri_sort, -(o_cov if np.isfinite(o_cov) else 1e9), lt_sort, o))
            ranked.sort(key=lambda t: (t[0], t[1], t[2], t[3]))

            while True:
                dest_exist_now = get_stock(dest, sku)
                gap_to_target  = max(0, target_units - dest_exist_now)
                if gap_to_target <= 0: break
                moved = False
                for _, _, _, _, o in ranked:
                    can_now = allowed_to_send_from_origin(o, sku)
                    if can_now <= 0: continue
                    origin_before = get_stock(o, sku); dest_before = dest_exist_now
                    gap_to_cap = max(0, MAX_STOCK_PER_SKU - dest_before)
                    qty = int(min(can_now, gap_to_target, gap_to_cap))
                    if qty <= 0: continue
                    origin_after = origin_before - qty; dest_after = dest_before + qty
                    transfers.append({
                        'Tienda origen': o,'Tienda destino': dest,
                        'Stock tienda origen antes traslado': origin_before,'Stock tienda origen despues traslado': origin_after,
                        'Stock tienda destino antes traslado': dest_before,'Stock tienda destino despues del traslado': dest_after,
                        'Unidades a trasladar': qty, 'Referencia': ref,'Talla': talla
                    })
                    idx_o = current_stock[(current_stock['Tienda']==o) & (current_stock['SKU']==sku)].index
                    current_stock.loc[idx_o, 'Existencia'] -= qty / max(1,len(idx_o))
                    idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
                    if len(idx_d)==0:
                        ref_d, talla_d = split_ref_talla_from_sku(sku)
                        new_row = {'Tienda': dest,'SKU': sku,'Referencia': ref_d,'Talla': talla_d,
                                   'RANGO_CAT': None,'Region': None,'RegionID': None,'Tipo': None,
                                   'IsEcom': looks_like_ecom(dest),'MinObjetivo': MIN_ECOM if looks_like_ecom(dest) else MIN_POR_SKU,
                                   'ADU': dest_adu,'Cobertura_dias': np.inf,'Existencia': 0}
                        current_stock = pd.concat([current_stock, pd.DataFrame([new_row])], ignore_index=True)
                        idx_d = current_stock[(current_stock['Tienda']==dest) & (current_stock['SKU']==sku)].index
                    current_stock.loc[idx_d, 'Existencia'] += qty / max(1,len(idx_d))
                    moved = True
                    break
                if not moved: break

    current_stock, transfers = completar_curva_desde_bodega(
    current_stock=current_stock,
    bodega_principal=bodega_principal,
    transfers=transfers,
    debug=args.debug
    )
    # === ADICIÓN: Drenaje final para vaciar bodega según rotación por SKU ===
    current_stock, transfers = drenar_residual_bodega_por_sku(
    current_stock=current_stock,
    bodega_principal=bodega_principal,
    transfers=transfers,
    debug=args.debug
    )

    current_stock['Existencia'] = current_stock['Existencia'].round().astype(int)
    current_stock.loc[current_stock['Existencia']<0, 'Existencia'] = 0

    traslados_df = pd.DataFrame(transfers, columns=[
        'Tienda origen','Tienda destino',
        'Stock tienda origen antes traslado','Stock tienda origen despues traslado',
        'Stock tienda destino antes traslado','Stock tienda destino despues del traslado',
        'Unidades a trasladar','Referencia','Talla'
    ])
    stock_despues = (current_stock.groupby(['Tienda','SKU','Referencia','Talla'], as_index=False)['Existencia']
                     .sum().sort_values(['Tienda','Referencia','Talla']))

    # Validaciones finales (para no dejar XLSX “en blanco” en modo estricto)
    if args.strict:
        if len(stock_despues)==0:
            print("[ERROR] stock_despues vacío: no hay datos para escribir."); sys.exit(2)
        # Traslados puede ser 0 y no es error, pero avisamos:
        if len(traslados_df)==0:
            print("[strict] aviso: no hay traslados. Se escribirá solo 'Stock_despues'.")

    out_path = Path(args.out).resolve()
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        traslados_df.to_excel(writer, sheet_name="Traslados", index=False)
        stock_despues.to_excel(writer, sheet_name="Stock_despues", index=False)

    if args.debug:
        print(f"[out] Traslados: {len(traslados_df)} filas")
        print(f"[out] Stock_despues: {len(stock_despues)} filas")
    print(f"[OK] Archivo final -> {out_path}")

if __name__ == "__main__":
    main()
