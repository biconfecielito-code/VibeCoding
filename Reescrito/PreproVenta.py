# PreproVenta.py — Normaliza Talla/SKU (strip + upper), selección opcional y exporta a 'Datos'
import pandas as pd
from pathlib import Path
import re, unicodedata
import argparse

# ----------------------------
# Utilidades
# ----------------------------
def _norm_txt(s: str) -> str:
    s = str(s).strip().lower()
    s = ''.join(ch for ch in unicodedata.normalize('NFD', s)
                if unicodedata.category(ch) != 'Mn')
    return re.sub(r'[^a-z0-9]+', '', s)

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {_norm_txt(c): c for c in df.columns.astype(str)}
    for cand in candidates:
        key = _norm_txt(cand)
        if key in norm_map:
            return norm_map[key]
    return None

def to_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def format_fecha_sin_hora(df: pd.DataFrame, col: str = "Fecha") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    return df

def build_sku(df: pd.DataFrame,
              ref_col: str = "Referencia",
              talla_src: str = "Desc. detalle ext. 2",
              sku_col: str = "SKU") -> pd.DataFrame:
    df[sku_col] = (df.get(ref_col, "").fillna("").astype(str)
                   + df.get(talla_src, "").fillna("").astype(str))
    return df

def filtrar_por_seleccion(
    df: pd.DataFrame,
    seleccion_df: pd.DataFrame | None = None,
    seleccion_path: str | None = None,
    seleccion_sheet: str | None = None,
    debug: bool = False,
) -> pd.DataFrame:
    """
    Si hay selección (DataFrame o archivo), filtra por Referencia ∈ Seleccion.Referencias o Seleccion.Referencia.
    Si NO hay selección, NO filtra (devuelve df intacto).
    """
    if seleccion_df is None and not (seleccion_path and seleccion_sheet):
        if debug: print("[VENTAS] Sin selección: no se filtra por referencias.")
        return df

    if seleccion_df is None:
        try:
            seleccion_df = pd.read_excel(seleccion_path, sheet_name=(seleccion_sheet or 0), engine="openpyxl")
        except Exception as e:
            if debug: print(f"[VENTAS] No se pudo leer selección ({e}); no se filtra.")
            return df

    col_sel = "Referencias" if "Referencias" in seleccion_df.columns else \
              ("Referencia" if "Referencia" in seleccion_df.columns else None)
    if not col_sel:
        if debug: print("[VENTAS] Selección sin columna 'Referencias'/'Referencia'; no se filtra.")
        return df

    refs = (seleccion_df[col_sel].astype("string")
            .str.strip().str.replace(r"[\x00-\x1F\x7F]", "", regex=True))
    refs = refs[refs.notna() & (refs != "")]
    if len(refs) == 0:
        if debug: print("[VENTAS] Selección vacía; no se filtra.")
        return df

    out = df[df["Referencia"].astype(str).isin(refs.astype(str).unique())].copy()
    if debug: print(f"[VENTAS] Filtrado por selección: {len(out)}/{len(df)} filas.")
    return out

def exportar_xlsx(df: pd.DataFrame, out_path: str | Path,
                  add_resumen: bool = True,
                  fecha_col: str = "Fecha",
                  valor_col: str = "Valor neto") -> None:
    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
        wb  = writer.book
        ws  = writer.sheets["Datos"]
        fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
        fmt_int  = wb.add_format({"num_format": "0"})
        cols = {c: i for i, c in enumerate(df.columns)}
        if fecha_col in cols: ws.set_column(cols[fecha_col], cols[fecha_col], None, fmt_date)
        if valor_col in cols: ws.set_column(cols[valor_col], cols[valor_col], None, fmt_int)
        if add_resumen:
            resumen = pd.DataFrame({
                "Métrica": ["Filas", f"Suma {valor_col}"],
                "Valor":  [len(df), df[valor_col].fillna(0).sum() if valor_col in df.columns else 0]
            })
            resumen.to_excel(writer, index=False, sheet_name="Resumen")

# ----------------------------
# Pipeline principal
# ----------------------------
def cargar_y_transformar(ventas_path: str | Path,
                         ventas_sheet: str = "Sheet1",
                         seleccion_df: pd.DataFrame | None = None,
                         seleccion_path: str | None = None,
                         seleccion_sheet: str | None = None,
                         debug: bool = False) -> pd.DataFrame:
    """Replica el flujo de Power Query; normaliza Talla/SKU para evitar mismatch."""
    df = pd.read_excel(ventas_path, sheet_name=ventas_sheet, engine="openpyxl")

    # 1) Filtros iniciales
    col_clas = find_col(df, ["CLASIFICACION", "Clasificación"])
    if col_clas:
        objetivo = _norm_txt("6301 - PRENDAS")
        mask = df[col_clas].astype(str).map(_norm_txt).eq(objetivo)
        df = df[mask].copy()
    else:
        if debug: print("[VENTAS] Aviso: no se encontró CLASIFICACION/Clasificación; omito ese filtro.")

    df = df[~df["Referencia"].astype(str).str.startswith("N")].copy()

    # 2) Normalizar numéricas
    money_cols = ["Precio unit.", "Valor bruto", "Valor descuentos", "Valor subtotal", "Valor neto"]
    df = to_numeric(df, money_cols)

    # 3) Eliminar columnas no usadas
    df = df.drop(columns=[c for c in ["Razón social cliente factura", "Costo promedio total", "Estado"]
                          if c in df.columns], errors="ignore")

    # 4) SKU (provisional antes de renombrar talla)
    df = build_sku(df, ref_col="Referencia", talla_src="Desc. detalle ext. 2", sku_col="SKU")

    # 5) Reemplazos de tienda (si corresponde)
    if "Desc. C.O." in df.columns:
        df["Desc. C.O."] = df["Desc. C.O."].astype("string").str.replace("PRINCIPAL", "ECOMMERCE", regex=False)

    # 6) Quitar columnas extra (como en M)
    df = df.drop(columns=[c for c in ["Nro documento", "Precio unit.", "Valor bruto",
                                      "Valor descuentos", "Valor subtotal", "CLASIFICACION", "SUBLINEA"]
                          if c in df.columns], errors="ignore")

    # 7) Renombrar talla
    if "Desc. detalle ext. 2" in df.columns:
        df = df.rename(columns={"Desc. detalle ext. 2": "Talla"})

    # 8) Eliminar columnas
    df = df.drop(columns=[c for c in ["GENERO", "CAPSULA"] if c in df.columns], errors="ignore")

    # 9) Filtrar PROMO
    df = df[~df["Referencia"].fillna("").astype(str).str.contains("PROMO", na=False)].copy()

    # 10) Limpiar referencia
    df["Referencia"] = (df["Referencia"].astype("string")
                        .str.strip().str.replace(r"[\x00-\x1F\x7F]", "", regex=True))

    # === PARCHE CLAVE: Normalizar Talla y SKU (strip + upper) ===
    if "Talla" in df.columns:
        df["Talla"] = df["Talla"].astype("string").str.strip().str.upper()
    if "SKU" in df.columns:
        # Reconstruir SKU ya normalizado (Ref + Talla)
        df["SKU"] = (df["Referencia"].fillna("").astype(str).str.strip()
                     + df.get("Talla", "").astype(str).str.strip().str.upper())

    # 11) Fecha sin hora
    df = format_fecha_sin_hora(df, col="Fecha")

    # 12) Valor neto entero
    if "Valor neto" in df.columns:
        df["Valor neto"] = pd.to_numeric(df["Valor neto"], errors="coerce").round(0).astype("Int64")

    # 13) (Opcional) Filtrar por Selección (inner)
    df = filtrar_por_seleccion(
        df,
        seleccion_df=seleccion_df,
        seleccion_path=seleccion_path,
        seleccion_sheet=seleccion_sheet,
        debug=debug,
    )

    # 14) Reorden opcional
    desired = ["C.O.", "Bodega", "Desc. C.O.", "Fecha", "Referencia", "Desc. item",
               "Talla", "Cantidad inv.", "Valor neto", "RANGO", "SKU"]
    front = [c for c in desired if c in df.columns]
    rest  = [c for c in df.columns if c not in front]
    df = df[front + rest]

    return df

# ----------------------------
# CLI
# ----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="ventas_path", required=True, help="Excel de ventas crudo")
    ap.add_argument("--sheet", dest="ventas_sheet", default="Sheet1", help="Hoja de ventas (default: Sheet1)")
    ap.add_argument("--out", dest="out_path", default="Ventas_procesadas_fmt.xlsx", help="Salida")
    ap.add_argument("--seleccion", dest="seleccion_path", default=None, help="(Opcional) Excel con referencias a incluir")
    ap.add_argument("--seleccion-sheet", dest="seleccion_sheet", default=None, help="Hoja de selección (default: primera)")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    df_final = cargar_y_transformar(
        ventas_path=args.ventas_path,
        ventas_sheet=args.ventas_sheet,
        seleccion_df=None,
        seleccion_path=args.seleccion_path,
        seleccion_sheet=args.seleccion_sheet,
        debug=args.debug,
    )

    out_path = Path(args.out_path).resolve()
    exportar_xlsx(df_final, out_path, add_resumen=True)
    print(f"OK -> {out_path}")

if __name__ == "__main__":
    main()
