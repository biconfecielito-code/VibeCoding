import pandas as pd
from pathlib import Path

# ----------------------------
# Utilidades
# ----------------------------
def _to_int64_nullable(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce")
    # redondeo por seguridad si viniera con decimales
    return s.round(0).astype("Int64")

def _clean_ref_col(series: pd.Series) -> pd.Series:
    return (series.astype("string")
                  .str.strip()
                  .str.replace(r"[\x00-\x1F\x7F]", "", regex=True))

# ----------------------------
# Pipeline principal
# ----------------------------
def procesar_stock(
    stock_path: str | Path,
    stock_sheet: str = "Sheet1",
    venta_df: pd.DataFrame | None = None,          # Debe tener columna 'Referencia'
    venta_path: str | None = None,
    venta_sheet: str | None = None,
    seleccion_df: pd.DataFrame | None = None,      # Debe tener columna 'Referencias'
    seleccion_path: str | None = None,
    seleccion_sheet: str | None = None,
) -> pd.DataFrame:
    """
    Replica en pandas el código M proporcionado para Stock.xlsx.
    """
    # 1) Cargar
    df = pd.read_excel(stock_path, sheet_name=stock_sheet, engine="openpyxl")

    # 2) Referencia no empieza por "N"
    df = df[~df["Referencia"].astype(str).str.startswith("N")].copy()

    # 3) Asegurar texto en Referencia
    df["Referencia"] = df["Referencia"].astype("string")

    # 4) SKU = Referencia & Desc. detalle ext. 2  (y renombrar a 'SKU')
    df["SKU"] = (df["Referencia"].fillna("") + df["Desc. detalle ext. 2"].fillna("")).astype("string")

    # 5) Quitar columna 'Existencia' original (si existe)
    df = df.drop(columns=[c for c in ["Existencia"] if c in df.columns], errors="ignore")

    # 6) Nueva 'Existencia' = Cant. disponible + Cant. transito ent.
    #    (primero hacemos numérico y llenamos nulos con 0)
    disp = pd.to_numeric(df.get("Cant. disponible", pd.NA), errors="coerce").fillna(0)
    trans = pd.to_numeric(df.get("Cant. transito ent.", pd.NA), errors="coerce").fillna(0)
    df["Existencia"] = disp + trans

    # 7) Eliminar columnas source de la suma
    df = df.drop(columns=[c for c in ["Cant. disponible", "Cant. transito ent."] if c in df.columns], errors="ignore")

    # 8) Renombrar 'Desc. detalle ext. 2' -> 'Talla'
    if "Desc. detalle ext. 2" in df.columns:
        df = df.rename(columns={"Desc. detalle ext. 2": "Talla"})

    # 9) Quitar 'CLASIFICACION' si existe
    df = df.drop(columns=[c for c in ["CLASIFICACION"] if c in df.columns], errors="ignore")

    # 10) Filtro por Desc. bodega (lista blanca)
    bodegas_ok = {
        "BARRANQUILLA BUENAVISTA", "BARRANQUILLA PORTAL DEL PRADO", "BARRANQUILLA UNICO",
        "BARRANQUILLA VIVA", "BODEGA ECOMMERCE", "BODEGA PRINCIPAL", "BOGOTA PLAZA CENTRAL",
        "BUGA PLAZA", "CALI CHIPICHAPE", "CALI JARDIN PLAZA", "CALI UNICENTRO", "CALI UNICO",
        "CARTAGENA CARIBE PLAZA", "CUCUTA UNICENTRO", "ECOMMERCE", "MONTERIA ALAMEDAS",
        "NEIVA SAN PEDRO", "PALMIRA LLANOGRANDE", "POPAYAN CAMPANARIO", "SABANETA MAYORCA",
        "TULUA LA HERRADURA"
    }
    if "Desc. bodega" in df.columns:
        df = df[df["Desc. bodega"].isin(bodegas_ok)].copy()

    # 11) Asegurar SKU como texto
    df["SKU"] = df["SKU"].astype("string")

    # 12) Filtrar: Referencia no contiene "PROMO"
    df = df[~df["Referencia"].fillna("").str.contains("PROMO", na=False)].copy()

    # 13) Filtrar: Referencia no empieza por "S"
    df = df[~df["Referencia"].astype(str).str.startswith("S")].copy()

    # 14) Trim + Clean en Referencia
    df["Referencia"] = _clean_ref_col(df["Referencia"])

    # 15) Join interno con Venta (por Referencia)
    if venta_df is None:
        if not (venta_path and venta_sheet):
            raise ValueError("Proporciona venta_df o (venta_path y venta_sheet).")
        venta_df = pd.read_excel(venta_path, sheet_name=venta_sheet, engine="openpyxl")
    if "Referencia" not in venta_df.columns:
        raise KeyError("La tabla 'Venta' debe contener la columna 'Referencia'.")
    refs_venta = _clean_ref_col(venta_df["Referencia"])
    df = df[df["Referencia"].isin(refs_venta)].copy()

    # 16) Convertir 'Existencia' a entero nullable (Int64)
    df["Existencia"] = _to_int64_nullable(df["Existencia"])

    # 17) Join interno con Seleccion (Referencia ∈ Seleccion.Referencias)
  # 17) (Opcional) Join interno con Seleccion (Referencia ∈ Seleccion.Referencias o Seleccion.Referencia)
    if (seleccion_df is not None) or (seleccion_path and seleccion_sheet):
        if seleccion_df is None:
            seleccion_df = pd.read_excel(seleccion_path, sheet_name=(seleccion_sheet or 0), engine="openpyxl")
        # Acepta 'Referencias' o 'Referencia'
        if "Referencias" in seleccion_df.columns:
            col_sel = "Referencias"
        elif "Referencia" in seleccion_df.columns:
            col_sel = "Referencia"
        else:
            raise KeyError("La tabla 'Seleccion' debe contener la columna 'Referencias' o 'Referencia'.")
        refs_sel = _clean_ref_col(seleccion_df[col_sel])
        df = df[df["Referencia"].isin(refs_sel)].copy()
# Si no se pasa selección, no se filtra por selección y seguimos.
    # Reorden opcional (para dejar lo relevante al frente si existe)
    desired = ["Referencia", "SKU", "Talla", "Existencia", "Desc. bodega"]
    front = [c for c in desired if c in df.columns]
    rest  = [c for c in df.columns if c not in front]
    df = df[front + rest]

    return df

# ----------------------------
# Ejemplo de uso
# ----------------------------
if __name__ == "__main__":
    import argparse
    from pathlib import Path

    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="stock_path", required=True, help="Excel de stock crudo")
    ap.add_argument("--sheet", dest="stock_sheet", default="Sheet1", help="Hoja de stock (default: Sheet1)")
    ap.add_argument("--out", dest="out_path", default="Stock_procesado.xlsx", help="Salida")
    ap.add_argument("--ventas", dest="venta_path", default=None, help="(Opcional) Ventas para filtrar referencias")
    ap.add_argument("--ventas-sheet", dest="venta_sheet", default="Datos", help="Hoja de ventas (default: Datos)")
    ap.add_argument("--seleccion", dest="seleccion_path", default=None, help="(Opcional) Excel con referencias/sku")
    ap.add_argument("--seleccion-sheet", dest="seleccion_sheet", default=None, help="Hoja de selección (default: primera)")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    # Llama a tu función existente
    df_final = procesar_stock(
        stock_path=args.stock_path,
        stock_sheet=args.stock_sheet,
        venta_path=args.venta_path,
        venta_sheet=args.venta_sheet,
        seleccion_path=args.seleccion_path,
        seleccion_sheet=args.seleccion_sheet
    )

    out_path = Path(args.out_path).resolve()
    df_final.to_excel(out_path, index=False)
    print(f"OK -> {out_path}")
