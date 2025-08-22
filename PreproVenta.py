import pandas as pd
from pathlib import Path

# ----------------------------
# Utilidades
# ----------------------------
def to_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """Convierte a número (float) de forma segura las columnas indicadas."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def format_fecha_sin_hora(df: pd.DataFrame, col: str = "Fecha") -> pd.DataFrame:
    """Convierte a datetime y deja solo fecha (date)."""
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    return df

def build_sku(df: pd.DataFrame,
              ref_col: str = "Referencia",
              talla_src: str = "Desc. detalle ext. 2",
              sku_col: str = "SKU") -> pd.DataFrame:
    """Crea el SKU concatenando Referencia + talla/origen."""
    df[sku_col] = (df.get(ref_col, "").fillna("").astype(str)
                   + df.get(talla_src, "").fillna("").astype(str))
    return df

def filtrar_por_seleccion(df: pd.DataFrame,
                          seleccion_df: pd.DataFrame | None = None,
                          seleccion_path: str | None = None,
                          seleccion_sheet: str | None = None) -> pd.DataFrame:
    """Aplica inner-filter por Referencia ∈ Seleccion.Referencias."""
    if seleccion_df is None:
        if not (seleccion_path and seleccion_sheet):
            raise ValueError("Proporciona seleccion_df o (seleccion_path y seleccion_sheet).")
        seleccion_df = pd.read_excel(seleccion_path, sheet_name=seleccion_sheet, engine="openpyxl")

    if "Referencias" not in seleccion_df.columns:
        raise KeyError("La tabla de selección debe contener la columna 'Referencias'.")

    refs = (seleccion_df["Referencias"].astype("string")
            .str.strip().str.replace(r"[\x00-\x1F\x7F]", "", regex=True))
    return df[df["Referencia"].isin(refs)].copy()

def exportar_xlsx(df: pd.DataFrame, out_path: str | Path,
                  add_resumen: bool = True,
                  fecha_col: str = "Fecha",
                  valor_col: str = "Valor neto") -> None:
    """Exporta a XLSX; opcionalmente agrega hoja 'Resumen' y formatos."""
    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        # Hoja principal
        df.to_excel(writer, index=False, sheet_name="Datos")
        wb  = writer.book
        ws  = writer.sheets["Datos"]

        # Formatos
        fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
        fmt_int  = wb.add_format({"num_format": "0"})

        # Aplicar formato a columnas si existen
        cols = {c: i for i, c in enumerate(df.columns)}
        if fecha_col in cols:
            ws.set_column(cols[fecha_col], cols[fecha_col], None, fmt_date)
        if valor_col in cols:
            ws.set_column(cols[valor_col], cols[valor_col], None, fmt_int)

        # Hoja Resumen
        if add_resumen:
            resumen = pd.DataFrame({
                "Métrica": ["Filas", f"Suma {valor_col}"],
                "Valor":  [len(df), df[valor_col].fillna(0).sum()]
            })
            resumen.to_excel(writer, index=False, sheet_name="Resumen")

# ----------------------------
# Pipeline principal
# ----------------------------
def cargar_y_transformar(ventas_path: str | Path,
                         ventas_sheet: str = "Sheet1",
                         seleccion_df: pd.DataFrame | None = None,
                         seleccion_path: str | None = None,
                         seleccion_sheet: str | None = None) -> pd.DataFrame:
    """Replica el flujo de Power Query, con fecha sin hora y Valor neto entero."""
    df = pd.read_excel(ventas_path, sheet_name=ventas_sheet, engine="openpyxl")

    # 1) Filtros iniciales
    df = df[df["CLASIFICACION"] == "6301 - PRENDAS"].copy()
    df = df[~df["Referencia"].astype(str).str.startswith("N")].copy()

    # 2) Normalizar numéricas (sin dividir entre 100)
    money_cols = ["Precio unit.", "Valor bruto", "Valor descuentos", "Valor subtotal", "Valor neto"]
    df = to_numeric(df, money_cols)

    # 3) Eliminar columnas
    df = df.drop(columns=[c for c in ["Razón social cliente factura", "Costo promedio total", "Estado"]
                          if c in df.columns],
                 errors="ignore")

    # 4) SKU
    df = build_sku(df, ref_col="Referencia", talla_src="Desc. detalle ext. 2", sku_col="SKU")

    # 5) Reemplazos
    if "Desc. C.O." in df.columns:
        df["Desc. C.O."] = df["Desc. C.O."].astype("string").str.replace("PRINCIPAL", "ECOMMERCE", regex=False)

    # 6) Quitar columnas extra (como en M)
    df = df.drop(columns=[c for c in ["Nro documento", "Precio unit.", "Valor bruto",
                                      "Valor descuentos", "Valor subtotal", "CLASIFICACION", "SUBLINEA"]
                          if c in df.columns],
                 errors="ignore")

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

    # 11) Fecha sin hora
    df = format_fecha_sin_hora(df, col="Fecha")

    # 12) Valor neto entero (redondeo al cercano)
    if "Valor neto" in df.columns:
        df["Valor neto"] = pd.to_numeric(df["Valor neto"], errors="coerce").round(0).astype("Int64")

    # 13) Filtrar por Selección (inner)
    df = filtrar_por_seleccion(
        df,
        seleccion_df=seleccion_df,
        seleccion_path=seleccion_path,
        seleccion_sheet=seleccion_sheet
    )

    # 14) Reorden opcional (mantiene lo principal al frente si existe)
    desired = ["C.O.", "Bodega", "Desc. C.O.", "Fecha", "Referencia", "Desc. item",
               "Talla", "Cantidad inv.", "Valor neto", "RANGO", "SKU"]
    front = [c for c in desired if c in df.columns]
    rest  = [c for c in df.columns if c not in front]
    df = df[front + rest]

    return df

# ----------------------------
# Ejemplo de uso
# ----------------------------
if __name__ == "__main__":
    ventas_path = "Ventas.xlsx"      # <-- cambia a tu ruta
    # Opción A: DataFrame de selección
    seleccion = pd.DataFrame({
        "Referencias": [
            "2485279","2485351","2185326","2485462","2585274","2485325","2585270",
            "1185437","2185449","2585353","2585327","2485329","2184989","1485456",
            "2485022","1485442","1485464","1185436","2485263","2485348"
        ]
    })

    df_final = cargar_y_transformar(
        ventas_path=ventas_path,
        ventas_sheet="Sheet1",
        seleccion_df=seleccion,                 # Opción A
        # seleccion_path=r"/ruta/a/Seleccion.xlsx", seleccion_sheet="Seleccion"  # Opción B
    )

    out_path = Path(ventas_path).with_name("Ventas_procesadas_fmt.xlsx")
    exportar_xlsx(df_final, out_path, add_resumen=True)
    print(f"OK -> {out_path}")
