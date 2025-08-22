import pandas as pd
from pathlib import Path
from configparser import ConfigParser
from typing import Literal

def load_ini(path: str | Path) -> ConfigParser:
    cfg = ConfigParser()
    with open(path, "r", encoding="utf-8") as f:
        cfg.read_file(f)
    return cfg

def parse_inline_refs(raw: str) -> list[str]:
    items = []
    for chunk in raw.replace("\n", ",").split(","):
        s = chunk.strip()
        if s:
            items.append(s)
    return items

def load_selection_from_cfg(cfg: ConfigParser) -> tuple[str, list[str]]:
    mode = cfg.get("seleccion", "mode", fallback="auto").strip().lower()
    source = cfg.get("seleccion", "source", fallback="inline").strip().lower()

    refs: list[str] = []
    if source == "inline":
        raw = cfg.get("seleccion", "refs_inline", fallback="")
        refs = parse_inline_refs(raw)
    elif source == "file":
        file_path = cfg.get("seleccion", "refs_file", fallback="")
        if file_path:
            p = Path(file_path)
            if p.is_file():
                txt = p.read_text(encoding="utf-8")
                refs = parse_inline_refs(txt)

    # Sanitizar
    refs = [r.strip() for r in refs if r and r.strip()]

    # Fallback: si no hay refs, no filtrar (usar all)
    if mode in ("subset", "auto") and len(refs) == 0:
        mode = "all"
    return mode, refs

def _clean_ref(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"[\x00-\x1F\x7F]", "", regex=True)
    )

def procesar_ventas(
    ventas_path: str | Path,
    ventas_sheet: str = "Sheet1",
    seleccion_mode: Literal["all", "subset", "auto"] = "auto",
    seleccion_refs: list[str] | None = None,
) -> pd.DataFrame:
    df = pd.read_excel(ventas_path, sheet_name=ventas_sheet, engine="openpyxl")

    # === Pipeline base (equivalente a Power Query anterior) ===
    df = df[df["CLASIFICACION"] == "6301 - PRENDAS"].copy()
    df = df[~df["Referencia"].astype(str).str.startswith("N")].copy()

    money_cols = ["Precio unit.", "Valor bruto", "Valor descuentos", "Valor subtotal", "Valor neto"]
    for c in money_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.drop(
        columns=[c for c in ["Razón social cliente factura", "Costo promedio total", "Estado"] if c in df.columns],
        errors="ignore",
    )

    df["SKU"] = (
        df["Referencia"].fillna("").astype(str) + df["Desc. detalle ext. 2"].fillna("").astype(str)
    )

    if "Desc. C.O." in df.columns:
        df["Desc. C.O."] = df["Desc. C.O."].astype("string").str.replace("PRINCIPAL", "ECOMMERCE", regex=False)

    df = df.drop(
        columns=[c for c in ["Nro documento", "Precio unit.", "Valor bruto", "Valor descuentos",
                             "Valor subtotal", "CLASIFICACION", "SUBLINEA"] if c in df.columns],
        errors="ignore",
    )

    if "Desc. detalle ext. 2" in df.columns:
        df = df.rename(columns={"Desc. detalle ext. 2": "Talla"})

    df = df.drop(columns=[c for c in ["GENERO", "CAPSULA"] if c in df.columns], errors="ignore")

    df = df[~df["Referencia"].fillna("").astype(str).str.contains("PROMO", na=False)].copy()

    df["Referencia"] = _clean_ref(df["Referencia"])

    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date

    if "Valor neto" in df.columns:
        df["Valor neto"] = pd.to_numeric(df["Valor neto"], errors="coerce").round(0).astype("Int64")

    # Reorden útil (si existen)
    desired = ["C.O.", "Bodega", "Desc. C.O.", "Fecha", "Referencia", "Desc. item", "Talla",
               "Cantidad inv.", "Valor neto", "RANGO", "SKU"]
    front = [c for c in desired if c in df.columns]
    df = df[front + [c for c in df.columns if c not in front]]

    # === Selección al final (clave para no duplicar en Stock) ===
    if seleccion_mode == "subset" and seleccion_refs:
        refs = _clean_ref(pd.Series(seleccion_refs))
        df = df[df["Referencia"].isin(refs)].copy()
    # 'auto' ya se resolvió a 'all' si no había refs; 'all' significa no filtrar

    return df

def main(config_path: str | Path = "config.ini"):
    cfg = load_ini(config_path)

    ventas_path = cfg.get("paths", "ventas", fallback="Ventas.xlsx")
    out_dir = Path(cfg.get("paths", "out_dir", fallback=".")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    # Salida de ventas
    ventas_out_cfg = cfg.get("paths", "ventas_out", fallback="ventas_procesadas.xlsx")
    ventas_out = Path(ventas_out_cfg)
    if not ventas_out.is_absolute():
        ventas_out = out_dir / ventas_out

    sel_mode, sel_refs = load_selection_from_cfg(cfg)

    df = procesar_ventas(
        ventas_path=ventas_path,
        ventas_sheet="Sheet1",
        seleccion_mode=sel_mode,
        seleccion_refs=sel_refs,
    )

    df.to_excel(ventas_out, index=False)
    print(f"Ventas procesadas -> {ventas_out}")

if __name__ == "__main__":
    main()
