# run_orquestador.py
# Orquesta: PreproVenta -> PreproStock -> Basecompleta (archivos separados)
# Requisitos en la MISMA carpeta:
#   - PreproVenta.py     (expone: cargar_y_transformar, exportar_xlsx)
#   - PreproStock.py     (expone: procesar_stock)
#   - Basecompleta.py    (acepta: --ventas --ventas-sheet --stock --stock-sheet --out [--tiendas] [--tiempos] [--strict] [--no-curve-filter] [--debug])
#   - Ventas.xlsx, Stock.xlsx
#   - (opcional) Seleccion.xlsx (hoja "Seleccion" con columna "Referencias")
#   - (opcional) Clasificacion_Tiendas.xlsx, Tiempos de entrega.xlsx

from __future__ import annotations

import sys
import subprocess
import argparse
from pathlib import Path
import pandas as pd

# --- asegurar imports locales (misma carpeta) ---
"""BASE = Path(__file__).parent.resolve()
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))"""

if getattr(sys, "frozen", False):  # ejecutable (PyInstaller)
    BASE = Path(sys.executable).parent.resolve()
else:                              # script normal
    BASE = Path(__file__).parent.resolve()

if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

# importa funciones de tus preprocesos
from PreproVenta import cargar_y_transformar, exportar_xlsx
from PreproStock import procesar_stock


def run_basecompleta(
    base_py: Path,
    ventas_out: Path,
    stock_out: Path,
    final_out: Path,
    tiendas: Path | None = None,
    tiempos: Path | None = None,
    ventas_sheet: str = "Datos",
    stock_sheet: str = "Sheet1",
    python_exe: str | None = None,
    debug: bool = False,
    no_curve_filter: bool = False,
) -> int:
    """
    Ejecuta Basecompleta.py como subproceso con rutas explícitas.
    Retorna el código de salida del proceso.
    """
    if not base_py.is_file():
        raise FileNotFoundError(f"No se encontró {base_py}")

    cmd = [
        python_exe or sys.executable, str(base_py),
        "--ventas", str(ventas_out),
        "--ventas-sheet", ventas_sheet,
        "--stock", str(stock_out),
        "--stock-sheet", stock_sheet,
        "--out", str(final_out),
        "--strict",  # fuerza error si se queda sin datos útiles
    ]
    if tiendas and tiendas.is_file():
        cmd += ["--tiendas", str(tiendas)]
    if tiempos and tiempos.is_file():
        cmd += ["--tiempos", str(tiempos)]
    if no_curve_filter:
        cmd.append("--no-curve-filter")
    if debug:
        cmd.append("--debug")

    print(f"[BASECOMPLETA] Ejecutando:\n  {' '.join(cmd)}")
    proc = subprocess.run(cmd)
    print(f"[BASECOMPLETA] Finalizado con código {proc.returncode}")
    return proc.returncode


def main():
    p = argparse.ArgumentParser(
        description="Orquestador: PreproVenta -> PreproStock -> Basecompleta (archivos separados)"
    )

    # Entradas de origen
    p.add_argument("--ventas", default="Ventas.xlsx", help="Excel de ventas de origen (default: Ventas.xlsx)")
    p.add_argument("--ventas-sheet", default="Sheet1", help="Hoja de ventas de origen (default: Sheet1)")
    p.add_argument("--stock", default="Stock.xlsx", help="Excel de stock de origen (default: Stock.xlsx)")
    p.add_argument("--stock-sheet", default="Sheet1", help="Hoja de stock de origen (default: Sheet1)")

    # Selección (opcional). Si no se pasa, se derivará de ventas procesadas.
    p.add_argument("--seleccion-xlsx", default="Seleccion.xlsx",
                   help="Excel con la tabla 'Seleccion' (columna 'Referencias'). (opcional)")
    p.add_argument("--seleccion-sheet", default="Hoja1",
                   help="Hoja con referencias (default: Hoja1)")

    # Salidas intermedias
    p.add_argument("--ventas-out", default="Ventas_procesadas_fmt.xlsx",
                   help="Salida de ventas preprocesadas (hoja 'Datos')")
    p.add_argument("--stock-out", default="Stock_procesado.xlsx",
                   help="Salida de stock preprocesado (hoja 'Sheet1')")

    # Tercer script y salida final
    p.add_argument("--base-script", default="Basecompleta.py",
                   help="Script final a ejecutar (default: Basecompleta.py)")
    p.add_argument("--final-out", default="Sugerencias_Traslados_Proyecto.xlsx",
                   help="Archivo XLSX final (default: Sugerencias_Traslados_Proyecto.xlsx)")
    p.add_argument("--skip-base", action="store_true",
                   help="No ejecutar Basecompleta.py (solo generar ventas/stock)")

    # Extras (opcionales)
    p.add_argument("--tiendas-xlsx", default="Clasificacion_Tiendas.xlsx",
                   help="Clasificación de tiendas (opcional)")
    p.add_argument("--tiempos-xlsx", default="Tiempos de entrega.xlsx",
                   help="Tiempos de entrega (opcional)")

    # Flags
    p.add_argument("--no-curve-filter", action="store_true",
                   help="Pasa --no-curve-filter a Basecompleta (desactiva filtro por curvas)")
    p.add_argument("--debug", action="store_true", help="Logs adicionales")

    args, _ = p.parse_known_args()

    # Rutas absolutas
    ventas_xlsx   = (BASE / args.ventas).resolve()
    stock_xlsx    = (BASE / args.stock).resolve()
    ventas_out    = (BASE / args.ventas_out).resolve()
    stock_out     = (BASE / args.stock_out).resolve()
    base_script   = (BASE / args.base_script).resolve()
    final_out     = (BASE / args.final_out).resolve()

    seleccion_xls = (BASE / args.seleccion_xlsx).resolve() if args.seleccion_xlsx else None
    tiendas_xls   = (BASE / args.tiendas_xlsx).resolve() if args.tiendas_xlsx else None
    tiempos_xls   = (BASE / args.tiempos_xlsx).resolve() if args.tiempos_xlsx else None

    # Validaciones mínimas de las entradas de origen
    if not ventas_xlsx.is_file():
        raise FileNotFoundError(f"No existe el archivo de ventas: {ventas_xlsx}")
    if not stock_xlsx.is_file():
        raise FileNotFoundError(f"No existe el archivo de stock: {stock_xlsx}")

    # ==========================
    # 1) PREPROCESAMIENTO VENTAS
    # ==========================
    if seleccion_xls and seleccion_xls.is_file():
        if args.debug:
            print(f"[VENTAS] Selección desde archivo: {seleccion_xls} (hoja={args.seleccion_sheet})")
        seleccion_df = pd.read_excel(seleccion_xls, sheet_name=args.seleccion_sheet, engine="openpyxl")
    else:
        seleccion_df = None
        if args.debug:
            print("[VENTAS] Sin selección explícita: se procesará todo y luego derivamos selección para stock de las ventas procesadas.")

    if args.debug:
        print(f"[VENTAS] Origen: {ventas_xlsx.name} (hoja={args.ventas_sheet})")

    df_ventas = cargar_y_transformar(
        ventas_path=ventas_xlsx,
        ventas_sheet=args.ventas_sheet,
        seleccion_df=seleccion_df,  # None = no filtro
        debug=args.debug,
    )
    exportar_xlsx(df_ventas, ventas_out, add_resumen=True)

    if args.debug:
        print(f"[VENTAS] Filas finales: {len(df_ventas)} -> {ventas_out.name}")

    # =========================
    # 2) PREPROCESAMIENTO STOCK
    # =========================
    # Si no hubo selección explícita, la derivamos de Ventas procesadas
    if seleccion_df is None:
        refs = pd.Series(df_ventas.get("Referencia", pd.Series([], dtype="string"))).dropna().astype(str).unique()
        seleccion_df = pd.DataFrame({"Referencias": refs})
        if args.debug:
            print(f"[STOCK] Selección derivada de ventas: {len(seleccion_df)} referencias")

    if args.debug:
        print(f"[STOCK] Origen: {stock_xlsx.name} (hoja={args.stock_sheet})")

    df_stock = procesar_stock(
        stock_path=stock_xlsx,
        stock_sheet=args.stock_sheet,
        venta_df=df_ventas,                   # usar ventas ya procesadas
        venta_path=None, venta_sheet=None,    # (no se requiere path)
        seleccion_df=seleccion_df,            # selección (archivo o derivada)
        seleccion_path=None, seleccion_sheet=None,
    )
    df_stock.to_excel(stock_out, index=False)

    if args.debug:
        print(f"[STOCK] Filas finales: {len(df_stock)} -> {stock_out.name}")

    # ===========================
    # 3) SCRIPT FINAL (opcional)
    # ===========================
    if args.skip_base:
        print("\n✅ Flujo OK (saltado Basecompleta.py por --skip-base)")
        print(f"  - {ventas_out}")
        print(f"  - {stock_out}")
        return

    rc = run_basecompleta(
        base_py=base_script,
        ventas_out=ventas_out,
        stock_out=stock_out,
        final_out=final_out,
        tiendas=tiendas_xls if (tiendas_xls and tiendas_xls.is_file()) else None,
        tiempos=tiempos_xls if (tiempos_xls and tiempos_xls.is_file()) else None,
        ventas_sheet="Datos",
        stock_sheet="Sheet1",
        debug=args.debug,
        no_curve_filter=args.no-curve-filter if hasattr(args, "no-curve-filter") else args.no_curve_filter,  # por si el guion falla en algunas shells
    )
    if rc != 0:
        raise SystemExit(rc)

    print("\n✅ Flujo completo OK")
    print(f"  - {ventas_out}")
    print(f"  - {stock_out}")
    print(f"  - {final_out}")


if __name__ == "__main__":
    main()
