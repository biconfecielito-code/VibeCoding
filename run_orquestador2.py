# run_orquestador.py
# Orquesta: PreproVenta -> PreproStock -> Basecompleta
# Requisitos en la MISMA carpeta que este script / .exe:
#   - PreproVenta.py  (expone: cargar_y_transformar, exportar_xlsx)
#   - PreproStock.py  (expone: procesar_stock)
#   - Basecompleta.py (script final con argparse)
#   - Ventas.xlsx, Stock.xlsx
#   - (opc) Seleccion.xlsx (hoja con col. 'Referencias')
#   - (opc) Clasificacion_Tiendas.xlsx, Tiempos de entrega.xlsx

from __future__ import annotations

import sys
import argparse
import runpy
import subprocess
from pathlib import Path
import pandas as pd

# --- base path para modo script y modo PyInstaller ---
if getattr(sys, "frozen", False):           # ejecutable (PyInstaller)
    BASE = Path(sys.executable).parent.resolve()
else:                                       # script normal
    BASE = Path(__file__).parent.resolve()

if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

# imports locales
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
    debug: bool = False,
    no_curve_filter: bool = False,
) -> int:
    """
    Ejecuta Basecompleta.py.
    - Si estamos en PyInstaller: lo ejecuta en el MISMO intérprete con runpy.
    - Si estamos en modo script: lanza subproceso de Python.
    """
    if not base_py.is_file():
        raise FileNotFoundError(f"No se encontró {base_py}")

    argv = [
        str(base_py),
        "--ventas", str(ventas_out),
        "--ventas-sheet", ventas_sheet,
        "--stock", str(stock_out),
        "--stock-sheet", stock_sheet,
        "--out", str(final_out),
        "--strict",
    ]
    if tiendas and tiendas.is_file():
        argv += ["--tiendas", str(tiendas)]
    if tiempos and tiempos.is_file():
        argv += ["--tiempos", str(tiempos)]
    if no_curve_filter:
        argv.append("--no-curve-filter")
    if debug:
        argv.append("--debug")

    print(f"[BASECOMPLETA] Ejecutando:\n  {' '.join([sys.executable if not getattr(sys, 'frozen', False) else str(BASE / (BASE.name))] + argv)}")

    if getattr(sys, "frozen", False):
        # Ejecutar dentro del mismo intérprete (PyInstaller)
        old_argv = sys.argv[:]
        try:
            sys.argv = argv  # Basecompleta usa argparse
            runpy.run_path(str(base_py), run_name="__main__")
            rc = 0
        except SystemExit as e:
            rc = int(e.code) if hasattr(e, "code") and e.code is not None else 0
        finally:
            sys.argv = old_argv
    else:
        # Modo script: subproceso de Python
        proc = subprocess.run([sys.executable] + argv)
        rc = proc.returncode

    print(f"[BASECOMPLETA] Finalizado con código {rc}")
    return rc


def main():
    p = argparse.ArgumentParser(description="Orquestador: PreproVenta -> PreproStock -> Basecompleta")

    # Entradas de origen
    p.add_argument("--ventas", default="Ventas.xlsx", help="Excel de ventas de origen")
    p.add_argument("--ventas-sheet", default="Sheet1", help="Hoja de ventas de origen")
    p.add_argument("--stock", default="Stock.xlsx", help="Excel de stock de origen")
    p.add_argument("--stock-sheet", default="Sheet1", help="Hoja de stock de origen")

    # Selección (opcional)
    p.add_argument("--seleccion-xlsx", default="Seleccion.xlsx", help="Excel con la tabla de selección (columna 'Referencias')")
    p.add_argument("--seleccion-sheet", default="Hoja1", help="Hoja con referencias")

    # Salidas intermedias
    p.add_argument("--ventas-out", default="Ventas_procesadas_fmt.xlsx", help="Salida ventas preprocesadas (hoja 'Datos')")
    p.add_argument("--stock-out", default="Stock_procesado.xlsx", help="Salida stock preprocesado (hoja 'Sheet1')")

    # Script final y salida
    p.add_argument("--base-script", default="Basecompleta.py", help="Script final a ejecutar")
    p.add_argument("--final-out", default="Sugerencias_Traslados_Proyecto.xlsx", help="Archivo XLSX final")
    p.add_argument("--skip-base", action="store_true", help="No ejecutar Basecompleta.py")

    # Extras
    p.add_argument("--tiendas-xlsx", default="Clasificacion_Tiendas.xlsx", help="Clasificación de tiendas (opcional)")
    p.add_argument("--tiempos-xlsx", default="Tiempos de entrega.xlsx", help="Tiempos de entrega (opcional)")

    # Flags
    p.add_argument("--no-curve-filter", action="store_true", help="Desactiva filtro por curvas en Basecompleta")
    p.add_argument("--debug", action="store_true", help="Logs adicionales")

    args, _ = p.parse_known_args()

    # Rutas absolutas (no reusar nombres)
    ventas_src  = (BASE / args.ventas).resolve()
    stock_src   = (BASE / args.stock).resolve()
    ventas_out  = (BASE / args.ventas_out).resolve()
    stock_out   = (BASE / args.stock_out).resolve()
    base_script = (BASE / args.base_script).resolve()
    final_out   = (BASE / args.final_out).resolve()

    seleccion_xls = (BASE / args.seleccion_xlsx).resolve() if args.seleccion_xlsx else None
    tiendas_xls   = (BASE / args.tiendas_xlsx).resolve() if args.tiendas_xlsx else None
    tiempos_xls   = (BASE / args.tiempos_xlsx).resolve() if args.tiempos_xlsx else None

    # Validación de entradas de origen
    if not ventas_src.is_file():
        raise FileNotFoundError(f"No existe el archivo de ventas: {ventas_src}")
    if not stock_src.is_file():
        raise FileNotFoundError(f"No existe el archivo de stock: {stock_src}")

    # ==========================
    # 1) PREPROCESAMIENTO VENTAS
    # ==========================
    if seleccion_xls and seleccion_xls.is_file():
        if args.debug:
            print(f"[VENTAS] Selección desde archivo: {seleccion_xls} (hoja={args.seleccion_sheet})")
        try:
            seleccion_df = pd.read_excel(seleccion_xls, sheet_name=args.seleccion_sheet, engine="openpyxl")
        except Exception as e:
            print(f"[VENTAS] Aviso: no se pudo leer selección ({e}). Se continuará sin filtro.")
            seleccion_df = None
    else:
        seleccion_df = None
        if args.debug:
            print("[VENTAS] Sin selección explícita: se procesará todo y luego derivamos selección para stock de las ventas procesadas.")

    if args.debug:
        print(f"[VENTAS] Origen: {ventas_src.name} (hoja={args.ventas_sheet})")

    # ¡OJO! procesar SIEMPRE el archivo de origen, no el de salida
    df_ventas = cargar_y_transformar(
        ventas_path=ventas_src,
        ventas_sheet=args.ventas_sheet,
        seleccion_df=seleccion_df,       # None = no filtro
        debug=args.debug,                 # tu PreproVenta ya debe aceptar 'debug'
    )
    exportar_xlsx(df_ventas, ventas_out, add_resumen=True)

    if args.debug:
        print(f"[VENTAS] Filas finales: {len(df_ventas)} -> {ventas_out.name}")

    # =========================
    # 2) PREPROCESAMIENTO STOCK
    # =========================
    # Si no hubo selección explícita, derivarla de las ventas procesadas
    if seleccion_df is None:
        refs = pd.Series(df_ventas.get("Referencia", pd.Series([], dtype="string"))).dropna().astype(str).unique()
        seleccion_df = pd.DataFrame({"Referencias": refs})
        if args.debug:
            print(f"[STOCK] Selección derivada de ventas: {len(seleccion_df)} referencias")

    if args.debug:
        print(f"[STOCK] Origen: {stock_src.name} (hoja={args.stock_sheet})")

    df_stock = procesar_stock(
        stock_path=stock_src,
        stock_sheet=args.stock_sheet,
        venta_df=df_ventas,                   # ventas ya procesadas (evita re-proceso)
        venta_path=None, venta_sheet=None,
        seleccion_df=seleccion_df,            # selección derivada o de archivo
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
        ventas_out=ventas_out,          # PASA la salida ya procesada
        stock_out=stock_out,
        final_out=final_out,
        tiendas=tiendas_xls if (tiendas_xls and tiendas_xls.is_file()) else None,
        tiempos=tiempos_xls if (tiempos_xls and tiempos_xls.is_file()) else None,
        ventas_sheet="Datos",
        stock_sheet="Sheet1",
        debug=args.debug,
        no_curve_filter=args.no_curve_filter,   # atributo correcto (guion_bajo)
    )
    if rc != 0:
        raise SystemExit(rc)

    print("\n✅ Flujo completo OK")
    print(f"  - {ventas_out}")
    print(f"  - {stock_out}")
    print(f"  - {final_out}")


if __name__ == "__main__":
    main()
