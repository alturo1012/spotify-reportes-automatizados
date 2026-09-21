"""Genera la semana BMAT: los 4 reportes, los intermedios y la lista de
tracks nuevos para revisar.

    python -m scripts.generar_bmat "C:/ruta/carpeta con los WK36-XX.xlsx"

Opciones:
    --revision ARCHIVO   una "Revisar clasificacion BMAT ..." ya corregida:
                         guarda las correcciones y recalcula la semana.
    --anio 2026          el año de la semana (por defecto se deduce de la
                         fecha de hoy; solo hace falta al cargar semanas de
                         un año anterior).
    --salida CARPETA     dónde dejar los archivos (por defecto data/output).

Ver src/bmat_proceso.py para el detalle del proceso.
"""
import argparse
import sys

from src import bmat_proceso


def main(argv=None):
    p = argparse.ArgumentParser(description="Genera los reportes BMAT de una semana.")
    p.add_argument("carpeta", help="Carpeta con los archivos WK<semana>-<mercado>.xlsx de BMAT")
    p.add_argument("--revision", help="Lista para revisar ya corregida")
    p.add_argument("--anio", type=int, help="Año de la semana (por defecto, el actual)")
    p.add_argument("--salida", help="Carpeta de salida (por defecto data/output)")
    args = p.parse_args(argv)
    try:
        r = bmat_proceso.generar_semana(args.carpeta, anio=args.anio, revision=args.revision,
                                        salida=args.salida)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    print(bmat_proceso.resumen(r))
    return r


if __name__ == "__main__":
    main()

