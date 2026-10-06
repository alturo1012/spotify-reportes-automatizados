"""Rellena los tracks y los streams de semanas que ya están en el histórico.

    python -m scripts.completar_semanas "C:\\ruta\\BQ_semana_32.xlsx" "...semana_33.xlsx"
    python -m scripts.completar_semanas "C:\\ruta\\*.xlsx"

PARA QUÉ SIRVE
--------------
El histórico sembrado (hasta la semana 33 de 2026) se sacó de los reportes
reales, que solo traían el % de market share: esas semanas no tienen los
streams ni el conteo de tracks, y el reporte MENSUAL los necesita para poder
cerrar el mes. Agosto de 2026, por ejemplo, necesita las semanas 32 y 33.

Cuando aparece la fuente de BQ de una de esas semanas, este script le mete
los números que le faltaban CONSERVANDO su número de semana.

POR QUÉ NO SIRVE recargar_semanas.py
------------------------------------
Ese salta a propósito los archivos cuya fecha ya viene en la siembra, para
no numerar la misma semana dos veces. Si le pasas la semana 33 te dice "ya
venía en la siembra, se salta" y el mes sigue sin poder cerrarse.

QUÉ NO TOCA
-----------
Solo reescribe las filas de esa semana en `ms_band_label_weekly`. No cambia
la numeración, no vacía nada, no toca las otras tablas ni el histórico
mensual. Correrlo dos veces con el mismo archivo da el mismo resultado.
"""
import argparse
import glob
import sys
from pathlib import Path

from src import history, load_data


def _expandir(patrones) -> list:
    rutas = []
    for patron in patrones:
        encontrados = sorted(glob.glob(patron))
        if encontrados:
            rutas.extend(encontrados)
        elif Path(patron).exists():
            rutas.append(patron)
        else:
            print(f"  (no existe, se ignora: {patron})")
    return list(dict.fromkeys(rutas))


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Rellena tracks y streams de semanas ya sembradas, sin renumerarlas")
    parser.add_argument("fuentes", nargs="+", help="Archivos fuente de BQ (admite comodines)")
    args = parser.parse_args(argv)

    rutas = _expandir(args.fuentes)
    if not rutas:
        sys.exit("No se encontró ningún archivo con lo que pasaste.")

    completadas = 0
    for ruta in rutas:
        nombre = Path(ruta).name
        try:
            df = load_data.load_source(Path(ruta))
        except Exception as e:  # noqa: BLE001 -- el motivo se le muestra al usuario
            print(f"  NO se pudo leer {nombre}: {e}")
            continue
        resultado = history.completar_semana_ms_bandas(df)
        if resultado is None:
            print(f"  {nombre}: su fecha NO está en el histórico. Esta semana hay que cargarla "
                  "con el flujo normal (el .exe o src.main), no con este script.")
            continue
        anio, semana, filas = resultado
        print(f"  {nombre} -> semana {semana} de {anio}: {filas} valores recalculados.")
        completadas += 1

    print(f"\nListo: {completadas} semana(s) completada(s).")
    if completadas:
        print("Ahora ya se puede generar el reporte mensual de esos meses.")
    return completadas


if __name__ == "__main__":
    main()
