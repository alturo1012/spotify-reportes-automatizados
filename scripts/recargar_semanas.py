"""Reconstruye el histórico cargando las semanas EN ORDEN DE FECHA.

    python -m scripts.recargar_semanas "ruta\\a\\fuentes\\*.xlsx"
    python -m scripts.recargar_semanas archivo1.xlsx archivo2.xlsx ...

PARA QUÉ SIRVE
--------------
El número de semana se asigna por orden de carga (la última guardada + 1),
no por la fecha del archivo. Mientras las semanas se carguen una tras otra
eso funciona; pero si una se salta -- porque falló, porque se volvió a
sembrar la base y se perdió, o porque simplemente se cargó fuera de orden --
todas las siguientes quedan corridas un número, y la que falta, si se carga
después, aparece al final de la cuadrícula en vez de en su lugar.

Este script arregla eso de la única forma que no deja dudas: vacía las cinco
tablas de histórico, vuelve a sembrar el histórico oficial y carga los
archivos que le pases ordenados por su chart_date.

QUÉ NO TOCA
-----------
- Los cachés de Spotify (las fechas de lanzamiento ya resueltas) se quedan
  donde están: no hay que volver a pedirlas a la API.
- Los archivos fuente: solo se leen.
- Los reportes ya generados en data/output/: no se borran ni se regeneran.
  Después de correr esto, vuelve a generar el reporte de la última semana.

QUÉ ARCHIVOS PASARLE
--------------------
TODOS los archivos fuente de las semanas posteriores a la siembra (la
siembra ya trae la historia hasta la semana 33 de 2026). Si le pasas uno
cuya fecha ya viene en la siembra, lo salta y lo dice: esa semana ya está.
"""
import argparse
import glob
import sys
from pathlib import Path

import pandas as pd

from src import chart_semanal, history, load_data, market_share


def _expandir(patrones) -> list:
    """Acepta rutas sueltas y comodines (en Windows, cmd no los expande)."""
    rutas = []
    for patron in patrones:
        encontrados = sorted(glob.glob(patron))
        if encontrados:
            rutas.extend(encontrados)
        elif Path(patron).exists():
            rutas.append(patron)
        else:
            print(f"  (no existe, se ignora: {patron})")
    # Sin repetidos, conservando el orden
    return list(dict.fromkeys(rutas))


def _leer_fuentes(rutas) -> list:
    """[(fecha, ruta, df), ...] ordenado por fecha. Descarta lo que no se
    pueda leer, diciendo por qué -- mejor saltar un archivo raro que abortar
    la reconstrucción entera a la mitad."""
    fuentes = []
    for ruta in rutas:
        try:
            df = load_data.load_source(Path(ruta))
        except Exception as e:  # noqa: BLE001 -- el motivo se le muestra al usuario
            print(f"  NO se pudo leer {Path(ruta).name}: {e}")
            continue
        fechas = df["chart_date"].unique()
        if len(fechas) != 1:
            print(f"  NO se usa {Path(ruta).name}: trae {len(fechas)} fechas distintas.")
            continue
        fuentes.append((pd.Timestamp(fechas[0]), ruta, df))
    return sorted(fuentes, key=lambda f: f[0])


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Vacía el histórico, lo vuelve a sembrar y carga las semanas en orden de fecha"
    )
    parser.add_argument("fuentes", nargs="+", help="Archivos fuente (admite comodines, ej. carpeta/*.xlsx)")
    args = parser.parse_args(argv)

    print("1. Buscando archivos...")
    rutas = _expandir(args.fuentes)
    if not rutas:
        sys.exit("No se encontró ningún archivo fuente con lo que pasaste.")
    print(f"   {len(rutas)} archivo(s).")

    print("2. Leyendo y ordenando por fecha...")
    fuentes = _leer_fuentes(rutas)
    if not fuentes:
        sys.exit("Ninguno de los archivos se pudo leer como fuente de una semana.")
    for fecha, ruta, _ in fuentes:
        print(f"   {fecha.date()}  {Path(ruta).name}")

    print("3. Vaciando el histórico y volviendo a sembrarlo...")
    history.vaciar_historico()
    history.seed_historico()
    ultima_sembrada = history.ultima_fecha_cargada()
    print(f"   Sembrado hasta {ultima_sembrada.date() if ultima_sembrada is not None else 'nada'}.")

    print("4. Cargando las semanas en orden...")
    cargadas = 0
    for fecha, ruta, df in fuentes:
        if history.semana_ya_cargada(fecha):
            print(f"   {fecha.date()} ya venía en la siembra -- se salta.")
            continue
        history.append_semana_chart(df)
        history.append_semana_tracks(df)
        history.append_semana_ms(df)
        history.append_semana_ms_bandas(df)
        anio, semana, _ = history.ultima_semana_cargada()
        print(f"   {fecha.date()} -> semana {semana} de {anio}")
        cargadas += 1

    ultima = history.ultima_semana_cargada()
    print(f"\nListo: {cargadas} semana(s) cargada(s).")
    if ultima is not None:
        anio, semana, chart_date = ultima
        print(f"El histórico llega ahora hasta la semana {semana} de {anio} ({chart_date}).")
    print("Ahora vuelve a generar el reporte de la última semana con el .exe.")


if __name__ == "__main__":
    main()
