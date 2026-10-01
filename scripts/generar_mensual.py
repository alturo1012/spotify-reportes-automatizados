"""Genera el reporte mensual "Market Share Spotify Latam".

    python -m scripts.generar_mensual --anio 2026 --mes 9
    python -m scripts.generar_mensual --anio 2026 --mes 9 --salida "D:/Reportes"

Sin --anio/--mes toma el último mes que ya terminó.

No hay que cargar nada aparte: el mes se arma con las semanas de BQ que ya
están en el histórico (las mismas que generan los reportes semanales). Si a
ese mes le falta alguna semana, el script lo avisa y la genera igual, para
que se vea qué falta.
"""
import argparse

from src import mensual


def main(argv=None):
    parser = argparse.ArgumentParser(description="Genera el reporte mensual de Market Share")
    parser.add_argument("--anio", type=int, help="Año del mes a cerrar (ej. 2026)")
    parser.add_argument("--mes", type=int, help="Número de mes a cerrar (1-12)")
    parser.add_argument("--salida", help="Carpeta donde dejar el reporte (por defecto data/output)")
    args = parser.parse_args(argv)

    anio, mes = mensual.ultimo_mes_cerrable()
    anio = args.anio or anio
    mes = args.mes or mes

    ruta, avisos = mensual.generar(anio, mes, salida=args.salida)
    print(mensual.resumen(ruta, avisos, anio, mes))
    return avisos


if __name__ == "__main__":
    main()
