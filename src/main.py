"""Orquesta la generación de ambos reportes a partir de la fuente BQ.

Uso:
    python -m src.main --fuente data/raw/Fuente_de_datos_BQ_Spotify.xlsx --semana 25

1. Carga y valida la fuente BQ de la semana (`load_data.load_source`).
2. Si la fecha de esa semana todavía no estaba en el histórico, la agrega a
   las dos tablas acumuladas (`chart_band_weekly` y `ms_label_weekly`, vía
   `history.py`). Si ya estaba -- por ejemplo si el script se corrió dos
   veces por error con el mismo archivo, o se volvió a correr después de
   una falla a la mitad -- NO se vuelve a agregar, para no duplicar la
   semana ni generar un número de semana fantasma; en ese caso solo se
   regeneran los reportes con el histórico que ya había.
3. Genera los dos reportes de salida (Chart Semanal y Market Share), cada
   uno recalculado con TODO el histórico acumulado hasta esa semana.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

from . import chart_semanal, config, history, load_data, market_share


def main(argv=None):
    parser = argparse.ArgumentParser(description="Genera los reportes semanales de Spotify Latam")
    parser.add_argument(
        "--fuente", required=True,
        help="Ruta al xlsx fuente de BigQuery de la semana (la 'Fuente de datos_BQ Spotify...')",
    )
    parser.add_argument(
        "--semana", required=True,
        help="Número de semana, solo para el nombre de los archivos de salida (ej. 25)",
    )
    args = parser.parse_args(argv)

    fuente_path = Path(args.fuente)
    if not fuente_path.exists():
        sys.exit(f"No se encontró el archivo fuente: {fuente_path}")

    print(f"Cargando fuente: {fuente_path}")
    df = load_data.load_source(fuente_path)

    fechas = df["chart_date"].unique()
    if len(fechas) != 1:
        sys.exit(
            f"La fuente trae {len(fechas)} fechas distintas en chart_date "
            f"({list(fechas)}); se esperaba una sola semana por archivo."
        )
    fecha = fechas[0]

    ya_cargada = history.semana_ya_cargada(fecha)
    if ya_cargada:
        print(
            f"Aviso: la fecha {pd.Timestamp(fecha).date()} ya estaba guardada en el "
            "histórico -- no se vuelve a agregar (para no duplicar la semana). "
            "Se regeneran los reportes igual, con el histórico que ya había."
        )
    guardar_en_historico = not ya_cargada

    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    chart_out = config.OUTPUT_DIR / f"Reporte_Chart_Top_Semanal_Sem_{args.semana}.xlsx"
    chart_semanal.generar_reporte(df, chart_out, guardar_en_historico=guardar_en_historico)
    print(f"Reporte de chart semanal generado: {chart_out}")

    ms_out = config.OUTPUT_DIR / f"Reporte_MS_TOP200_Sem_{args.semana}.xlsx"
    market_share.generar_reporte(df, ms_out, guardar_en_historico=guardar_en_historico)
    print(f"Reporte de market share generado: {ms_out}")


if __name__ == "__main__":
    main()
