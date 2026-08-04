"""Orquesta la generación de ambos reportes a partir de la fuente BQ.

Uso:
    python -m src.main --fuente data/raw/Fuente_de_datos_BQ_Spotify.xlsx --semana 25
"""
import argparse
from pathlib import Path

from . import config, load_data, chart_semanal, market_share


def main():
    parser = argparse.ArgumentParser(description="Genera reportes Spotify Latam")
    parser.add_argument("--fuente", required=True, help="Ruta al xlsx fuente de BigQuery")
    parser.add_argument("--semana", required=True, help="Número de semana (para el nombre de salida)")
    args = parser.parse_args()

    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    df = load_data.load_source(Path(args.fuente))

    chart_out = config.OUTPUT_DIR / f"Reporte_Chart_Top_Semanal_Sem_{args.semana}.xlsx"
    chart_semanal.generar_reporte(df, chart_out)
    print(f"Reporte de chart semanal generado: {chart_out}")

    ms_out = config.OUTPUT_DIR / f"Reporte_MS_TOP200_Sem_{args.semana}.xlsx"
    market_share.generar_reporte(df, ms_out)
    print(f"Reporte de market share generado: {ms_out}")


if __name__ == "__main__":
    main()
