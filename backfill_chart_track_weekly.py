"""Script de un solo uso: rellena `chart_track_weekly` para UNA semana que ya
estaba guardada en `ms_label_weekly`/`chart_band_weekly` antes de que
existiera `chart_track_weekly` (Ajuste 5), a partir del mismo archivo fuente
de esa semana. NO vuelve a tocar `chart_band_weekly`/`ms_label_weekly` --
esos ya están bien, este script solo llena la tabla nueva que faltó.

Uso (desde la raíz del repo, con el venv activado):
    python backfill_chart_track_weekly.py --fuente "ruta al xlsx de esa semana"

Se puede correr una vez por cada semana vieja que todavía tengas el archivo
fuente a mano (por ejemplo, si probaste varios ajustes con la misma semana
antes de que existiera chart_track_weekly). Si la fecha de esa fuente ya
está guardada en chart_track_weekly, avisa y no hace nada -- seguro de
correr más de una vez con el mismo archivo por error.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd

from src import history, load_data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fuente", required=True, help="Ruta al xlsx fuente de esa semana")
    args = parser.parse_args()

    fuente_path = Path(args.fuente)
    if not fuente_path.exists():
        sys.exit(f"No se encontró el archivo fuente: {fuente_path}")

    df = load_data.load_source(fuente_path)
    fechas = df["chart_date"].unique()
    if len(fechas) != 1:
        sys.exit(f"La fuente trae {len(fechas)} fechas distintas, se esperaba una sola.")
    fecha = pd.Timestamp(fechas[0])
    fecha_str = fecha.date().isoformat()

    ya_guardado = history.cargar_chart_track_weekly()
    if not ya_guardado.empty and fecha_str in set(ya_guardado["chart_date"]):
        print(f"La fecha {fecha_str} ya está en chart_track_weekly -- no se hace nada (para no duplicar).")
        return

    history.append_semana_tracks(df)
    print(f"Listo: se guardó el detalle track por track de la semana {fecha_str} en chart_track_weekly.")


if __name__ == "__main__":
    main()
