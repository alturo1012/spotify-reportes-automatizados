"""Script de un solo uso: siembra el histórico real (2019-2026 para Chart
Semanal, 2025-2026 para Market Share) en la base local `universal_data.db`.

Correlo UNA VEZ después de agregar `src/history.py` y los CSV de
`data/history/seed/` al repo:

    python -m scripts.sembrar_historico

Es seguro correrlo más de una vez — `seed_historico()` no duplica filas si
ya estaban cargadas.
"""
from src import history

if __name__ == "__main__":
    history.seed_historico()
    chart_df = history.cargar_chart_band_weekly()
    ms_df = history.cargar_ms_label_weekly()
    print(f"chart_band_weekly: {len(chart_df)} filas ({chart_df['anio'].min()}-{chart_df['anio'].max()})")
    print(f"ms_label_weekly: {len(ms_df)} filas ({ms_df['anio'].min()}-{ms_df['anio'].max()})")
