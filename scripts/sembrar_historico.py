"""Script de un solo uso: siembra el histórico real en la base local
`universal_data.db`, a partir de los CSV de `data/history/seed/`.

    python -m scripts.sembrar_historico

Siembra las tres tablas de histórico:
  - chart_band_weekly    (Chart Semanal, 2019 -> 2026 semana 33)
  - ms_label_weekly      (streams por sello, 2025 -> 2026 semana 24)
  - ms_band_label_weekly (% por banda/sello de cada país, 2021 -> 2026 sem 33)

Es seguro correrlo más de una vez: `seed_historico()` usa INSERT OR IGNORE y
no duplica lo que ya estaba cargado. OJO: por lo mismo, si un CSV de siembra
se ACTUALIZA con semanas nuevas, correr el script las agrega sin problema;
pero si se corrigiera un valor ya cargado, hay que borrar
`data/history/universal_data.db` y volver a sembrar para que tome el nuevo.
"""
from src import history

if __name__ == "__main__":
    history.seed_historico()

    tablas = [
        ("chart_band_weekly", history.cargar_chart_band_weekly),
        ("ms_label_weekly", history.cargar_ms_label_weekly),
        ("ms_band_label_weekly", history.cargar_ms_band_label_weekly),
    ]
    for nombre, cargar in tablas:
        df = cargar()
        if df.empty:
            print(f"{nombre}: vacía (¿falta su CSV en data/history/seed/?)")
            continue
        ultimo_anio = int(df["anio"].max())
        ultima_semana = int(df.loc[df["anio"] == ultimo_anio, "semana"].max())
        print(
            f"{nombre}: {len(df)} filas "
            f"({int(df['anio'].min())}-{ultimo_anio}, hasta la semana {ultima_semana})"
        )