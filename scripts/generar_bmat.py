"""Genera el reporte MS BMAT (Promúsica Colombia) con el histórico guardado.

    python -m scripts.generar_bmat

A diferencia de los dos reportes de Spotify, este NO recibe una fuente de
datos: todavía no está definida la lógica para calcular una semana nueva a
partir de un archivo de BMAT, así que por ahora el reporte es un reflejo del
histórico ya sembrado (ver src/bmat.py). Cuando se defina, este script
recibirá la fuente igual que `src/main.py`.
"""
from src import bmat, config

if __name__ == "__main__":
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    salida = config.OUTPUT_DIR / "Reporte_MS_BMAT_COL.xlsx"

    historico = bmat.cargar_historico()
    if historico.empty:
        print(
            "Aviso: no hay histórico de BMAT guardado todavía "
            "(¿falta correr `python -m scripts.sembrar_historico`?). "
            "El reporte se genera igual, solo con la estructura."
        )
    else:
        semanas = historico[["anio", "semana"]].drop_duplicates()
        ultima = semanas.sort_values(["anio", "semana"]).iloc[-1]
        print(
            f"Histórico BMAT: {len(semanas)} semanas, "
            f"hasta la semana {int(ultima.semana)} de {int(ultima.anio)}."
        )

    bmat.generar_reporte(salida)
    print(f"Reporte generado: {salida}")
