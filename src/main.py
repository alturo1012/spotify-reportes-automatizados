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


# Cuántas semanas de diferencia se toleran entre el número que le vamos a
# asignar a esta semana y la semana del calendario que le corresponde a su
# fecha, antes de sospechar que faltan semanas en el histórico. Se deja
# holgado (3) porque la numeración del proyecto es secuencial por año, no
# ISO, y puede desfasarse un poco de forma legítima.
TOLERANCIA_SEMANAS = 3


def ultima_semana_guardada(anio: int):
    """Hasta qué semana de `anio` llega el histórico que alimenta las
    pestañas por país del Market Share (`ms_band_label_weekly`), o "ninguna"
    si ese año todavía no tiene nada.

    Sirve para que el aviso de "esta fecha ya estaba cargada" pueda decir
    hasta dónde llegan los reportes, que es justo la duda que aparece cuando
    uno genera "la semana 34" y en las pestañas ve la 33.
    """
    historico = history.cargar_ms_band_label_weekly()
    if historico.empty:
        return "ninguna"
    del_anio = historico.loc[historico["anio"] == anio, "semana"]
    return int(del_anio.max()) if not del_anio.empty else "ninguna"


def revisar_orden(fecha) -> str:
    """Avisa si esta semana es ANTERIOR a la última cargada. Devuelve el
    motivo, o None si viene en orden.

    Por qué importa: el número de semana se asigna por orden de carga (la
    última guardada + 1), no por fecha. Si una semana se salta y se carga
    después, se le asigna el número más alto y termina dibujada al FINAL de
    la cuadrícula, detrás de semanas posteriores a ella -- y de paso todas
    las que se cargaron en el medio quedan corridas un número.

    Le pasó al usuario con el 20 de agosto de 2026: esa semana no quedó en
    la base (la base se volvió a sembrar y se perdió), se siguieron cargando
    las siguientes, y la cuadrícula quedó 32, 33, 34 (27-ago), 35 (03-sep),
    sin el 20-ago y con todo corrido. Cargarla en ese momento la habría
    puesto al final, que es peor. La salida está en
    `scripts/recargar_semanas.py`.
    """
    ultima = history.ultima_fecha_cargada()
    if ultima is None:
        return None
    fecha = pd.Timestamp(fecha)
    if fecha >= ultima:
        return None
    return (
        f"esta semana ({fecha.date()}) es ANTERIOR a la última que hay en el histórico "
        f"({ultima.date()}). Se guardó igual, pero con el número más alto, así que en "
        "las pestañas por país va a aparecer al final, fuera de orden. Para dejar la "
        "numeración bien, corre `python -m scripts.recargar_semanas` con TODOS los "
        "archivos fuente posteriores a la siembra: los vuelve a cargar en orden de fecha."
    )


def revisar_historico(fecha) -> str:
    """Avisa si el histórico parece vacío o incompleto. Devuelve el motivo,
    o None si todo se ve bien.

    Existe por un problema real: al usuario se le borró la base sin querer y
    cargó una semana de agosto sobre un histórico vacío. El programa la
    numeró como "semana 1", generó los dos reportes sin quejarse, y salieron
    con una sola fila de historia y el YTD todo en ceros -- tuvo que darse
    cuenta él abriendo los archivos.

    Cómo lo detecta: compara el número que le va a tocar a esta semana
    contra la semana del calendario de su fecha. Si vamos a guardar la
    "semana 1" y la fecha es de agosto (semana ~34), falta casi todo el
    año -- señal de que la base se borró o nunca se sembró.
    """
    fecha = pd.Timestamp(fecha)
    anio = fecha.year
    semana_calendario = fecha.isocalendar()[1]

    historico = history.cargar_chart_band_weekly()
    if historico.empty:
        return (
            "el histórico está VACÍO: esta semana se va a guardar como la número 1 "
            "y los reportes van a salir con una sola fila de historia. "
            "Revisa que exista data/history/universal_data.db y, si no, corre "
            "`python -m scripts.sembrar_historico` antes de generar."
        )

    del_anio = historico.loc[historico["anio"] == anio, "semana"]
    proxima_semana = (int(del_anio.max()) + 1) if not del_anio.empty else 1

    if semana_calendario - proxima_semana >= TOLERANCIA_SEMANAS:
        return (
            f"esta semana se va a guardar como la número {proxima_semana} de {anio}, "
            f"pero su fecha ({fecha.date()}) corresponde a la semana {semana_calendario} "
            "del calendario. Parece que al histórico le faltan semanas: los reportes "
            "van a salir incompletos. Si borraste la base, corre "
            "`python -m scripts.sembrar_historico` y vuelve a cargar esta semana."
        )
    return None


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

    avisos = []

    ya_cargada = history.semana_ya_cargada(fecha)
    if ya_cargada:
        aviso_repetida = (
            f"la fecha {pd.Timestamp(fecha).date()} de este archivo YA estaba en el "
            "histórico, así que no se agregó otra vez (para no duplicar la semana). "
            f"Los reportes salieron con el histórico que ya había, hasta la semana "
            f"{ultima_semana_guardada(pd.Timestamp(fecha).year)}. Si esperabas ver una "
            "semana nueva, revisa el archivo fuente: lo que manda es la fecha que trae "
            "adentro (chart_date), no el número de semana que escribiste ni el nombre "
            "del archivo."
        )
        avisos.append(aviso_repetida)
        print(f"Aviso: {aviso_repetida}")
    guardar_en_historico = not ya_cargada

    if guardar_en_historico:
        for revision in (revisar_historico, revisar_orden):
            aviso = revision(fecha)
            if aviso:
                avisos.append(aviso)
                print(f"Aviso: {aviso}")

    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    chart_out = config.OUTPUT_DIR / f"Reporte_Chart_Top_Semanal_Sem_{args.semana}.xlsx"
    chart_semanal.generar_reporte(df, chart_out, guardar_en_historico=guardar_en_historico)
    print(f"Reporte de chart semanal generado: {chart_out}")

    ms_out = config.OUTPUT_DIR / f"Reporte_MS_TOP200_Sem_{args.semana}.xlsx"
    market_share.generar_reporte(df, ms_out, guardar_en_historico=guardar_en_historico)
    print(f"Reporte de market share generado: {ms_out}")

    return avisos


if __name__ == "__main__":
    main()
