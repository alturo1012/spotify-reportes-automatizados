"""Reporte MENSUAL "Market Share Spotify Latam".

Es el hermano mensual del Reporte_MS_TOP200: el mismo % de streams por sello
y banda, pero con los meses como columnas (más los cierres de Q1..Q4, H1/H2
y año) y con una banda más, el Top 20.

CÓMO SE ARMA UN MES
-------------------
El mes NO se calcula aparte: sale de las mismas semanas de BQ que ya se
cargan todos los viernes. Por cada país, banda y sello:

- **streams**: SUMA de los streams de las semanas del mes;
- **tracks**: PROMEDIO de los tracks de las semanas del mes (por eso en el
  reporte aparecen valores como 3,25 = promedio de 4 semanas);
- **%**: se RECALCULA sobre el total del mes (streams del sello / streams de
  los 7 sellos en esa banda). NO es el promedio de los % semanales.

Qué semana cae en qué mes: la de su fecha de corte (`chart_date`). Una
semana que va del 28 de agosto al 3 de septiembre cuenta para septiembre.

Las tres reglas se comprobaron contra las hojas "XX-Det" de la plantilla
real (ver claude/mensual_market_share_analisis.md):

- streams: 1.190 valores de meses completos de 2025 y 2026 (17 países, 7
  sellos), 1.190 exactos;
- tracks: 800 de 800 en 2025 y 812 de 816 en 2024;
- la regla de la fecha de corte da 89,8% de coincidencia contra 18,4% si la
  semana se asignara al mes en que empieza.

HISTORIA
--------
De mayo de 2017 hasta el último mes que traía la plantilla, el histórico
mensual viene sembrado (`history.SEED_MS_MENSUAL_CSV`, extraído de las hojas
Det). De ahí en adelante lo cierra este módulo mes a mes.
"""
from datetime import date
from pathlib import Path

import pandas as pd

from . import config, history, mensual_reporte


ORIGEN_CALCULADO = "calculado"


def _fecha(valor):
    return pd.Timestamp(valor).date()


def semanas_del_mes(anio: int, mes: int, bandas: pd.DataFrame = None) -> pd.DataFrame:
    """Las filas del histórico semanal por banda/sello cuya fecha de corte
    cae en ese mes."""
    bandas = history.cargar_ms_band_label_weekly() if bandas is None else bandas
    if bandas.empty:
        return bandas
    fechas = pd.to_datetime(bandas["chart_date"], errors="coerce")
    return bandas[(fechas.dt.year == anio) & (fechas.dt.month == mes)]


def semanas_esperadas(anio: int, mes: int, una_fecha) -> int:
    """Cuántas fechas de corte debería tener ese mes, deducido de una fecha
    de corte conocida: las fechas van de 7 en 7 días, así que basta contar
    cuántos días del mes caen en el mismo día de la semana.

    Sirve para avisar "este mes tiene 4 semanas cargadas pero le
    corresponden 5" sin tener que saber de antemano el calendario de Spotify.
    """
    referencia = _fecha(una_fecha)
    primero = date(anio, mes, 1)
    siguiente = date(anio + (mes == 12), (mes % 12) + 1, 1)
    dias = (siguiente - primero).days
    desfase = (primero.toordinal() - referencia.toordinal()) % 7
    primer_corte = 1 + ((7 - desfase) % 7)
    return len(range(primer_corte, dias + 1, 7))


def calcular_mes(anio: int, mes: int, bandas: pd.DataFrame = None) -> tuple:
    """Devuelve `(DataFrame del mes, avisos)`.

    El DataFrame trae una fila por país/banda/sello con tracks, streams y %.
    Si no hay ninguna semana de ese mes en el histórico, viene vacío.
    """
    delmes = semanas_del_mes(anio, mes, bandas)
    avisos = []
    columnas = ["anio", "mes", "country_code", "banda", "label_group",
                "tracks", "streams_millones", "pct_streams", "semanas"]
    if delmes.empty:
        avisos.append(
            f"No hay ninguna semana de {config.MESES_ES[mes].lower()} de {anio} en el "
            "histórico: hay que cargar sus fuentes de BQ antes de cerrar el mes."
        )
        return pd.DataFrame(columns=columnas), avisos

    todas = sorted(set(delmes["chart_date"]))
    esperadas = semanas_esperadas(anio, mes, todas[0])

    # Las semanas sembradas de la plantilla vieja solo traen el %: no sirven
    # para sumar streams ni promediar tracks. Se dejan fuera del cálculo y se
    # avisa, en vez de sumar como si valieran cero.
    con_datos = delmes[delmes["streams_millones"].notna()]
    fechas = sorted(set(con_datos["chart_date"]))
    sin_datos = [f for f in todas if f not in set(fechas)]
    if sin_datos:
        avisos.append(
            f"Estas semanas de {config.MESES_ES[mes].lower()} de {anio} están en el "
            f"histórico pero sin streams ni tracks ({', '.join(sin_datos)}): vienen del "
            "histórico sembrado, que solo traía el %. Consigue sus fuentes de BQ y córrelas "
            "con `python -m scripts.completar_semanas <archivo.xlsx>` para que el mes quede "
            "completo."
        )
    if not fechas:
        return pd.DataFrame(columns=columnas), avisos
    if len(fechas) < esperadas:
        avisos.append(
            f"{config.MESES_ES[mes].lower().capitalize()} de {anio} se calculó con "
            f"{len(fechas)} semana(s) ({', '.join(fechas)}) y le corresponden {esperadas}: "
            f"falta(n) {esperadas - len(fechas)}. Los streams del mes van a salir bajos "
            "hasta que se carguen las que faltan (los tracks y el % no, que son promedio "
            "y proporción)."
        )

    agrupado = con_datos.groupby(["country_code", "banda", "label_group"], as_index=False).agg(
        tracks=("tracks", "mean"),
        streams_millones=("streams_millones", "sum"),
    )
    totales = agrupado.groupby(["country_code", "banda"])["streams_millones"].transform("sum")
    agrupado["pct_streams"] = (agrupado["streams_millones"] / totales).where(totales > 0, 0.0)
    agrupado["anio"] = anio
    agrupado["mes"] = mes
    agrupado["semanas"] = len(fechas)
    return agrupado[columnas].sort_values(
        ["country_code", "banda", "label_group"]).reset_index(drop=True), avisos


MENSAJE_INCOMPLETO = "se calculó con"


def cerrar_mes(anio: int, mes: int, bandas: pd.DataFrame = None) -> tuple:
    """Calcula el mes y, si está completo, lo guarda en el histórico mensual.
    Devuelve `(DataFrame, avisos, guardado)`.

    Un mes al que le falta una semana NO se guarda: si se guardara, sus
    streams bajos se arrastrarían a los cierres de Q, semestre y año, y ahí
    ya no se nota que faltaba algo. Se devuelve igual, para que el reporte
    lo muestre esta vez junto con el aviso; cuando lleguen las semanas que
    faltan y se vuelva a cerrar, queda guardado.
    """
    df, avisos = calcular_mes(anio, mes, bandas)
    if df.empty:
        return df, avisos, False
    if any(MENSAJE_INCOMPLETO in a for a in avisos):
        return df, avisos, False
    orden = config.LABEL_GROUPS_MS
    df = df.copy()
    df["_o"] = df["label_group"].map({lab: i for i, lab in enumerate(orden)})
    df = df.sort_values(["country_code", "banda", "_o"]).drop(columns="_o")
    history.guardar_mes(
        (fila.anio, fila.mes, fila.country_code, int(fila.banda), fila.label_group,
         float(fila.tracks), float(fila.streams_millones), float(fila.pct_streams),
         int(fila.semanas), ORIGEN_CALCULADO)
        for fila in df.itertuples(index=False)
    )
    return df, avisos, True


def meses_cargados(mensual: pd.DataFrame = None) -> list:
    """Los (año, mes) que hay en el histórico mensual, en orden."""
    mensual = history.cargar_ms_mensual() if mensual is None else mensual
    if mensual.empty:
        return []
    return sorted({(int(a), int(m)) for a, m in zip(mensual["anio"], mensual["mes"])})


def mes_anterior(anio: int, mes: int) -> tuple:
    return (anio - 1, 12) if mes == 1 else (anio, mes - 1)


def ultimo_mes_cerrable(hoy: date = None) -> tuple:
    """El último mes que ya terminó. Es el que propone la interfaz cuando el
    usuario abre la ventana del reporte mensual."""
    hoy = hoy or date.today()
    return mes_anterior(hoy.year, hoy.month)


def generar(anio: int, mes: int, salida=None) -> tuple:
    """Cierra el mes y escribe el Excel. Devuelve `(ruta, avisos)`.

    Es lo que llaman la ventana y la línea de comandos: no hace falta cargar
    nada aparte, el mes sale de las semanas de BQ que ya están en el
    histórico.
    """
    df, avisos, guardado = cerrar_mes(anio, mes)
    historico = history.cargar_ms_mensual()
    if not guardado and not df.empty:
        # El mes quedó incompleto y no se guardó: se muestra solo en este
        # archivo, para que se vea lo que hay hasta ahora.
        parcial = df.copy()
        parcial["origen"] = "parcial"
        historico = pd.concat([historico, parcial], ignore_index=True)
    salida = Path(salida or config.OUTPUT_DIR)
    ruta = salida / mensual_reporte.nombre_reporte(anio, mes)
    mensual_reporte.generar_reporte(
        historico, ruta, hasta=(anio, mes),
        cierres=history.cargar_ms_mensual_cierre(),
    )
    return ruta, avisos


def resumen(ruta, avisos, anio: int, mes: int) -> str:
    lineas = [f"Reporte mensual de {config.MESES_ES[mes].lower()} de {anio}: {ruta}"]
    if avisos:
        lineas.append("Avisos:")
        lineas += [f"  ! {a}" for a in avisos]
    return "\n".join(lineas)
