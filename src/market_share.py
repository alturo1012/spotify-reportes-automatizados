"""Genera el Reporte_MS_MS_TOP_200_Spotify (Market Share YTD) a partir del
histórico acumulado en src/history.py.

Fórmula real (verificada contra PLANTILLA_SEMANAL_MS_TOP200.xlsx, filas
99-112 — ver claude/mapeo_logica_plantillas.md sección 1):

    % Market Share YTD (sello, país) =
        suma streams_top200 del sello, semanas 1..N del año
        -------------------------------------------------------
        suma streams_top200 de los 7 sellos, semanas 1..N del año

Se compara el mismo número de semanas (1..N) entre el año en curso y el
año anterior. NO es un promedio de porcentajes semanales.

La pestaña resumen "% Market Share" se escribe como una cuadrícula de
tablas, una por país (4 por fila de bloques), replicando el formato visual
de PLANTILLA_SEMANAL_MS_TOP200.xlsx -- ver _escribir_resumen_pct.

Las pestañas individuales por país (una por config.PAISES_MS) son ahora una
cuadrícula semanal: columnas = semanas ya guardadas en
history.chart_track_weekly (fecha + n° de semana), filas = las 5 bandas de
config.BANDAS_MARKET_SHARE, cada una con las 7 filas de sello de
config.LABEL_GROUPS_MS -- replica la sub-tabla "Streams (%)" de
PLANTILLA_SEMANAL_MS_TOP200.xlsx (ver _escribir_pagina_pais / Ajuste 7 en
el plan). Las otras dos sub-tablas de la plantilla real por banda ("Tracks"
y "Streams" crudo) se dejaron fuera a pedido del usuario -- solo la de "%".

El histórico de esa cuadrícula vive en `history.ms_band_label_weekly` y
tiene dos orígenes que se complementan: se sembró de una vez con los años
que ya traía calculados el reporte real (2021 semana 1 -> 2026 semana 33,
293 semanas x 17 países) y, de ahí en adelante, cada semana nueva se agrega
sola al generar el reporte (`history.append_semana_ms_bandas`, misma
fórmula). Antes se calculaba al vuelo desde `chart_track_weekly` y por eso
salía casi vacía -- ver "Actualización a la semana 33" en
claude/plan_fusion_paso_a_paso.md.
"""
from pathlib import Path
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import config, history


def calcular_ytd_por_pais(
    country_code: str, anio_actual: int, hasta_semana: int
) -> pd.DataFrame:
    """% Market Share YTD de un país: un DataFrame con una fila por sello
    (en el orden de config.LABEL_GROUPS_MS), comparando anio_actual vs.
    anio_actual - 1, mismas semanas 1..N en ambos.

    N = min(hasta_semana, semanas realmente disponibles de anio_actual - 1).
    Esto evita comparar, por ejemplo, 25 semanas de 2026 contra solo 24 de
    2025 si el histórico del año anterior no llega tan lejos todavía — la
    comparación YTD deja de tener sentido si no son las mismas semanas.
    """
    anio_anterior = anio_actual - 1

    historico_pais_anterior = history.cargar_ms_label_weekly()
    historico_pais_anterior = historico_pais_anterior[
        (historico_pais_anterior["country_code"] == country_code)
        & (historico_pais_anterior["anio"] == anio_anterior)
    ]
    semana_max_anterior = (
        int(historico_pais_anterior["semana"].max())
        if not historico_pais_anterior.empty
        else 0
    )
    hasta_semana_efectiva = min(hasta_semana, semana_max_anterior)

    ytd_actual = history.query_ytd_ms(anio_actual, hasta_semana_efectiva)
    ytd_actual = ytd_actual[ytd_actual["country_code"] == country_code]
    ytd_anterior = history.query_ytd_ms(anio_anterior, hasta_semana_efectiva)
    ytd_anterior = ytd_anterior[ytd_anterior["country_code"] == country_code]

    streams_actual = ytd_actual.groupby("label_group")["streams_top200"].sum()
    streams_anterior = ytd_anterior.groupby("label_group")["streams_top200"].sum()

    total_actual = streams_actual.sum()
    total_anterior = streams_anterior.sum()

    filas = []
    for label in config.LABEL_GROUPS_MS:
        s_actual = float(streams_actual.get(label, 0.0))
        s_anterior = float(streams_anterior.get(label, 0.0))
        pct_actual = s_actual / total_actual if total_actual else 0.0
        pct_anterior = s_anterior / total_anterior if total_anterior else 0.0
        filas.append({
            "label_group": label,
            f"pct_YTD_{anio_actual}": pct_actual,
            f"pct_YTD_{anio_anterior}": pct_anterior,
            "g_l": pct_actual - pct_anterior,
        })
    return pd.DataFrame(filas)


def calcular_streams_pct_grid(country_code: str) -> pd.DataFrame:
    """Tabla larga (tidy) con el % de streams por banda y sello, semana a
    semana, para un país -- lo que muestran las sub-tablas "Streams (%)
    TOP N" del reporte real.

    Sale directo de `history.ms_band_label_weekly`, que trae dos cosas
    juntas: el histórico sembrado del reporte oficial (2021 en adelante) y
    cada semana nueva que se va cargando (la calcula
    `history.append_semana_ms_bandas` con la misma fórmula). Antes esto se
    calculaba al vuelo desde `chart_track_weekly`, que solo tenía las
    semanas cargadas después del Ajuste 5 -- por eso la cuadrícula salía
    casi vacía.

    Una fila por (anio, semana, chart_date, banda, label_group, pct).
    """
    columnas = ["anio", "semana", "chart_date", "banda", "label_group", "pct"]
    historico = history.cargar_ms_band_label_weekly()
    if historico.empty:
        return pd.DataFrame(columns=columnas)

    del_pais = historico[historico["country_code"] == country_code]
    if del_pais.empty:
        return pd.DataFrame(columns=columnas)

    grid = del_pais.rename(columns={"pct_streams": "pct"})[columnas]
    return grid.reset_index(drop=True)


def construir_resumen_pct(anio_actual: int, hasta_semana: int) -> pd.DataFrame:
    """Tabla larga (tidy) con el % Market Share YTD de los 17 países, para
    la pestaña resumen "% Market Share". Una fila por país/sello.
    """
    partes = []
    for country_code in config.PAISES_MS:
        tabla_pais = calcular_ytd_por_pais(country_code, anio_actual, hasta_semana)
        tabla_pais.insert(0, "country_code", country_code)
        partes.append(tabla_pais)
    return pd.concat(partes, ignore_index=True)


# Layout de la pestaña resumen "% Market Share", replicando el formato
# visual de PLANTILLA_SEMANAL_MS_TOP200.xlsx (verificado 1:1 contra ese
# archivo real -- fila 2 banner, fila 4 encabezado de bloque, fila 6
# subencabezado, filas 7-13 sellos, luego 1 fila en blanco y se repite):
#   - Banner "TOP 200 WEEKLY MARKET SHARE" arriba de todo (fila 2), y
#     freeze_panes en A4 para que las filas 1-3 queden siempre visibles al
#     desplazarse hacia abajo (pedido explícito del usuario).
#   - Un bloque de 4 columnas por país (nombre, YTD año actual, YTD año
#     anterior, G/L), con el nombre del país en un banner arriba del bloque
#     y 7 filas de sellos debajo (config.ORDEN_LABELS_MS_RESUMEN). 4 bloques
#     de país por fila, con una columna angosta de separación entre cada
#     uno, igual que la plantilla real.
_MS_COL_INICIAL = 2  # columna B, igual que la plantilla real (A queda vacía)
_MS_COLS_POR_BLOQUE = 4  # nombre, YTD actual, YTD anterior, G/L
_MS_BLOQUES_POR_FILA = 4
_MS_FILAS_POR_BLOQUE = 11  # banner + blanco + subencabezado + 7 sellos + blanco
_MS_FILA_BANNER_TOP200 = 2
_MS_FILA_PRIMER_BLOQUE = 4


# ---------------------------------------------------------------------------
# Semáforo del Reporte_MS_TOP200. La regla se sacó del formato condicional
# REAL del reporte oficial (las fórmulas y los dxf de MS TOP 200 a la Sem 33),
# no de mirar una captura, y es LA MISMA en las dos pestañas:
#
#   - Se pinta de ROJO el sello cuyo porcentaje le GANA a Universal dentro de
#     su grupo. Es un reporte de Universal: lo que se quiere ver de un vistazo
#     es quién va por delante.
#   - Si Universal es el más alto del grupo, su propia celda va en VERDE.
#
#   El "grupo" es lo único que cambia entre las dos pestañas:
#     · resumen "% Market Share": los 7 sellos de un país, por columna YTD.
#     · pestaña de país: los 7 sellos de una banda, en la columna de esa semana.
#
#   La columna G/L (solo en el resumen) tiene su propia regla, por signo:
#   VERDE si ganó participación, ROJO si perdió, AMARILLO si quedó igual. En
#   el archivo real son tres reglas contra los umbrales de las celdas
#   U1/V1/W1/X1, que están VACÍAS, y Excel las lee como 0 -- o sea, el signo.
#
# DOS ERRORES DE LA PLANTILLA ORIGINAL QUE ACÁ NO SE REPLICAN (a propósito).
# Los dos están en la pestaña resumen; las pestañas de país sí la tienen bien:
#   1. En 4 de los 5 bloques de fila, las reglas de las dos últimas filas
#      (Indies y Virgin) comparan contra la fila del SUBENCABEZADO en vez de
#      contra la de Universal (C17 en vez de C18, C28 en vez de C29, etc.).
#      Como esa celda tiene el nombre del país (texto), en Excel un número
#      nunca es "mayor que" un texto: esas dos filas NO se pintan nunca,
#      aunque le ganen a Universal. Afecta a 13 de los 17 países.
#   2. Los bloques 3 y 4 evalúan el verde con la fórmula del bloque 2
#      (C18=MAX(C18:C24) sin actualizar al copiar), y el MAX del bloque 1
#      deja a Virgin fuera del rango (C7:C12 en vez de C7:C13).
# Acá la regla se aplica igual para los 7 sellos de los 17 países, así que
# van a salir en rojo algunas celdas que en el archivo viejo no se pintaban.
# ---------------------------------------------------------------------------

def _referencia_grupo(valores_por_label: dict):
    """Arma la referencia del semáforo para un grupo de 7 sellos.

    `valores_por_label` es {label: valor o None}. Devuelve
    {"universal": v, "maximo": m} o None si el grupo no tiene datos (todo en
    cero o vacío): en ese caso no se pinta nada, porque si no Universal
    saldría en verde por empate técnico y eso engaña.
    """
    presentes = [v for v in valores_por_label.values() if v is not None]
    if not presentes or sum(presentes) == 0:
        return None
    return {"universal": valores_por_label.get("Universal"), "maximo": max(presentes)}


def _pintar(celda, fondo: str, texto: str) -> None:
    celda.fill = PatternFill(start_color=fondo, end_color=fondo, fill_type="solid")
    celda.font = Font(color=texto)


def _pintar_vs_universal(celda, valor, label: str, referencia) -> None:
    """Rojo si le gana a Universal; verde en la celda de Universal si lidera."""
    if valor is None or referencia is None or referencia["universal"] is None:
        return
    if label == "Universal":
        if valor < referencia["maximo"]:
            return
        _pintar(celda, config.COLOR_SEMAFORO_VERDE, config.COLOR_TEXTO_SEMAFORO_VERDE)
    elif valor > referencia["universal"]:
        _pintar(celda, config.COLOR_SEMAFORO_ROJO, config.COLOR_TEXTO_SEMAFORO_ROJO)


def _pintar_gl(celda, valor) -> None:
    """Columna G/L del resumen: verde si ganó, rojo si perdió, amarillo si igual."""
    if valor is None:
        return
    if valor > 0:
        _pintar(celda, config.COLOR_SEMAFORO_VERDE, config.COLOR_TEXTO_SEMAFORO_VERDE)
    elif valor < 0:
        _pintar(celda, config.COLOR_SEMAFORO_ROJO, config.COLOR_TEXTO_SEMAFORO_ROJO)
    else:
        _pintar(celda, config.COLOR_SEMAFORO_AMARILLO, config.COLOR_TEXTO_SEMAFORO_AMARILLO)


def _escribir_resumen_pct(ws, anio_actual: int, hasta_semana: int) -> None:
    """Escribe la pestaña resumen "% Market Share" como una cuadrícula de
    tablas por país (ver constantes _MS_* arriba) -- se arma directo con
    openpyxl (no con `.to_excel(...)`) por el mismo motivo que en
    chart_semanal.py: acá hace falta control celda por celda (banners
    combinados, freeze_panes) que `.to_excel(...)` no ofrece.
    """
    negrita_banner = Font(bold=True, color=config.COLOR_BANNER_MS_TEXTO)
    relleno_banner = PatternFill(
        start_color=config.COLOR_BANNER_MS_FONDO, end_color=config.COLOR_BANNER_MS_FONDO, fill_type="solid"
    )
    centrado = Alignment(horizontal="center", vertical="center")
    negrita = Font(bold=True)

    ancho_total_columnas = _MS_BLOQUES_POR_FILA * (_MS_COLS_POR_BLOQUE + 1) - 1
    columna_final = _MS_COL_INICIAL + ancho_total_columnas - 1
    celda_top200 = ws.cell(row=_MS_FILA_BANNER_TOP200, column=_MS_COL_INICIAL, value="TOP 200 WEEKLY MARKET SHARE")
    ws.merge_cells(
        start_row=_MS_FILA_BANNER_TOP200, start_column=_MS_COL_INICIAL,
        end_row=_MS_FILA_BANNER_TOP200, end_column=columna_final,
    )
    celda_top200.font = Font(bold=True, size=14, color=config.COLOR_BANNER_MS_TEXTO)
    celda_top200.fill = relleno_banner
    celda_top200.alignment = centrado

    etiqueta_actual = f"YTD {str(anio_actual)[-2:]}"
    etiqueta_anterior = f"YTD {str(anio_actual - 1)[-2:]}"

    for i, country_code in enumerate(config.ORDEN_PAISES_MS_RESUMEN):
        fila_bloque = i // _MS_BLOQUES_POR_FILA
        col_bloque = i % _MS_BLOQUES_POR_FILA
        fila_header = _MS_FILA_PRIMER_BLOQUE + fila_bloque * _MS_FILAS_POR_BLOQUE
        fila_subheader = fila_header + 2
        fila_primer_dato = fila_subheader + 1
        col_inicio = _MS_COL_INICIAL + col_bloque * (_MS_COLS_POR_BLOQUE + 1)
        col_fin = col_inicio + _MS_COLS_POR_BLOQUE - 1

        nombre_visible = config.NOMBRE_PAIS_MS_RESUMEN[country_code]

        celda_pais = ws.cell(row=fila_header, column=col_inicio, value=nombre_visible)
        ws.merge_cells(start_row=fila_header, start_column=col_inicio, end_row=fila_header, end_column=col_fin)
        celda_pais.font = negrita_banner
        celda_pais.fill = relleno_banner
        celda_pais.alignment = centrado

        encabezados = [nombre_visible, etiqueta_actual, etiqueta_anterior, "G/L"]
        for j, texto in enumerate(encabezados):
            celda = ws.cell(row=fila_subheader, column=col_inicio + j, value=texto)
            celda.font = negrita
            celda.alignment = centrado

        tabla_pais = calcular_ytd_por_pais(country_code, anio_actual, hasta_semana)
        tabla_pais = tabla_pais.set_index("label_group")

        campo_actual = f"pct_YTD_{anio_actual}"
        campo_anterior = f"pct_YTD_{anio_actual - 1}"

        def valor_de(label, campo):
            if label not in tabla_pais.index or campo not in tabla_pais.columns:
                return None
            return float(tabla_pais.loc[label, campo])

        # Una referencia por columna YTD: el grupo es "los 7 sellos del país".
        referencia = {
            campo: _referencia_grupo({lab: valor_de(lab, campo) for lab in config.ORDEN_LABELS_MS_RESUMEN})
            for campo in (campo_actual, campo_anterior)
        }

        for k, label in enumerate(config.ORDEN_LABELS_MS_RESUMEN):
            r = fila_primer_dato + k
            ws.cell(row=r, column=col_inicio, value=label)
            for col_offset, campo in enumerate(
                [campo_actual, campo_anterior, "g_l"], start=1
            ):
                valor = valor_de(label, campo)
                celda_valor = ws.cell(row=r, column=col_inicio + col_offset, value=valor)
                celda_valor.number_format = "0.0%"
                if campo == "g_l":
                    _pintar_gl(celda_valor, valor)
                else:
                    _pintar_vs_universal(celda_valor, valor, label, referencia[campo])

        for col in range(col_inicio, col_fin + 1):
            ws.column_dimensions[get_column_letter(col)].width = 11
        ws.column_dimensions[get_column_letter(col_fin + 1)].width = 2  # separadora

    ws.freeze_panes = "A4"


# Layout de la pestaña individual de cada país (ver calcular_streams_pct_grid
# y el docstring del módulo). Verificado contra la pestaña "CO" real de
# PLANTILLA_SEMANAL_MS_TOP200.xlsx: columna A = banda (combinada verticalmente
# sobre sus 7 filas de sello), columna B = sello, columnas C+ = una por
# semana. Fila 2 = fecha, fila 3 = n° de semana -- esas dos filas y las
# columnas A/B quedan fijas (freeze_panes), igual que las columnas de tops y
# disquera quedan fijas al desplazarse a la derecha (pedido explícito del
# usuario).
_MSPAIS_COL_BANDA = 1
_MSPAIS_COL_LABEL = 2
_MSPAIS_COL_PRIMERA_SEMANA = 3
_MSPAIS_FILA_TITULO = 1
_MSPAIS_FILA_FECHA = 2
_MSPAIS_FILA_SEMANA = 3
_MSPAIS_FILA_PRIMER_DATO = 4


def _escribir_pagina_pais(ws, country_code: str) -> None:
    """Escribe la pestaña individual de un país como cuadrícula semanal
    (ver constantes _MSPAIS_* arriba). Si todavía no hay ninguna semana
    guardada en chart_track_weekly para este país, la cuadrícula sale sin
    columnas de datos (solo encabezados) -- se va llenando sola a medida
    que se cargan bases nuevas.
    """
    negrita = Font(bold=True)
    centrado = Alignment(horizontal="center", vertical="center", wrap_text=True)

    nombre_visible = config.NOMBRE_PAIS_MS_RESUMEN.get(country_code, country_code)
    celda_titulo = ws.cell(
        row=_MSPAIS_FILA_TITULO, column=_MSPAIS_COL_BANDA,
        value=f"{nombre_visible} - % Market Share por banda (Streams)",
    )
    celda_titulo.font = Font(bold=True, size=12)

    grid = calcular_streams_pct_grid(country_code)
    semanas = (
        grid[["anio", "semana", "chart_date"]]
        .drop_duplicates()
        .sort_values(["anio", "semana"])
        .reset_index(drop=True)
    )
    lista_semanas = list(semanas.itertuples(index=False))
    pct_por_celda = {
        (r.anio, r.semana, r.banda, r.label_group): r.pct for r in grid.itertuples(index=False)
    }

    ws.cell(row=_MSPAIS_FILA_SEMANA, column=_MSPAIS_COL_LABEL, value="Etiquetas\nde fila").font = negrita

    for j, s in enumerate(lista_semanas):
        col = _MSPAIS_COL_PRIMERA_SEMANA + j
        celda_fecha = ws.cell(row=_MSPAIS_FILA_FECHA, column=col, value=pd.Timestamp(s.chart_date))
        celda_fecha.number_format = "yyyy-mm-dd"
        celda_fecha.font = negrita
        celda_fecha.alignment = centrado
        celda_semana = ws.cell(row=_MSPAIS_FILA_SEMANA, column=col, value=int(s.semana))
        celda_semana.font = negrita
        celda_semana.alignment = centrado
        ws.column_dimensions[get_column_letter(col)].width = 11

    # Referencia del semáforo: acá el grupo son los 7 sellos de UNA banda en
    # UNA semana (en el resumen son los 7 sellos del país). Misma regla, ver
    # _pintar_vs_universal.
    referencia_por_celda = {
        (s.anio, s.semana, banda): _referencia_grupo({
            lab: pct_por_celda.get((s.anio, s.semana, banda, lab))
            for lab in config.LABEL_GROUPS_MS
        })
        for s in lista_semanas
        for banda in config.BANDAS_MARKET_SHARE
    }

    fila = _MSPAIS_FILA_PRIMER_DATO
    for banda in config.BANDAS_MARKET_SHARE:
        fila_inicio_banda = fila
        for label in config.LABEL_GROUPS_MS:
            ws.cell(row=fila, column=_MSPAIS_COL_LABEL, value=label)
            for j, s in enumerate(lista_semanas):
                col = _MSPAIS_COL_PRIMERA_SEMANA + j
                valor = pct_por_celda.get((s.anio, s.semana, banda, label))
                celda_valor = ws.cell(row=fila, column=col, value=valor)
                celda_valor.number_format = "0.0%"
                _pintar_vs_universal(
                    celda_valor, valor, label, referencia_por_celda[(s.anio, s.semana, banda)]
                )
            fila += 1

        celda_banda = ws.cell(row=fila_inicio_banda, column=_MSPAIS_COL_BANDA, value=f"Streams\n(%)\nTOP {banda}")
        ws.merge_cells(
            start_row=fila_inicio_banda, start_column=_MSPAIS_COL_BANDA,
            end_row=fila - 1, end_column=_MSPAIS_COL_BANDA,
        )
        celda_banda.font = negrita
        celda_banda.alignment = centrado
        fila += 1  # fila en blanco separadora entre bandas, igual que la plantilla real

    ws.column_dimensions[get_column_letter(_MSPAIS_COL_BANDA)].width = 12
    ws.column_dimensions[get_column_letter(_MSPAIS_COL_LABEL)].width = 14

    ws.freeze_panes = ws.cell(row=_MSPAIS_FILA_PRIMER_DATO, column=_MSPAIS_COL_PRIMERA_SEMANA).coordinate


def generar_reporte(
    df_semana: pd.DataFrame, output_path: Path, guardar_en_historico: bool = True
) -> Path:
    """Genera el reporte de Market Share a partir del DataFrame de la semana
    nueva (el que devuelve load_data.load_source). Guarda esa semana en el
    histórico (a menos que ya se haya guardado antes, con
    guardar_en_historico=False) y calcula el YTD con todo el histórico
    acumulado hasta esa semana.
    """
    if guardar_en_historico:
        history.append_semana_ms(df_semana)
        # Alimenta la cuadrícula semanal de las pestañas por país. Va junto
        # a append_semana_ms (mismo reporte, misma corrida) para que las dos
        # tablas avancen a la par.
        history.append_semana_ms_bandas(df_semana)

    fecha = pd.Timestamp(df_semana["chart_date"].unique()[0])
    anio_actual = fecha.year

    todo_el_historico = history.cargar_ms_label_weekly()
    hasta_semana = int(todo_el_historico.loc[todo_el_historico["anio"] == anio_actual, "semana"].max())

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ws_resumen = writer.book.create_sheet(config.MS_SHEET_PORCENTAJE)
        _escribir_resumen_pct(ws_resumen, anio_actual, hasta_semana)
        for country_code in config.PAISES_MS:
            ws_pais = writer.book.create_sheet(country_code)
            _escribir_pagina_pais(ws_pais, country_code)

    return output_path
