"""Genera el reporte MS BMAT (Promúsica Colombia).

Es un reporte aparte de los dos de Spotify y con otro universo de datos:
BMAT/Promúsica consolida varias plataformas (Napster, GooglePlay, Spotify,
Deezer), tiene sus propias bandas -- hasta el Top 10.000 -- y sus propios 8
sellos (ver config.BMAT_*). Por eso vive en su propio módulo y no reutiliza
nada de market_share.py: coinciden en la forma, no en los datos.

POR AHORA SOLO GENERA LA ESTRUCTURA Y EL FORMATO, con el histórico ya
sembrado (2020 semana 1 -> 2026 semana 23, 336 semanas). Todavía no hay
lógica para calcular una semana nueva a partir de una fuente: cuando se
defina, el lugar para engancharla es una función `append_semana_bmat()` en
history.py, igual que las otras tres tablas, y esto no tendría que cambiar.

Dos hojas, replicando 1:1 el archivo real "MS BMAT COL a Sem 23 de 2026.xlsx":

- "Resumen Promusica Colombia": el detalle. Ocho bandas, y cada una con tres
  sub-tablas de 9 filas (los 8 sellos + Total): TRACKS (cuántas canciones),
  Streams (Millones) y % Streams.
- "Market Share Promusica Colombia": el resumen. Las mismas ocho bandas con
  solo el % de cada sello, y el semáforo de color.

El semáforo es EXACTAMENTE la misma regla del Reporte_MS_TOP200 (verificado
en el formato condicional del archivo real): rojo para el sello que le gana
a Universal dentro de su banda, verde en Universal cuando lidera. Se reusa
la implementación de market_share.py para no tener dos copias de la regla.
"""
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import config, history, market_share

# --- Layout de "Resumen Promusica Colombia" (verificado contra el archivo) ---
# Fila 1 = año (solo en la primera columna de cada año), fila 2 = n° de
# semana, datos desde la columna D. La primera banda arranca en la fila 4 y
# de ahí en adelante cada 32 filas -- salvo el primer salto, que en el
# archivo real es de 33; se respeta para que las filas coincidan si alguien
# compara los dos archivos lado a lado.
_RES_COL_ETIQUETA = 1      # A: TRACKS / Streams (Millones) / % Streams
_RES_COL_SELLO = 2         # B
_RES_COL_PRIMERA_SEMANA = 4  # D
_RES_FILA_ANIO = 1
_RES_FILA_SEMANA = 2
_RES_FILA_PRIMERA_BANDA = 4
_RES_SALTOS_BANDA = [33] + [32] * 6   # entre banda y banda
# Desplazamiento de cada sub-tabla respecto de la fila de su banda.
_RES_SUBTABLAS = [("TRACKS", "tracks", 2, "#,##0"),
                  ("Streams \n(Millones)", "streams_millones", 12, "0.00"),
                  ("% \nStreams", "pct_streams", 22, "0.0%")]

# --- Layout de "Market Share Promusica Colombia" ---
# Fila 2 = "Sem N", datos desde la columna E. Banda en la columna B
# (combinada sobre sus 8 filas de sello), sello en la columna C.
_MS_COL_BANDA = 2          # B
_MS_COL_SELLO = 3          # C
_MS_COL_PRIMERA_SEMANA = 5  # E
_MS_FILA_ANIO = 1
_MS_FILA_SEMANA = 2
_MS_FILA_PRIMERA_BANDA = 4
_MS_SALTO_BANDA = 9        # 8 sellos + 1 fila en blanco


def cargar_historico(country_code: str = None) -> pd.DataFrame:
    """Histórico BMAT tal cual está guardado, ordenado por semana."""
    country_code = country_code or config.BMAT_PAIS
    df = history.cargar_bmat_weekly()
    if df.empty:
        return df
    return df[df["country_code"] == country_code].sort_values(["anio", "semana"])


def _semanas(df: pd.DataFrame) -> list:
    """Las semanas con datos, en orden: [(anio, semana), ...]."""
    if df.empty:
        return []
    return [
        (int(r.anio), int(r.semana))
        for r in df[["anio", "semana"]].drop_duplicates()
                   .sort_values(["anio", "semana"]).itertuples(index=False)
    ]


def _nombre_banda(banda: int, mayusculas: bool = False) -> str:
    """10 -> 'Top 10'; 1000 -> 'Top 1.000' (separador de miles, como el
    archivo real en sus bandas grandes)."""
    numero = f"{banda:,}".replace(",", ".")
    return f"{'TOP' if mayusculas else 'Top'} {numero}"


def _escribir_encabezado_semanas(ws, semanas, fila_anio, fila_semana,
                                 primera_col, con_prefijo_sem, relleno, negrita_blanca):
    """Fila de año + fila de semana. El año se escribe UNA vez, en su primera
    columna, igual que el archivo real."""
    anio_anterior = None
    for i, (anio, semana) in enumerate(semanas):
        col = primera_col + i
        if anio != anio_anterior:
            ws.cell(row=fila_anio, column=col, value=anio)
            anio_anterior = anio
        celda = ws.cell(row=fila_semana, column=col,
                        value=f"Sem\n{semana}" if con_prefijo_sem else semana)
        celda.font = negrita_blanca
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 9.5


def _escribir_resumen(ws, df: pd.DataFrame) -> None:
    """La hoja de detalle: por banda, las tres sub-tablas (TRACKS, Streams,
    % Streams) con los 8 sellos y su Total."""
    negrita = Font(bold=True)
    blanca_negrita = Font(bold=True, color="FFFFFF")
    relleno = PatternFill(start_color=config.COLOR_BANNER_BMAT,
                          end_color=config.COLOR_BANNER_BMAT, fill_type="solid")
    centrado = Alignment(horizontal="center", vertical="center", wrap_text=True)

    titulo = ws.cell(row=2, column=_RES_COL_ETIQUETA, value=config.BMAT_TITULO_RESUMEN)
    titulo.font = negrita
    titulo.alignment = centrado
    ws.merge_cells(start_row=2, start_column=_RES_COL_ETIQUETA, end_row=2, end_column=_RES_COL_SELLO)

    semanas = _semanas(df)
    _escribir_encabezado_semanas(ws, semanas, _RES_FILA_ANIO, _RES_FILA_SEMANA,
                                 _RES_COL_PRIMERA_SEMANA, False, relleno, blanca_negrita)

    # (anio, semana, banda, sello) -> fila del histórico, para no filtrar el
    # DataFrame una vez por celda.
    por_celda = {
        (int(r.anio), int(r.semana), int(r.banda), r.label_group): r
        for r in df.itertuples(index=False)
    }

    fila_banda = _RES_FILA_PRIMERA_BANDA
    for i, banda in enumerate(config.BMAT_BANDAS):
        celda_banda = ws.cell(row=fila_banda, column=_RES_COL_PRIMERA_SEMANA,
                              value=_nombre_banda(banda))
        celda_banda.font = blanca_negrita
        celda_banda.fill = relleno
        celda_banda.alignment = centrado

        for etiqueta, campo, offset, formato in _RES_SUBTABLAS:
            fila0 = fila_banda + offset
            celda_et = ws.cell(row=fila0, column=_RES_COL_ETIQUETA, value=etiqueta)
            celda_et.font = negrita
            celda_et.alignment = centrado

            for k, label in enumerate(config.BMAT_LABELS):
                ws.cell(row=fila0 + k, column=_RES_COL_SELLO, value=label)
            celda_total = ws.cell(row=fila0 + len(config.BMAT_LABELS),
                                  column=_RES_COL_SELLO, value="Total")
            celda_total.font = negrita

            for j, (anio, semana) in enumerate(semanas):
                col = _RES_COL_PRIMERA_SEMANA + j
                total = 0.0
                hay_dato = False
                for k, label in enumerate(config.BMAT_LABELS):
                    fila_hist = por_celda.get((anio, semana, banda, label))
                    valor = getattr(fila_hist, campo, None) if fila_hist is not None else None
                    if valor is not None and pd.notna(valor):
                        total += float(valor)
                        hay_dato = True
                    else:
                        valor = None
                    celda = ws.cell(row=fila0 + k, column=col, value=valor)
                    celda.number_format = formato
                celda_tot = ws.cell(row=fila0 + len(config.BMAT_LABELS), column=col,
                                    value=total if hay_dato else None)
                celda_tot.number_format = formato
                celda_tot.font = negrita

        if i < len(_RES_SALTOS_BANDA):
            fila_banda += _RES_SALTOS_BANDA[i]

    # La nota de fuente va en la fila de la segunda banda, como en el archivo.
    fila_fuente = _RES_FILA_PRIMERA_BANDA + _RES_SALTOS_BANDA[0]
    celda_fuente = ws.cell(row=fila_fuente, column=_RES_COL_ETIQUETA, value=config.BMAT_FUENTE)
    celda_fuente.font = negrita
    ws.merge_cells(start_row=fila_fuente, start_column=_RES_COL_ETIQUETA,
                   end_row=fila_fuente, end_column=_RES_COL_SELLO)

    ws.column_dimensions[get_column_letter(_RES_COL_ETIQUETA)].width = 8.5
    ws.column_dimensions[get_column_letter(_RES_COL_SELLO)].width = 13.5
    ws.column_dimensions[get_column_letter(3)].width = 2
    ws.freeze_panes = ws.cell(row=_RES_FILA_PRIMERA_BANDA,
                              column=_RES_COL_PRIMERA_SEMANA).coordinate


def _escribir_market_share(ws, df: pd.DataFrame) -> None:
    """La hoja resumen: solo el % por banda y sello, con el semáforo."""
    negrita = Font(bold=True)
    blanca_negrita = Font(bold=True, color="FFFFFF")
    relleno = PatternFill(start_color=config.COLOR_BANNER_BMAT,
                          end_color=config.COLOR_BANNER_BMAT, fill_type="solid")
    centrado = Alignment(horizontal="center", vertical="center", wrap_text=True)

    titulo = ws.cell(row=2, column=_MS_COL_BANDA, value=config.BMAT_TITULO_MARKET_SHARE)
    titulo.font = negrita
    titulo.alignment = centrado
    ws.merge_cells(start_row=2, start_column=_MS_COL_BANDA, end_row=2, end_column=_MS_COL_SELLO)

    semanas = _semanas(df)
    _escribir_encabezado_semanas(ws, semanas, _MS_FILA_ANIO, _MS_FILA_SEMANA,
                                 _MS_COL_PRIMERA_SEMANA, True, relleno, blanca_negrita)

    pct_por_celda = {
        (int(r.anio), int(r.semana), int(r.banda), r.label_group): r.pct_streams
        for r in df.itertuples(index=False)
    }

    for i, banda in enumerate(config.BMAT_BANDAS):
        fila0 = _MS_FILA_PRIMERA_BANDA + i * _MS_SALTO_BANDA
        celda_banda = ws.cell(row=fila0, column=_MS_COL_BANDA,
                              value=_nombre_banda(banda, mayusculas=True).replace(" ", "\n"))
        celda_banda.font = negrita
        celda_banda.alignment = centrado
        ws.merge_cells(start_row=fila0, start_column=_MS_COL_BANDA,
                       end_row=fila0 + len(config.BMAT_LABELS) - 1, end_column=_MS_COL_BANDA)

        for k, label in enumerate(config.BMAT_LABELS):
            ws.cell(row=fila0 + k, column=_MS_COL_SELLO, value=label)

        for j, (anio, semana) in enumerate(semanas):
            col = _MS_COL_PRIMERA_SEMANA + j
            valores = {lab: pct_por_celda.get((anio, semana, banda, lab))
                       for lab in config.BMAT_LABELS}
            referencia = market_share._referencia_grupo(valores)
            for k, label in enumerate(config.BMAT_LABELS):
                valor = valores[label]
                celda = ws.cell(row=fila0 + k, column=col, value=valor)
                celda.number_format = "0.0%"
                market_share._pintar_vs_universal(celda, valor, label, referencia)

    ws.column_dimensions[get_column_letter(1)].width = 3.5
    ws.column_dimensions[get_column_letter(_MS_COL_BANDA)].width = 10.5
    ws.column_dimensions[get_column_letter(_MS_COL_SELLO)].width = 22.5
    ws.column_dimensions[get_column_letter(4)].width = 1.2
    ws.freeze_panes = ws.cell(row=_MS_FILA_PRIMERA_BANDA,
                              column=_MS_COL_PRIMERA_SEMANA).coordinate


def generar_reporte(output_path: Path, country_code: str = None) -> Path:
    """Escribe el reporte BMAT completo con el histórico que haya guardado.

    No recibe una fuente de datos porque todavía no está definida la lógica
    para calcular una semana nueva (ver el docstring del módulo): hoy el
    reporte es un reflejo del histórico sembrado.
    """
    df = cargar_historico(country_code)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ws_resumen = writer.book.create_sheet(config.BMAT_SHEET_RESUMEN)
        _escribir_resumen(ws_resumen, df)
        ws_ms = writer.book.create_sheet(config.BMAT_SHEET_MARKET_SHARE)
        _escribir_market_share(ws_ms, df)

    return output_path
