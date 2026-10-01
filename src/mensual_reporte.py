"""Escribe el Excel del reporte mensual "Market Share Spotify Latam".

Tres tipos de hoja, igual que el archivo real:

- **"XX-Det"** (una por país): el cálculo. Por cada banda, tres bloques de 7
  filas -- `Tracks TOP N`, `Streams TOP N` y `Streams (%) TOP N`.
- **hojas de país** ("CO", "PE"...): solo el % de streams, con los sellos en
  el orden de presentación y las 5 bandas.
- **"Resumen"**: el % de Universal de cada país, en las bandas 10/50/100/200.

Las columnas (y en el Resumen, las filas) son el calendario: los meses de
cada año más sus cierres Q1..Q4, H1 y el año.

REGLAS DE LOS CIERRES (Q / H1 / año)
------------------------------------
Salen de la plantilla real y se comprobaron celda por celda contra ella
(2019 en adelante, 100% de coincidencia; 2018, que fue el primer año y tiene
valores escritos a mano, no cuadra siempre):

- **streams**: suma de los meses del período.
- **tracks**: promedio simple de los meses del período.
- **%**: se recalcula sobre los streams del período (streams del sello /
  streams de los 7 sellos), no es el promedio de los % mensuales.
- **el año, en tracks**, es la excepción: en la plantilla NO es el promedio
  de los 12 meses sino el promedio de tres columnas, `(H1 + Q3 + Q4) / 3`.
  Da distinto que el promedio de los meses (en Colombia 2018: 3,74 contra
  3,41). Se replica tal cual para que el reporte generado dé lo mismo que el
  que venían haciendo a mano. DECIDIDO el 30/09/2026: se deja así, no se
  "corrige" al promedio de los doce meses.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config


AZUL = "FF002060"
GRIS = "FFD9D9D9"
BLANCO = "FFFFFFFF"

FMT_TRACKS = '#,##0_ ;[Red]\\-#,##0\\ '
FMT_STREAMS = '#,##0.00_ ;[Red]\\-#,##0.00\\ '
FMT_PCT = "0%"

_BORDE = Border(*[Side(style="thin", color="FFBFBFBF")] * 4)

# Etiquetas de los tres bloques de cada banda en las hojas Det.
BLOQUES = [
    ("tracks", "Tracks\nTOP\n{banda}", FMT_TRACKS),
    ("streams", "Streams\nTOP\n{banda}", FMT_STREAMS),
    ("pct", "Streams\n(%)\nTOP {banda}", FMT_PCT),
]

# Cierres: qué meses entran en cada uno y en qué orden van las columnas.
CIERRES = {"Q1": (1, 2, 3), "Q2": (4, 5, 6), "Q3": (7, 8, 9), "Q4": (10, 11, 12),
           "H1": (1, 2, 3, 4, 5, 6)}
SECUENCIA = [1, 2, 3, "Q1", 4, 5, 6, "Q2", "H1", 7, 8, 9, "Q3", 10, 11, 12, "Q4", "YEAR"]


def calendario(desde: tuple, hasta: tuple) -> list:
    """La secuencia de columnas (o de filas, en el Resumen) entre dos meses.

    Cada entrada es `(anio, clave)`, donde clave es un número de mes o el
    nombre de un cierre. El primer año arranca en el mes de `desde` (la
    plantilla real empieza en mayo de 2017, sin enero-abril ni Q1), y el
    último año se dibuja completo aunque todavía no tenga datos: así las
    columnas no se mueven de lugar cada mes.
    """
    anio_desde, mes_desde = desde
    salida = []
    for anio in range(anio_desde, hasta[0] + 1):
        for clave in SECUENCIA:
            if anio == anio_desde:
                if isinstance(clave, int) and clave < mes_desde:
                    continue
                if clave in CIERRES and max(CIERRES[clave]) < mes_desde:
                    continue
            salida.append((anio, clave))
    return salida


def etiqueta(anio: int, clave) -> str:
    if isinstance(clave, int):
        return config.MESES_ES_ABREV[clave].capitalize()
    if clave == "YEAR":
        return f"Total\n{anio}"
    if clave == "H1":
        return f"H1\n{anio}"
    return clave


class Valores:
    """Los números de un país: mes a mes y por banda, y los cierres
    calculados a partir de ellos."""

    def __init__(self, filas, hasta: tuple = None, cierres=None):
        # (banda, mes) -> {label: (tracks, streams)}
        self.por_mes = {}
        # Meses que recalculó el programa desde sus semanas: si alguno cae
        # dentro de un período, el cierre se recalcula en vez de usar el que
        # venía en la plantilla.
        self.recalculados = set()
        for fila in filas:
            mes = (int(fila["anio"]), int(fila["mes"]))
            self.por_mes.setdefault((int(fila["banda"]), mes), {})[fila["label_group"]] = (
                fila["tracks"], fila["streams_millones"])
            if fila.get("origen") and fila["origen"] != "plantilla":
                self.recalculados.add(mes)
        # (anio, periodo, banda, label) -> (tracks, streams, pct)
        self.cierres = cierres or {}
        self.hasta = hasta
        self._cache = {}

    def cierre_heredado(self, banda: int, anio: int, clave):
        """El cierre tal cual lo traía la plantilla, si ese período no tiene
        ningún mes recalculado por el programa.

        En 2017 y 2018 los cierres de la plantilla no cuadran con sus meses
        (fórmulas viejas: el Q1 de 2018, por ejemplo, se olvidó de enero).
        Recalcularlos daría un archivo "más correcto" pero distinto del que
        el área viene mirando hace años, así que para los períodos que ya
        venían cerrados se respeta el número de la plantilla.
        """
        if not self.cierres or isinstance(clave, int):
            return None
        if any(m in self.recalculados for m in self.meses_de(anio, clave)):
            return None
        return self.cierres.get((anio, clave, banda))

    def cerrado(self, banda: int, anio: int, clave) -> bool:
        """Un cierre (Q, H1, año) solo se escribe cuando el período terminó y
        no le falta ningún mes.

        Lo primero, porque un "Total 2026" con nueve meses sería engañoso (la
        plantilla real hace lo mismo: deja esas columnas en blanco hasta que
        el período termina). Lo segundo, porque si a un trimestre le falta un
        mes -- por ejemplo agosto, cuyas semanas todavía no se cargaron -- el
        Q3 saldría con dos meses y nadie lo notaría al mirarlo.
        """
        if isinstance(clave, int) or self.hasta is None:
            return True
        meses = self.meses_de(anio, clave)
        if max(meses) > self.hasta:
            return False
        dentro = [m for m in meses if m >= config.MENSUAL_INICIO]
        if not dentro:
            return True
        faltan = [m for m in dentro if (banda, m) not in self.por_mes]
        return not faltan or len(faltan) == len(dentro)

    def meses_de(self, anio: int, clave) -> list:
        if isinstance(clave, int):
            return [(anio, clave)]
        if clave == "YEAR":
            return [(anio, m) for m in range(1, 13)]
        return [(anio, m) for m in CIERRES[clave]]

    def _bruto(self, banda: int, anio: int, clave):
        """(tracks promedio, streams sumados) por sello, o None si no hay
        ningún mes con datos."""
        presentes = [self.por_mes[(banda, m)] for m in self.meses_de(anio, clave)
                     if (banda, m) in self.por_mes]
        if not presentes:
            return None
        salida = {}
        for label in config.LABEL_GROUPS_MS:
            tracks = [d[label][0] for d in presentes if label in d and d[label][0] is not None]
            streams = [d[label][1] for d in presentes if label in d and d[label][1] is not None]
            salida[label] = (sum(tracks) / len(tracks) if tracks else None,
                             sum(streams) if streams else None)
        return salida

    def celda(self, banda: int, anio: int, clave, tipo: str, label: str):
        """El valor que va en una celda, o None si ese mes/cierre no tiene
        datos todavía."""
        # El cierre heredado de la plantilla manda: ese período ya lo cerró
        # el área en su momento, con los datos que tenía.
        if not self.cerrado(banda, anio, clave) and not self.cierre_heredado(banda, anio, clave):
            return None
        datos = self._cache.get((banda, anio, clave))
        if datos is None:
            datos = self._calcular(banda, anio, clave)
            self._cache[(banda, anio, clave)] = datos
        if datos is False:
            return None
        return datos[tipo].get(label)

    def _calcular(self, banda: int, anio: int, clave):
        heredado = self.cierre_heredado(banda, anio, clave)
        if heredado:
            return {tipo: {lab: v[i] for lab, v in heredado.items()}
                    for i, tipo in enumerate(("tracks", "streams", "pct"))}
        bruto = self._bruto(banda, anio, clave)
        if bruto is None:
            return False
        tracks = {lab: v[0] for lab, v in bruto.items()}
        streams = {lab: v[1] for lab, v in bruto.items()}
        if clave == "YEAR":
            # El año, en tracks, es (H1 + Q3 + Q4) / 3 -- ver el docstring.
            tracks = {}
            for lab in config.LABEL_GROUPS_MS:
                trozos = [self.celda(banda, anio, k, "tracks", lab) for k in ("H1", "Q3", "Q4")]
                trozos = [t for t in trozos if t is not None]
                tracks[lab] = sum(trozos) / 3 if len(trozos) == 3 else (
                    sum(trozos) / len(trozos) if trozos else None)
        total = sum(v for v in streams.values() if v)
        pct = {lab: (streams[lab] / total if total and streams[lab] is not None else None)
               for lab in config.LABEL_GROUPS_MS}
        return {"tracks": tracks, "streams": streams, "pct": pct}


def _encabezado(ws, fila: int, columnas: list, col_inicio: int, con_fecha: bool) -> None:
    for i, (anio, clave) in enumerate(columnas):
        celda = ws.cell(fila, col_inicio + i)
        celda.value = etiqueta(anio, clave)
        celda.font = Font(bold=True, size=9, color=BLANCO)
        celda.fill = PatternFill("solid", start_color=AZUL, end_color=AZUL)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_inicio + i)].width = 9.6


def hoja_det(wb, pais: str, valores: Valores, columnas: list) -> None:
    ws = wb.create_sheet(config.MENSUAL_HOJA_DET.format(sigla=pais))
    ws.sheet_state = "visible"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 18
    _encabezado(ws, 1, columnas, 4, con_fecha=True)

    fila = 3
    for banda in config.MENSUAL_BANDAS:
        for tipo, titulo, formato in BLOQUES:
            celda = ws.cell(fila, 1)
            celda.value = titulo.format(banda=banda)
            celda.font = Font(bold=True, size=9)
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila + 6, end_column=1)
            for j, label in enumerate(config.LABEL_GROUPS_MS):
                ws.cell(fila + j, 2).value = config.MENSUAL_NOMBRES_DET[tipo][label]
                ws.cell(fila + j, 2).font = Font(size=9)
                for i, (anio, clave) in enumerate(columnas):
                    c = ws.cell(fila + j, 4 + i)
                    c.value = valores.celda(banda, anio, clave, tipo, label)
                    c.number_format = formato
                    c.font = Font(size=9)
            fila += 7
        fila += 1
    ws.freeze_panes = "D2"


def hoja_pais(wb, pais: str, nombre: str, valores: Valores, columnas: list) -> None:
    ws = wb.create_sheet(config.MENSUAL_HOJAS_PAIS[pais])
    ws.column_dimensions["A"].width = 9.5
    ws.column_dimensions["B"].width = 20.6
    titulo = ws.cell(2, 1)
    titulo.value = f"% Streams - {nombre}"
    titulo.font = Font(bold=True, size=11, color=BLANCO)
    titulo.fill = PatternFill("solid", start_color=AZUL, end_color=AZUL)
    _encabezado(ws, 2, columnas, 3, con_fecha=False)

    fila = 4
    for banda in config.MENSUAL_BANDAS:
        celda = ws.cell(fila, 1)
        celda.value = f"TOP\n{banda}"
        celda.font = Font(bold=True, size=10)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila + 6, end_column=1)
        for j, label in enumerate(config.MENSUAL_ORDEN_PAIS):
            ws.cell(fila + j, 2).value = config.MENSUAL_NOMBRES_PAIS[label]
            ws.cell(fila + j, 2).font = Font(size=10)
            for i, (anio, clave) in enumerate(columnas):
                c = ws.cell(fila + j, 3 + i)
                c.value = valores.celda(banda, anio, clave, "pct", label)
                c.number_format = FMT_PCT
                c.font = Font(size=10)
                c.border = _BORDE
        fila += 8
    ws.freeze_panes = "C3"


def hoja_resumen(wb, valores_por_pais: dict, columnas: list) -> None:
    ws = wb.create_sheet("Resumen", 0)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 3
    titulo = ws.cell(1, 2)
    titulo.value = "Resumen / Market Share UM"
    titulo.font = Font(bold=True, size=12, color=BLANCO)
    titulo.fill = PatternFill("solid", start_color=AZUL, end_color=AZUL)

    ws.cell(3, 2).value = "MES /TOP"
    ws.cell(3, 2).font = Font(bold=True, size=9, color=BLANCO)
    ws.cell(3, 2).fill = PatternFill("solid", start_color=AZUL, end_color=AZUL)

    col = 4
    for pais, nombre in config.MENSUAL_PAISES:
        cab = ws.cell(3, col)
        cab.value = nombre
        cab.font = Font(bold=True, size=9, color=BLANCO)
        cab.fill = PatternFill("solid", start_color=AZUL, end_color=AZUL)
        cab.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=3, start_column=col,
                       end_row=3, end_column=col + len(config.MENSUAL_BANDAS_RESUMEN) - 1)
        for k, banda in enumerate(config.MENSUAL_BANDAS_RESUMEN):
            c = ws.cell(4, col + k)
            c.value = banda
            c.font = Font(bold=True, size=9)
            c.fill = PatternFill("solid", start_color=GRIS, end_color=GRIS)
            c.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(col + k)].width = 7
        col += len(config.MENSUAL_BANDAS_RESUMEN) + 1

    for i, (anio, clave) in enumerate(columnas):
        fila = 5 + i
        etq = ws.cell(fila, 2)
        etq.value = etiqueta(anio, clave).replace("\n", " ")
        etq.font = Font(size=9, bold=not isinstance(clave, int))
        col = 4
        for pais, _ in config.MENSUAL_PAISES:
            valores = valores_por_pais.get(pais)
            for k, banda in enumerate(config.MENSUAL_BANDAS_RESUMEN):
                c = ws.cell(fila, col + k)
                c.value = valores.celda(banda, anio, clave, "pct", "Universal") if valores else None
                c.number_format = FMT_PCT
                c.font = Font(size=9)
                c.border = _BORDE
            col += len(config.MENSUAL_BANDAS_RESUMEN) + 1
    ws.freeze_panes = "D5"


def _cierres_por_pais(cierres) -> dict:
    """De la tabla de cierres heredados al diccionario que usa `Valores`."""
    salida = {}
    if cierres is None or getattr(cierres, "empty", True):
        return salida
    for fila in cierres.to_dict("records"):
        clave = (int(fila["anio"]), fila["periodo"], int(fila["banda"]))
        salida.setdefault(fila["country_code"], {}).setdefault(clave, {})[fila["label_group"]] = (
            fila["tracks"], fila["streams_millones"], fila["pct_streams"])
    return salida


def generar_reporte(mensual, path, hasta: tuple, desde: tuple = None, cierres=None):
    """Escribe el Excel completo con el histórico `mensual` (el DataFrame de
    `history.cargar_ms_mensual()`), hasta el mes `hasta`.

    `cierres` es lo que devuelve `history.cargar_ms_mensual_cierre()`: los
    Q/H1/año tal cual venían en la plantilla, que se respetan mientras el
    período no tenga ningún mes recalculado (ver `Valores.cierre_heredado`).
    """
    desde = desde or config.MENSUAL_INICIO
    columnas = calendario(desde, hasta)
    datos = mensual[(mensual["anio"] * 100 + mensual["mes"]) <= hasta[0] * 100 + hasta[1]]
    heredados = _cierres_por_pais(cierres)

    valores_por_pais = {}
    for pais, _ in config.MENSUAL_PAISES:
        valores_por_pais[pais] = Valores(
            datos[datos["country_code"] == pais].to_dict("records"), hasta=hasta,
            cierres=heredados.get(pais))

    wb = Workbook()
    wb.remove(wb.active)
    hoja_resumen(wb, valores_por_pais, columnas)
    for pais, nombre in config.MENSUAL_PAISES:
        hoja_pais(wb, pais, nombre, valores_por_pais[pais], columnas)
    for pais, _ in config.MENSUAL_PAISES:
        hoja_det(wb, pais, valores_por_pais[pais], columnas)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def nombre_reporte(anio: int, mes: int) -> str:
    return config.MENSUAL_ARCHIVO.format(mes=config.MESES_ES_ABREV[mes], anio=anio)
