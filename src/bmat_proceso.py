"""El proceso semanal BMAT completo, de la carpeta con los WK a los reportes.

    resultado = bmat_proceso.generar_semana("C:/.../BMAT semana 36")

1. Siembra (solo la primera vez) el histórico y la clasificación de la
   semana 35.
2. Si se pasa una "Lista para revisar" corregida, guarda las correcciones.
3. Lee cada "WK<semana>-<mercado>.xlsx" de la carpeta, clasifica sus tracks,
   calcula las bandas y las guarda en el histórico (si la semana ya estaba,
   la reemplaza: correr dos veces da lo mismo).
4. Escribe, en data/output/BMAT Sem S de A/ (hoy solo Colombia, ver
   config.BMAT_REPORTES_ACTIVOS):
   - el reporte final "MS BMAT COL a Sem S de A.xlsx";
   - el intermedio "Top 10000 BMAT COL a sem S de A.xlsx" (Streams Catalogo,
     Resumen, TOP 200 Nuevos y Tracks Independientes);
   - "Revisar clasificacion BMAT sem S de A.xlsx": los tracks nuevos.

Para corregir la clasificación: llenar la columna "Disqueras corregida" en
la lista y volver a correr la MISMA carpeta pasando esa lista. La semana se
recalcula con las correcciones y quedan guardadas para las siguientes.
"""
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from . import bmat, bmat_calculo, config
from . import bmat_clasificacion as clasif


@dataclass
class ResultadoBMAT:
    anio: int
    semana: int
    carpeta_salida: Path
    reportes: list = field(default_factory=list)
    intermedios: list = field(default_factory=list)
    lista_revision: Path = None
    tracks_nuevos: int = 0
    correcciones: int = 0
    mercados: list = field(default_factory=list)
    avisos: list = field(default_factory=list)


def etiqueta_semana(anio: int, semana: int) -> str:
    return f"semana {semana} de {anio}"


def generar_semana(carpeta, anio: int = None, revision=None, salida=None,
                   hoy: date = None) -> ResultadoBMAT:
    bmat_calculo.asegurar_semilla()
    clasif.sembrar()

    correcciones = clasif.aplicar_revision(revision) if revision else 0

    fuentes = bmat_calculo.buscar_fuentes(carpeta)
    semana = bmat_calculo.identificar_archivo(next(iter(fuentes.values())))[0]
    anio = anio or bmat_calculo.anio_de_la_semana(semana, hoy)
    etiqueta = etiqueta_semana(anio, semana)
    salida = Path(salida or config.OUTPUT_DIR) / f"BMAT Sem {semana} de {anio}"
    salida.mkdir(parents=True, exist_ok=True)

    resultado = ResultadoBMAT(anio=anio, semana=semana, carpeta_salida=salida,
                              correcciones=correcciones)

    # Todos los mercados se clasifican contra la MISMA foto de la tabla: un
    # track nuevo en Colombia y en Perú sale igual en los dos.
    tabla = clasif.cargar()
    titularidad = bmat_calculo.cargar_titularidad()
    clasificados = {}
    conflictos = {}   # tracks compartidos cuya clasificación no es ninguno de sus dos sellos
    for mercado in config.bmat_mercados_activos():
        if mercado not in fuentes:
            continue
        df = bmat_calculo.leer_fuente(fuentes[mercado])
        tope = config.BMAT_MERCADOS[mercado]["top"]
        if len(df) < tope:
            resultado.avisos.append(
                f"{fuentes[mercado].name} trae {len(df)} tracks (se esperaban {tope}): "
                f"las bandas más grandes salen NA."
            )
        previa = bmat_calculo.ultima_semana(mercado)
        if previa and semana > 1 and previa < (anio, semana - 1):
            resultado.avisos.append(
                f"{config.BMAT_MERCADOS[mercado]['nombre']}: la última semana guardada era la "
                f"{previa[1]} de {previa[0]}; las semanas del medio quedan NA."
            )
        c = bmat_calculo.marcar_titularidad(clasif.clasificar(df, tabla, etiqueta), titularidad)
        clasificados[mercado] = c
        for _, r in c[c["_reparto"].notna()].iterrows():
            if r["Sello"] not in {s for s, _ in r["_reparto"]}:
                clave = (r.get("Track"), r["Sello"], r["Titularidad compartida"])
                conflictos.setdefault(clave, []).append(config.BMAT_MERCADOS[mercado]["sigla"])
        bandas = bmat_calculo.calcular_bandas(c, mercado)
        bmat_calculo.guardar_semana(anio, semana, mercado, bandas)
        resultado.intermedios.append(bmat_calculo.escribir_intermedio(
            salida / bmat_calculo.nombre_intermedio(mercado, anio, semana), mercado, anio, semana, c, bandas))
        resultado.mercados.append(mercado)

    for (track, sello, reparto), siglas in conflictos.items():
        resultado.avisos.append(
            f"\"{track}\" está clasificado {sello}, pero la lista de titularidad compartida lo "
            f"reparte {reparto} ({', '.join(siglas)}). Se usó la lista; revisa cuál está bien.")

    for c in clasificados.values():
        clasif.guardar_nuevos(c, etiqueta)

    lista = clasif.lista_para_revisar(clasificados)
    resultado.tracks_nuevos = len(lista)
    resultado.lista_revision = clasif.escribir_lista_revision(
        lista, salida / f"Revisar clasificacion BMAT sem {semana} de {anio}.xlsx",
        f"Tracks nuevos por revisar - BMAT semana {semana} de {anio}")

    for reporte in config.BMAT_REPORTES_ACTIVOS:
        conf = config.BMAT_REPORTES[reporte]
        faltan = [config.BMAT_MERCADOS[m]["nombre"] for m in conf["hojas"] if m not in fuentes]
        if len(faltan) == len(conf["hojas"]):
            resultado.avisos.append(
                f"No se generó {conf['archivo']}: no había archivo WK{semana} para "
                f"{', '.join(faltan)}.")
            continue
        if faltan:
            resultado.avisos.append(
                f"{conf['archivo']}: faltó el WK{semana} de {', '.join(faltan)} "
                f"(esa semana sale NA en sus hojas).")
        resultado.reportes.append(bmat.generar_reporte(
            reporte, salida / bmat.nombre_reporte(reporte, anio, semana), hasta=(anio, semana)))
    return resultado


def resumen(r: ResultadoBMAT) -> str:
    lineas = [f"BMAT semana {r.semana} de {r.anio} "
              f"({', '.join(config.BMAT_MERCADOS[m]['nombre'] for m in r.mercados)})."]
    if r.correcciones:
        lineas.append(f"Se aplicaron {r.correcciones} correcciones de clasificación.")
    lineas.append(f"Tracks nuevos para revisar: {r.tracks_nuevos} -> {r.lista_revision.name}")
    lineas.append(f"Reportes en: {r.carpeta_salida}")
    lineas += [f"  - {p.name}" for p in r.reportes + r.intermedios]
    if r.avisos:
        lineas.append("Avisos:")
        lineas += [f"  ! {a}" for a in r.avisos]
    return "\n".join(lineas)
