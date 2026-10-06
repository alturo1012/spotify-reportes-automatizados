"""Pruebas del reporte mensual: cómo se arma un mes con las semanas y cómo
sale el Excel."""
import openpyxl
import pandas as pd
import pytest

from src import config, history, mensual, mensual_reporte


SELLOS = config.LABEL_GROUPS_MS


def semana(fecha, country="CO", streams=None, tracks=None, banda=200):
    """Una semana suelta de `ms_band_label_weekly`, con un valor por sello."""
    streams = streams or {"Universal": 10.0, "Sony": 30.0}
    tracks = tracks or {"Universal": 4.0, "Sony": 6.0}
    total = sum(streams.get(s, 0.0) for s in SELLOS)
    return pd.DataFrame([
        {"anio": pd.Timestamp(fecha).year, "semana": pd.Timestamp(fecha).isocalendar()[1],
         "chart_date": fecha, "country_code": country, "banda": banda, "label_group": s,
         "pct_streams": streams.get(s, 0.0) / total if total else 0.0,
         "tracks": tracks.get(s, 0.0), "streams_millones": streams.get(s, 0.0)}
        for s in SELLOS
    ])


def cuatro_semanas_de_septiembre(country="CO"):
    return pd.concat([
        semana("2026-09-03", country, {"Universal": 10.0, "Sony": 30.0}, {"Universal": 4.0, "Sony": 6.0}),
        semana("2026-09-10", country, {"Universal": 20.0, "Sony": 20.0}, {"Universal": 5.0, "Sony": 5.0}),
        semana("2026-09-17", country, {"Universal": 30.0, "Sony": 10.0}, {"Universal": 6.0, "Sony": 4.0}),
        semana("2026-09-24", country, {"Universal": 40.0, "Sony": 40.0}, {"Universal": 5.0, "Sony": 5.0}),
    ], ignore_index=True)


class TestCalculoDelMes:
    def test_streams_se_suman_tracks_se_promedian(self):
        df, avisos = mensual.calcular_mes(2026, 9, cuatro_semanas_de_septiembre())
        assert avisos == []
        universal = df[df["label_group"] == "Universal"].iloc[0]
        assert universal["streams_millones"] == pytest.approx(100.0)   # 10+20+30+40
        assert universal["tracks"] == pytest.approx(5.0)               # (4+5+6+5)/4
        assert universal["semanas"] == 4

    def test_el_pct_se_recalcula_sobre_el_total_del_mes(self):
        """No es el promedio de los % semanales: es streams del sello sobre
        streams de todos. Con estos números se nota la diferencia (el
        promedio de los % daría 0,4375)."""
        df, _ = mensual.calcular_mes(2026, 9, cuatro_semanas_de_septiembre())
        universal = df[df["label_group"] == "Universal"].iloc[0]
        assert universal["pct_streams"] == pytest.approx(100.0 / 200.0)

    def test_la_semana_cuenta_para_el_mes_de_su_fecha_de_corte(self):
        """La semana del 3 de septiembre empieza en agosto, pero cuenta para
        septiembre."""
        bandas = pd.concat([semana("2026-08-27"), semana("2026-09-03")], ignore_index=True)
        agosto, _ = mensual.calcular_mes(2026, 8, bandas)
        septiembre, _ = mensual.calcular_mes(2026, 9, bandas)
        assert agosto["semanas"].max() == 1
        assert septiembre["semanas"].max() == 1

    def test_avisa_cuando_al_mes_le_falta_una_semana(self):
        bandas = cuatro_semanas_de_septiembre()
        bandas = bandas[bandas["chart_date"] != "2026-09-17"]
        df, avisos = mensual.calcular_mes(2026, 9, bandas)
        assert any("falta" in a for a in avisos)
        assert df["semanas"].max() == 3

    def test_avisa_cuando_el_mes_no_tiene_ninguna_semana(self):
        df, avisos = mensual.calcular_mes(2026, 11, cuatro_semanas_de_septiembre())
        assert df.empty
        assert any("No hay ninguna semana" in a for a in avisos)

    def test_las_semanas_sembradas_sin_streams_no_cuentan(self):
        """El histórico viejo solo guardaba el %. Esas semanas no se pueden
        sumar, así que quedan fuera y se avisa."""
        bandas = cuatro_semanas_de_septiembre()
        bandas.loc[bandas["chart_date"] == "2026-09-10", ["tracks", "streams_millones"]] = None
        df, avisos = mensual.calcular_mes(2026, 9, bandas)
        assert any("sin streams ni tracks" in a for a in avisos)
        assert df["semanas"].max() == 3

    def test_semanas_esperadas_cuenta_los_cortes_del_mes(self):
        # Septiembre de 2026 tiene 4 jueves; julio, 5.
        assert mensual.semanas_esperadas(2026, 9, "2026-09-03") == 4
        assert mensual.semanas_esperadas(2026, 7, "2026-09-03") == 5


class TestGuardadoDelMes:
    def test_un_mes_completo_se_guarda(self):
        df, avisos, guardado = mensual.cerrar_mes(2026, 9, cuatro_semanas_de_septiembre())
        assert guardado and not df.empty
        guardadas = history.cargar_ms_mensual()
        assert len(guardadas[(guardadas["anio"] == 2026) & (guardadas["mes"] == 9)]) == len(df)

    def test_un_mes_incompleto_NO_se_guarda(self):
        """Si se guardara, sus streams bajos se arrastrarían al trimestre y
        al año sin que se note."""
        bandas = cuatro_semanas_de_septiembre()
        bandas = bandas[bandas["chart_date"] != "2026-09-17"]
        _, avisos, guardado = mensual.cerrar_mes(2026, 9, bandas)
        assert not guardado
        assert any("falta" in a for a in avisos)
        assert history.cargar_ms_mensual().empty

    def test_volver_a_cerrar_el_mes_lo_reemplaza(self):
        mensual.cerrar_mes(2026, 9, cuatro_semanas_de_septiembre())
        dobles = pd.concat([cuatro_semanas_de_septiembre()] * 1, ignore_index=True)
        dobles["streams_millones"] = dobles["streams_millones"] * 2
        df, _, guardado = mensual.cerrar_mes(2026, 9, dobles)
        assert guardado
        guardadas = history.cargar_ms_mensual()
        universal = guardadas[(guardadas["mes"] == 9) & (guardadas["label_group"] == "Universal")]
        assert len(universal) == 1
        assert universal.iloc[0]["streams_millones"] == pytest.approx(200.0)


class TestCalendario:
    def test_el_primer_anio_arranca_en_el_mes_de_inicio(self):
        cols = mensual_reporte.calendario((2017, 5), (2017, 12))
        assert cols[0] == (2017, 5)
        assert (2017, 1) not in cols and (2017, "Q1") not in cols
        assert (2017, "Q2") in cols and (2017, "YEAR") in cols

    def test_el_ultimo_anio_se_dibuja_completo(self):
        cols = mensual_reporte.calendario((2026, 1), (2026, 9))
        assert (2026, 12) in cols and (2026, "Q4") in cols


class TestExcel:
    def _historico(self):
        filas = []
        for (anio, mes) in [(2026, 7), (2026, 8), (2026, 9)]:
            for banda in config.MENSUAL_BANDAS:
                for i, sello in enumerate(SELLOS):
                    filas.append({"anio": anio, "mes": mes, "country_code": "CO", "banda": banda,
                                  "label_group": sello, "tracks": float(i), "origen": "calculado",
                                  "streams_millones": float(10 * (i + 1)), "pct_streams": 0.0})
        return pd.DataFrame(filas)

    def test_escribe_las_hojas_esperadas(self, tmp_path):
        ruta = mensual_reporte.generar_reporte(
            self._historico(), tmp_path / "mensual.xlsx", hasta=(2026, 9))
        wb = openpyxl.load_workbook(ruta)
        assert wb.sheetnames[0] == "Resumen"
        assert "CO" in wb.sheetnames and "CO-Det" in wb.sheetnames
        # Chile es "CH" aunque su código sea CL.
        assert "CH" in wb.sheetnames and "CL-Det" in wb.sheetnames
        assert len(wb.sheetnames) == 1 + 2 * len(config.MENSUAL_PAISES)

    def test_el_trimestre_queda_en_blanco_si_le_falta_un_mes(self, tmp_path):
        historico = self._historico()
        historico = historico[historico["mes"] != 8]
        ruta = mensual_reporte.generar_reporte(
            historico, tmp_path / "mensual.xlsx", hasta=(2026, 9))
        wb = openpyxl.load_workbook(ruta)
        cols = mensual_reporte.calendario(config.MENSUAL_INICIO, (2026, 9))
        columna = 4 + cols.index((2026, "Q3"))
        assert wb["CO-Det"].cell(98, columna).value is None

    def test_el_trimestre_completo_suma_los_streams(self, tmp_path):
        ruta = mensual_reporte.generar_reporte(
            self._historico(), tmp_path / "mensual.xlsx", hasta=(2026, 9))
        wb = openpyxl.load_workbook(ruta)
        cols = mensual_reporte.calendario(config.MENSUAL_INICIO, (2026, 9))
        columna = 4 + cols.index((2026, "Q3"))
        # Universal: 10 por mes, tres meses del Q3.
        assert wb["CO-Det"].cell(98, columna).value == pytest.approx(30.0)

    def test_el_anio_en_tracks_es_h1_q3_q4_sobre_tres(self, tmp_path):
        """Regla rara pero real de la plantilla: el año, en tracks, no es el
        promedio de los 12 meses."""
        valores = mensual_reporte.Valores(
            [{"anio": 2026, "mes": m, "banda": 200, "label_group": "Universal",
              "tracks": float(m), "streams_millones": 1.0, "origen": "calculado"}
             for m in range(1, 13)], hasta=(2026, 12))
        h1 = valores.celda(200, 2026, "H1", "tracks", "Universal")
        q3 = valores.celda(200, 2026, "Q3", "tracks", "Universal")
        q4 = valores.celda(200, 2026, "Q4", "tracks", "Universal")
        anio = valores.celda(200, 2026, "YEAR", "tracks", "Universal")
        assert anio == pytest.approx((h1 + q3 + q4) / 3)
        assert anio != pytest.approx(sum(range(1, 13)) / 12)

    def test_el_cierre_heredado_de_la_plantilla_manda(self, tmp_path):
        """Mientras el período no tenga ningún mes recalculado, se respeta el
        número que traía la plantilla aunque no cuadre con sus meses."""
        historico = self._historico()
        historico["origen"] = "plantilla"
        cierres = pd.DataFrame([
            {"anio": 2026, "periodo": "Q3", "country_code": "CO", "banda": 200,
             "label_group": s, "tracks": 1.0, "streams_millones": 999.0, "pct_streams": 0.5}
            for s in SELLOS])
        ruta = mensual_reporte.generar_reporte(
            historico, tmp_path / "mensual.xlsx", hasta=(2026, 9), cierres=cierres)
        wb = openpyxl.load_workbook(ruta)
        cols = mensual_reporte.calendario(config.MENSUAL_INICIO, (2026, 9))
        columna = 4 + cols.index((2026, "Q3"))
        assert wb["CO-Det"].cell(98, columna).value == pytest.approx(999.0)

    def test_un_mes_recalculado_pisa_el_cierre_heredado(self):
        filas = [{"anio": 2026, "mes": m, "banda": 200, "label_group": "Universal",
                  "tracks": 1.0, "streams_millones": 10.0,
                  "origen": "plantilla" if m != 8 else "calculado"} for m in (7, 8, 9)]
        cierres = {(2026, "Q3", 200): {"Universal": (1.0, 999.0, 0.5)}}
        valores = mensual_reporte.Valores(filas, hasta=(2026, 9), cierres=cierres)
        assert valores.celda(200, 2026, "Q3", "streams", "Universal") == pytest.approx(30.0)


class TestHistorico:
    def test_las_semanas_nuevas_guardan_tracks_y_streams(self):
        """El reporte mensual necesita los dos números, no solo el %."""
        df = pd.DataFrame([
            {"country_code": "CO", "position": p, "label_group": "Universal",
             "stream_count": 1_000_000.0, "chart_date": pd.Timestamp("2026-09-03"),
             "album_copyright": "© 2026 Universal"}
            for p in range(1, 11)
        ])
        history.append_semana_ms_bandas(df)
        guardado = history.cargar_ms_band_label_weekly()
        fila = guardado[(guardado["banda"] == 10) & (guardado["label_group"] == "Universal")].iloc[0]
        assert fila["tracks"] == 10
        assert fila["streams_millones"] == pytest.approx(10.0)


class TestCompletarSemanaSembrada:
    """Las semanas sembradas traen solo el %. Cuando aparece su fuente de BQ
    hay que rellenarles tracks y streams SIN cambiarles el número de semana
    (`recargar_semanas` no sirve: salta las fechas que ya vienen sembradas)."""

    def _sembrar_semana_sin_numeros(self, chart_date="2026-08-13", anio=2026, semana=33):
        conn = history.conectar()
        conn.executemany(
            "INSERT INTO ms_band_label_weekly (anio, semana, chart_date, country_code, "
            "banda, label_group, pct_streams) VALUES (?, ?, ?, ?, ?, ?, ?)",
            [(anio, semana, chart_date, "CO", banda, sello, 0.0)
             for banda in config.MENSUAL_BANDAS for sello in SELLOS])
        conn.commit()
        conn.close()

    def _fuente(self, chart_date="2026-08-13"):
        return pd.DataFrame([
            {"country_code": "CO", "position": p, "label_group": "Universal",
             "stream_count": 1_000_000.0, "chart_date": pd.Timestamp(chart_date),
             "album_copyright": "© 2026 Universal"}
            for p in range(1, 11)])

    def test_rellena_conservando_el_numero_de_semana(self):
        self._sembrar_semana_sin_numeros()
        resultado = history.completar_semana_ms_bandas(self._fuente())
        assert resultado is not None
        anio, semana, filas = resultado
        assert (anio, semana) == (2026, 33)   # NO la renumera
        assert filas == len(config.MENSUAL_BANDAS) * len(SELLOS)
        guardado = history.cargar_ms_band_label_weekly()
        fila = guardado[(guardado["banda"] == 10) & (guardado["label_group"] == "Universal")].iloc[0]
        assert fila["tracks"] == 10
        assert fila["streams_millones"] == pytest.approx(10.0)

    def test_no_agrega_semanas_nuevas(self):
        self._sembrar_semana_sin_numeros()
        antes = history.cargar_ms_band_label_weekly()
        history.completar_semana_ms_bandas(self._fuente())
        despues = history.cargar_ms_band_label_weekly()
        assert len(antes) == len(despues)

    def test_devuelve_none_si_la_fecha_no_esta_en_el_historico(self):
        self._sembrar_semana_sin_numeros()
        assert history.completar_semana_ms_bandas(self._fuente("2026-10-01")) is None

    def test_despues_de_completarla_el_mes_ya_se_puede_cerrar(self):
        """Antes de completar, la semana no cuenta y el mes avisa; después, sí."""
        self._sembrar_semana_sin_numeros()
        _, avisos_antes = mensual.calcular_mes(2026, 8)
        assert any("sin streams ni tracks" in a for a in avisos_antes)
        history.completar_semana_ms_bandas(self._fuente())
        df, avisos = mensual.calcular_mes(2026, 8)
        assert not any("sin streams ni tracks" in a for a in avisos)
        assert not df.empty


class TestSemaforo:
    """Los colores de la plantilla. Umbrales tomados del formato condicional
    real (celdas BG2/BH2/BI2/BJ2/BL2 de su hoja "Resumen")."""

    def test_el_resumen_usa_el_objetivo_del_30_por_ciento(self):
        rojo = config.MENSUAL_COLOR_ROJO
        amarillo = config.MENSUAL_COLOR_AMARILLO
        verde = config.MENSUAL_COLOR_VERDE
        celeste = config.MENSUAL_COLOR_DESTACADO
        casos = [
            (0.29, rojo), (0.294, rojo),         # 29% o menos
            (0.295, amarillo), (0.30, amarillo), (0.3049, amarillo),   # 30%
            (0.305, verde), (0.33, verde), (0.39, verde),              # 31% o más
            (0.40, celeste), (0.51, celeste),                          # 40% o más
        ]
        for valor, esperado in casos:
            assert mensual_reporte.color_universal(valor) == esperado, valor

    def test_sin_valor_no_hay_color(self):
        assert mensual_reporte.color_universal(None) is None

    def test_universal_va_en_verde_cuando_es_el_sello_mas_alto(self):
        colores = mensual_reporte.colores_del_bloque(
            {"Universal": 0.40, "Sony": 0.30, "Warner": 0.20})
        assert colores == {"Universal": config.MENSUAL_COLOR_VERDE}

    def test_los_sellos_que_le_ganan_a_universal_van_en_rojo(self):
        colores = mensual_reporte.colores_del_bloque(
            {"Universal": 0.20, "Sony": 0.45, "Warner": 0.25, "Indies": 0.10})
        assert colores == {"Sony": config.MENSUAL_COLOR_ROJO,
                           "Warner": config.MENSUAL_COLOR_ROJO}
        # Universal no se pinta de verde si no es el más alto.
        assert "Universal" not in colores

    def test_empatar_no_cuenta_como_ganarle(self):
        colores = mensual_reporte.colores_del_bloque({"Universal": 0.30, "Sony": 0.30})
        assert colores == {"Universal": config.MENSUAL_COLOR_VERDE}

    def test_el_excel_sale_pintado(self, tmp_path):
        filas = []
        for banda in config.MENSUAL_BANDAS:
            for sello, pct in [("Universal", 0.2), ("Sony", 0.5), ("INgrooves", 0.075),
                               ("Virgin", 0.075), ("Orchard", 0.05), ("Warner", 0.05),
                               ("Indies", 0.05)]:
                filas.append({"anio": 2026, "mes": 9, "country_code": "CO", "banda": banda,
                              "label_group": sello, "tracks": 1.0, "origen": "calculado",
                              "streams_millones": pct * 100, "pct_streams": pct})
        ruta = mensual_reporte.generar_reporte(
            pd.DataFrame(filas), tmp_path / "mensual.xlsx", hasta=(2026, 9))
        wb = openpyxl.load_workbook(ruta)
        cols = mensual_reporte.calendario(config.MENSUAL_INICIO, (2026, 9))
        columna = 3 + cols.index((2026, 9))
        co = wb["CO"]
        # Fila 4 = Universal, fila 5 = Sony (ver config.MENSUAL_ORDEN_PAIS).
        assert co.cell(5, columna).fill.start_color.rgb[-6:] == config.MENSUAL_COLOR_ROJO[0]
        # Universal con 20% en el Resumen: rojo (por debajo del objetivo).
        resumen = wb["Resumen"]
        columna_res = 4 + cols.index((2026, 9))
        celda = resumen.cell(5 + cols.index((2026, 9)), 4)
        assert celda.fill.start_color.rgb[-6:] == config.MENSUAL_COLOR_ROJO[0]
