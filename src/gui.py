"""Interfaz gráfica mínima para generar los reportes semanales sin usar la
terminal ni git. Pensada para empaquetarse como un .exe con PyInstaller (ver
`build.bat` en la raíz del repo) y usarse haciendo doble clic, semana a
semana.

Uso con Python (para probar antes de empaquetar):
    python -m src.gui
"""
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# Al empaquetar con PyInstaller en modo "--windowed" (sin consola), sys.stdout
# / sys.stderr pueden quedar en None -- y main.py usa print() para mostrar el
# progreso. Sin este resguardo, la primera llamada a print() dentro del .exe
# empaquetado lanzaría un error y la GUI se cerraría sola sin explicación.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import pandas as pd

from . import bmat_calculo, bmat_proceso, chart_semanal, config, history, mensual, preferencias
from . import main as main_module


def texto_ultima_semana() -> str:
    """Una línea para la ventana con hasta dónde llega el histórico.

    Se muestra al abrir y se actualiza al terminar de generar. Nació de un
    problema real: cuando el archivo fuente traía una fecha ya cargada, la
    semana no se agregaba y los reportes salían hasta la semana anterior --
    y no había forma de darse cuenta sin abrir el Excel. Con esto se ve
    antes y después, en la misma ventana.
    """
    ultima = history.ultima_semana_cargada()
    if ultima is None:
        return "Histórico VACÍO: no hay ninguna semana cargada."
    anio, semana, chart_date = ultima
    try:
        fecha = pd.Timestamp(chart_date)
        legible = f"{fecha.day:02d} {config.MESES_ES_ABREV[fecha.month]} {fecha.year}"
    except (ValueError, TypeError):  # una fecha rara no debería tumbar la ventana
        legible = str(chart_date)
    return f"Última semana cargada: {semana} de {anio}  ({legible})"


def generar(fuente_path: str, semana: str, salida: str = None):
    """Corre el mismo proceso que `main.py` (carga la fuente, guarda la
    semana en el histórico si hace falta, genera los dos reportes) y
    devuelve las rutas de los reportes generados, para que la GUI pueda
    mostrarlas. Separado de la clase `App` para poder probarlo con pytest
    sin necesitar una pantalla (tkinter no se importa en los tests).

    Devuelve (chart_out, ms_out, aviso). `aviso` es None si todo salió
    bien, o el texto de lo que haya que advertir: que el histórico está
    vacío o incompleto (ver main.revisar_historico) y/o que no se pudieron
    resolver las fechas de lanzamiento vía Spotify. Existe porque el .exe se
    empaqueta con --windowed (sin consola): esos avisos se imprimen con
    `print` y ahí no los ve nadie, así que la ventana los tiene que mostrar
    en el mensaje final.
    """
    carpeta = Path(salida) if salida else preferencias.carpeta_salida()
    avisos = list(main_module.main(
        ["--fuente", fuente_path, "--semana", semana, "--salida", str(carpeta)]) or [])
    aviso_fechas = chart_semanal.ultimo_aviso_fechas()
    if aviso_fechas:
        avisos.append(aviso_fechas)

    preferencias.recordar_carpeta_salida(carpeta)
    chart_out = carpeta / f"Reporte_Chart_Top_Semanal_Sem_{semana}.xlsx"
    ms_out = carpeta / f"Reporte_MS_TOP200_Sem_{semana}.xlsx"
    return chart_out, ms_out, "\n\n".join(avisos) if avisos else None


def texto_ultima_semana_bmat() -> str:
    """Igual que texto_ultima_semana, para BMAT (su histórico es aparte)."""
    ultima = bmat_calculo.ultima_semana()
    if ultima is None:
        return "BMAT: todavía no hay semanas calculadas (se siembra hasta la 35 de 2026 al generar)."
    return f"BMAT - última semana cargada: {ultima[1]} de {ultima[0]}"


def generar_bmat(carpeta: str, revision: str = None, salida: str = None):
    """Corre el proceso BMAT completo (ver bmat_proceso.generar_semana) y
    devuelve (resultado, texto para mostrar). Sin tkinter, para probarlo."""
    destino = Path(salida) if salida else preferencias.carpeta_salida(preferencias.CARPETA_SALIDA_BMAT)
    r = bmat_proceso.generar_semana(carpeta, revision=revision or None, salida=destino)
    preferencias.recordar_carpeta_salida(destino, preferencias.CARPETA_SALIDA_BMAT)
    return r, bmat_proceso.resumen(r)


class VentanaBMAT(tk.Toplevel):
    """Ventana de los reportes BMAT: carpeta con los WK de la semana y, si
    ya se revisó, la lista de tracks nuevos corregida."""

    def __init__(self, master):
        super().__init__(master)
        self.title("Reportes BMAT")
        self.geometry("620x400")
        self.resizable(False, False)
        self.carpeta_var = tk.StringVar()
        self.revision_var = tk.StringVar()
        self.salida_var = tk.StringVar(
            value=str(preferencias.carpeta_salida(preferencias.CARPETA_SALIDA_BMAT)))
        self._cola: "queue.Queue" = queue.Queue()

        self.estado_hist = tk.StringVar(value=texto_ultima_semana_bmat())
        tk.Label(self, textvariable=self.estado_hist, fg="#1D7A3E", font=("Segoe UI", 9, "bold")).pack(
            anchor="w", padx=14, pady=(14, 0))

        tk.Label(self, text="1. Carpeta con los archivos WK de BMAT de la semana (WK36-CO.xlsx, ...):").pack(
            anchor="w", padx=14, pady=(14, 4))
        f1 = tk.Frame(self)
        f1.pack(fill="x", padx=14)
        tk.Entry(f1, textvariable=self.carpeta_var, width=56).pack(side="left")
        tk.Button(f1, text="Elegir carpeta...", command=self._elegir_carpeta).pack(side="left", padx=6)

        tk.Label(self, text="2. (Opcional) Lista \"Revisar clasificacion BMAT\" ya corregida:").pack(
            anchor="w", padx=14, pady=(14, 4))
        f2 = tk.Frame(self)
        f2.pack(fill="x", padx=14)
        tk.Entry(f2, textvariable=self.revision_var, width=56).pack(side="left")
        tk.Button(f2, text="Elegir archivo...", command=self._elegir_revision).pack(side="left", padx=6)

        tk.Label(self, text="3. Carpeta donde dejar los reportes:").pack(anchor="w", padx=14, pady=(14, 4))
        f3 = tk.Frame(self)
        f3.pack(fill="x", padx=14)
        tk.Entry(f3, textvariable=self.salida_var, width=56).pack(side="left")
        tk.Button(f3, text="Cambiar...", command=self._elegir_salida).pack(side="left", padx=6)

        self.boton = tk.Button(self, text="4. Generar reportes BMAT", command=self._generar,
                               bg="#002060", fg="white", font=("Segoe UI", 11, "bold"))
        self.boton.pack(pady=20)
        self.estado = tk.StringVar()
        tk.Label(self, textvariable=self.estado, fg="#555", wraplength=560, justify="left").pack(padx=14)

    def _elegir_carpeta(self):
        ruta = filedialog.askdirectory(title="Carpeta con los WK de BMAT", parent=self)
        if ruta:
            self.carpeta_var.set(ruta)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Carpeta donde dejar los reportes BMAT", parent=self,
                                       initialdir=self.salida_var.get() or None)
        if ruta:
            self.salida_var.set(ruta)

    def _elegir_revision(self):
        ruta = filedialog.askopenfilename(title="Lista para revisar corregida", parent=self,
                                          filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        if ruta:
            self.revision_var.set(ruta)

    def _generar(self):
        carpeta = self.carpeta_var.get().strip()
        if not carpeta:
            messagebox.showerror("Falta la carpeta", "Elige la carpeta con los archivos WK de BMAT de la semana.", parent=self)
            return
        self.boton.config(state="disabled")
        self.estado.set("Generando, un momento...")
        threading.Thread(target=self._en_hilo,
                         args=(carpeta, self.revision_var.get().strip(), self.salida_var.get().strip()),
                         daemon=True).start()
        self.after(200, self._revisar)

    def _en_hilo(self, carpeta, revision, salida):
        try:
            self._cola.put(("exito", generar_bmat(carpeta, revision, salida)))
        except (SystemExit, Exception) as e:  # noqa: BLE001 -- se muestra tal cual
            self._cola.put(("error", str(e)))

    def _revisar(self):
        try:
            tipo, dato = self._cola.get_nowait()
        except queue.Empty:
            self.after(200, self._revisar)
            return
        self.boton.config(state="normal")
        self.estado_hist.set(texto_ultima_semana_bmat())
        if tipo == "error":
            self.estado.set("Ocurrió un error, revisa el mensaje.")
            messagebox.showerror("Error en BMAT", dato, parent=self)
            return
        resultado, texto = dato
        self.revision_var.set("")
        self.estado.set(f"Listo. Archivos en: {resultado.carpeta_salida}")
        if resultado.avisos:
            messagebox.showwarning("Reportes BMAT generados (con avisos)", texto, parent=self)
        else:
            messagebox.showinfo("Reportes BMAT generados", texto, parent=self)


def generar_mensual(anio, mes, salida: str = None):
    """Genera el reporte mensual y devuelve (ruta, texto para mostrar, avisos).
    Sin tkinter, para poder probarlo."""
    destino = Path(salida) if salida else preferencias.carpeta_salida(
        preferencias.CARPETA_SALIDA_MENSUAL)
    ruta, avisos = mensual.generar(int(anio), int(mes), salida=destino)
    preferencias.recordar_carpeta_salida(destino, preferencias.CARPETA_SALIDA_MENSUAL)
    return ruta, mensual.resumen(ruta, avisos, int(anio), int(mes)), avisos


def texto_ultimo_mes() -> str:
    meses = mensual.meses_cargados()
    if not meses:
        return "Mensual: todavía no hay meses en el histórico."
    anio, mes = meses[-1]
    return f"Mensual - último mes en el histórico: {config.MESES_ES[mes].lower()} de {anio}"


class VentanaMensual(tk.Toplevel):
    """Ventana del reporte mensual: solo hay que elegir el mes. Los datos
    salen de las semanas de BQ que ya se cargaron."""

    def __init__(self, master):
        super().__init__(master)
        self.title("Reporte mensual - Market Share Spotify Latam")
        self.geometry("620x330")
        self.resizable(False, False)
        anio, mes = mensual.ultimo_mes_cerrable()
        self.anio_var = tk.StringVar(value=str(anio))
        self.mes_var = tk.StringVar(value=config.MESES_ES[mes].capitalize())
        self.salida_var = tk.StringVar(
            value=str(preferencias.carpeta_salida(preferencias.CARPETA_SALIDA_MENSUAL)))
        self._cola: "queue.Queue" = queue.Queue()

        self.estado_hist = tk.StringVar(value=texto_ultimo_mes())
        tk.Label(self, textvariable=self.estado_hist, fg="#1D7A3E",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(14, 0))

        tk.Label(self, text="1. Mes que quieres cerrar:").pack(anchor="w", padx=14, pady=(14, 4))
        f1 = tk.Frame(self)
        f1.pack(fill="x", padx=14)
        self.combo_mes = ttk.Combobox(
            f1, textvariable=self.mes_var, width=14, state="readonly",
            values=[config.MESES_ES[m].capitalize() for m in range(1, 13)])
        self.combo_mes.pack(side="left")
        tk.Entry(f1, textvariable=self.anio_var, width=8).pack(side="left", padx=8)

        tk.Label(self, text="2. Carpeta donde dejar el reporte:").pack(
            anchor="w", padx=14, pady=(14, 4))
        f2 = tk.Frame(self)
        f2.pack(fill="x", padx=14)
        tk.Entry(f2, textvariable=self.salida_var, width=56).pack(side="left")
        tk.Button(f2, text="Cambiar...", command=self._elegir_salida).pack(side="left", padx=6)

        self.boton = tk.Button(self, text="3. Generar reporte mensual", command=self._generar,
                               bg="#002060", fg="white", font=("Segoe UI", 11, "bold"))
        self.boton.pack(pady=18)
        self.estado = tk.StringVar()
        tk.Label(self, textvariable=self.estado, fg="#555", wraplength=560,
                 justify="left").pack(padx=14)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Carpeta donde dejar el reporte mensual", parent=self,
                                       initialdir=self.salida_var.get() or None)
        if ruta:
            self.salida_var.set(ruta)

    def _numero_de_mes(self):
        nombre = self.mes_var.get().strip().upper()
        for numero, texto in config.MESES_ES.items():
            if texto == nombre:
                return numero
        return None

    def _generar(self):
        mes = self._numero_de_mes()
        anio = self.anio_var.get().strip()
        if mes is None or not anio.isdigit():
            messagebox.showerror("Falta el mes",
                                 "Elige el mes y escribe el año (por ejemplo 2026).", parent=self)
            return
        self.boton.config(state="disabled")
        self.estado.set("Generando, un momento...")
        threading.Thread(target=self._en_hilo,
                         args=(anio, mes, self.salida_var.get().strip()),
                         daemon=True).start()
        self.after(200, self._revisar)

    def _en_hilo(self, anio, mes, salida):
        try:
            self._cola.put(("exito", generar_mensual(anio, mes, salida)))
        except (SystemExit, Exception) as e:  # noqa: BLE001 -- se muestra tal cual
            self._cola.put(("error", str(e)))

    def _revisar(self):
        try:
            tipo, dato = self._cola.get_nowait()
        except queue.Empty:
            self.after(200, self._revisar)
            return
        self.boton.config(state="normal")
        self.estado_hist.set(texto_ultimo_mes())
        if tipo == "error":
            self.estado.set("Ocurrió un error, revisa el mensaje.")
            messagebox.showerror("Error en el reporte mensual", dato, parent=self)
            return
        ruta, texto, avisos = dato
        self.estado.set(f"Listo. Archivo en: {ruta}")
        if avisos:
            messagebox.showwarning("Reporte mensual generado (con avisos)", texto, parent=self)
        else:
            messagebox.showinfo("Reporte mensual generado", texto, parent=self)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reportes Spotify Latam")
        self.geometry("560x400")
        self.resizable(False, False)

        self.fuente_var = tk.StringVar()
        self.semana_var = tk.StringVar()
        self.salida_var = tk.StringVar(value=str(preferencias.carpeta_salida()))
        # Tkinter no es seguro para llamarlo desde otro hilo (el trabajo
        # pesado corre en un hilo aparte para no congelar la ventana -- ver
        # _generar). En vez de que ese hilo llame directo a self.after(...),
        # deja el resultado en esta cola thread-safe, y el hilo principal
        # (el único que toca la ventana) la revisa periódicamente con
        # self.after -- ese sí programado siempre desde el hilo principal.
        self._resultado_queue: "queue.Queue" = queue.Queue()

        # Estado del histórico, arriba de todo: es el contexto con el que uno
        # decide qué archivo cargar (ver texto_ultima_semana).
        self.historico_var = tk.StringVar()
        self.label_historico = tk.Label(
            self, textvariable=self.historico_var, anchor="w", justify="left",
            font=("Segoe UI", 9, "bold"),
        )
        self.label_historico.pack(anchor="w", padx=14, pady=(14, 0))
        self._refrescar_historico()

        tk.Label(self, text="1. Elige el archivo fuente de la semana (el Excel de BigQuery):").pack(
            anchor="w", padx=14, pady=(14, 4)
        )
        frame_fuente = tk.Frame(self)
        frame_fuente.pack(fill="x", padx=14)
        tk.Entry(frame_fuente, textvariable=self.fuente_var, width=56).pack(side="left")
        tk.Button(frame_fuente, text="Elegir archivo...", command=self._elegir_archivo).pack(
            side="left", padx=6
        )

        tk.Label(self, text="2. Número de semana (solo para nombrar los archivos de salida, ej. 25):").pack(
            anchor="w", padx=14, pady=(18, 4)
        )
        tk.Entry(self, textvariable=self.semana_var, width=10).pack(anchor="w", padx=14)

        tk.Label(self, text="3. Carpeta donde dejar los reportes:").pack(
            anchor="w", padx=14, pady=(18, 4)
        )
        frame_salida = tk.Frame(self)
        frame_salida.pack(fill="x", padx=14)
        tk.Entry(frame_salida, textvariable=self.salida_var, width=56).pack(side="left")
        tk.Button(frame_salida, text="Cambiar...", command=self._elegir_salida).pack(side="left", padx=6)

        self.boton_generar = tk.Button(
            self, text="4. Generar reportes", command=self._generar,
            bg="#1DB954", fg="white", font=("Segoe UI", 11, "bold"),
        )
        self.boton_generar.pack(pady=22)

        self.estado_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self.estado_var, fg="#555", wraplength=520, justify="left").pack(
            padx=14
        )

        # BMAT y el mensual son otros procesos (otras fuentes, otros
        # reportes): cada uno con su propia ventana.
        tk.Button(self, text="Reportes BMAT...", command=lambda: VentanaBMAT(self)).place(
            relx=1.0, rely=1.0, x=-12, y=-10, anchor="se")
        tk.Button(self, text="Reporte mensual...",
                  command=lambda: VentanaMensual(self)).place(
            relx=1.0, rely=1.0, x=-140, y=-10, anchor="se")

    def _refrescar_historico(self) -> None:
        texto = texto_ultima_semana()
        self.historico_var.set(texto)
        # En rojo cuando no hay nada cargado: ahí los reportes saldrían con
        # una sola fila de historia.
        self.label_historico.config(fg="#C00000" if "VACÍO" in texto else "#1D7A3E")

    def _elegir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Elige el archivo fuente de la semana",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.fuente_var.set(ruta)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Carpeta donde dejar los reportes",
                                       initialdir=self.salida_var.get() or None)
        if ruta:
            self.salida_var.set(ruta)

    def _generar(self):
        fuente = self.fuente_var.get().strip()
        semana = self.semana_var.get().strip()
        if not fuente:
            messagebox.showerror("Falta el archivo", "Elige primero el archivo fuente de la semana.")
            return
        if not semana:
            messagebox.showerror("Falta la semana", "Escribe el número de semana (ej. 25).")
            return

        self.boton_generar.config(state="disabled")
        self.estado_var.set("Generando reportes, un momento...")

        hilo = threading.Thread(target=self._generar_en_hilo,
                                args=(fuente, semana, self.salida_var.get().strip()), daemon=True)
        hilo.start()
        # Programado desde el hilo principal (el único punto donde se toca
        # Tkinter desde fuera del hilo de trabajo) -- revisa la cola cada
        # 150ms hasta que el hilo de trabajo deje un resultado.
        self.after(150, self._revisar_resultado)

    def _generar_en_hilo(self, fuente: str, semana: str, salida: str = None) -> None:
        # Corre en un hilo aparte para que la ventana no se quede "congelada"
        # (sin responder) mientras se procesan las ~3000 filas del archivo
        # fuente -- puede tardar varios segundos. No debe llamar a NINGÚN
        # método de Tkinter directamente (ver nota en __init__): solo deja
        # el resultado en la cola.
        try:
            chart_out, ms_out, aviso = generar(fuente, semana, salida)
        except (SystemExit, Exception) as e:  # noqa: BLE001 -- se la mostramos tal cual al usuario
            self._resultado_queue.put(("error", str(e)))
        else:
            self._resultado_queue.put(("exito", (chart_out, ms_out, aviso)))

    def _revisar_resultado(self) -> None:
        try:
            tipo, dato = self._resultado_queue.get_nowait()
        except queue.Empty:
            self.after(150, self._revisar_resultado)
            return

        if tipo == "exito":
            self._exito(*dato)
        else:
            self._error(dato)

    def _exito(self, chart_out: Path, ms_out: Path, aviso: str = None) -> None:
        self.boton_generar.config(state="normal")
        # Después de generar, el histórico cambió: mostrar el estado nuevo.
        self._refrescar_historico()
        texto = (
            f"Se generaron los dos reportes en:\n\n{chart_out.parent}\n\n"
            f"- {chart_out.name}\n- {ms_out.name}"
        )
        if aviso:
            # Sin consola (--windowed) este es el único lugar donde el
            # usuario puede enterarse de que algo quedó a medias.
            self.estado_var.set("Listo, pero con un aviso. Revisa el mensaje.")
            messagebox.showwarning("Reportes generados (con un aviso)", f"{texto}\n\n⚠ {aviso}")
            return
        self.estado_var.set(f"Listo. Reportes guardados en: {chart_out.parent}")
        messagebox.showinfo("Reportes generados", texto)

    def _error(self, mensaje: str) -> None:
        self.boton_generar.config(state="normal")
        self.estado_var.set("Ocurrió un error, revisa el mensaje.")
        messagebox.showerror("Error al generar los reportes", mensaje)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()