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
from tkinter import filedialog, messagebox

# Al empaquetar con PyInstaller en modo "--windowed" (sin consola), sys.stdout
# / sys.stderr pueden quedar en None -- y main.py usa print() para mostrar el
# progreso. Sin este resguardo, la primera llamada a print() dentro del .exe
# empaquetado lanzaría un error y la GUI se cerraría sola sin explicación.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

from . import config
from . import main as main_module


def generar(fuente_path: str, semana: str) -> tuple[Path, Path]:
    """Corre el mismo proceso que `main.py` (carga la fuente, guarda la
    semana en el histórico si hace falta, genera los dos reportes) y
    devuelve las rutas de los reportes generados, para que la GUI pueda
    mostrarlas. Separado de la clase `App` para poder probarlo con pytest
    sin necesitar una pantalla (tkinter no se importa en los tests).
    """
    main_module.main(["--fuente", fuente_path, "--semana", semana])
    chart_out = config.OUTPUT_DIR / f"Reporte_Chart_Top_Semanal_Sem_{semana}.xlsx"
    ms_out = config.OUTPUT_DIR / f"Reporte_MS_TOP200_Sem_{semana}.xlsx"
    return chart_out, ms_out


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reportes Spotify Latam")
        self.geometry("560x240")
        self.resizable(False, False)

        self.fuente_var = tk.StringVar()
        self.semana_var = tk.StringVar()
        # Tkinter no es seguro para llamarlo desde otro hilo (el trabajo
        # pesado corre en un hilo aparte para no congelar la ventana -- ver
        # _generar). En vez de que ese hilo llame directo a self.after(...),
        # deja el resultado en esta cola thread-safe, y el hilo principal
        # (el único que toca la ventana) la revisa periódicamente con
        # self.after -- ese sí programado siempre desde el hilo principal.
        self._resultado_queue: "queue.Queue" = queue.Queue()

        tk.Label(self, text="1. Elige el archivo fuente de la semana (el Excel de BigQuery):").pack(
            anchor="w", padx=14, pady=(18, 4)
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

        self.boton_generar = tk.Button(
            self, text="3. Generar reportes", command=self._generar,
            bg="#1DB954", fg="white", font=("Segoe UI", 11, "bold"),
        )
        self.boton_generar.pack(pady=22)

        self.estado_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self.estado_var, fg="#555", wraplength=520, justify="left").pack(
            padx=14
        )

    def _elegir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Elige el archivo fuente de la semana",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.fuente_var.set(ruta)

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

        hilo = threading.Thread(target=self._generar_en_hilo, args=(fuente, semana), daemon=True)
        hilo.start()
        # Programado desde el hilo principal (el único punto donde se toca
        # Tkinter desde fuera del hilo de trabajo) -- revisa la cola cada
        # 150ms hasta que el hilo de trabajo deje un resultado.
        self.after(150, self._revisar_resultado)

    def _generar_en_hilo(self, fuente: str, semana: str) -> None:
        # Corre en un hilo aparte para que la ventana no se quede "congelada"
        # (sin responder) mientras se procesan las ~3000 filas del archivo
        # fuente -- puede tardar varios segundos. No debe llamar a NINGÚN
        # método de Tkinter directamente (ver nota en __init__): solo deja
        # el resultado en la cola.
        try:
            chart_out, ms_out = generar(fuente, semana)
        except (SystemExit, Exception) as e:  # noqa: BLE001 -- se la mostramos tal cual al usuario
            self._resultado_queue.put(("error", str(e)))
        else:
            self._resultado_queue.put(("exito", (chart_out, ms_out)))

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

    def _exito(self, chart_out: Path, ms_out: Path) -> None:
        self.boton_generar.config(state="normal")
        self.estado_var.set(f"Listo. Reportes guardados en: {chart_out.parent}")
        messagebox.showinfo(
            "Reportes generados",
            f"Se generaron los dos reportes en:\n\n{chart_out.parent}\n\n"
            f"- {chart_out.name}\n- {ms_out.name}",
        )

    def _error(self, mensaje: str) -> None:
        self.boton_generar.config(state="normal")
        self.estado_var.set("Ocurrió un error, revisa el mensaje.")
        messagebox.showerror("Error al generar los reportes", mensaje)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
