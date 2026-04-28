import os
import sys
import subprocess
import threading
import customtkinter as ctk

from tkinter import filedialog, messagebox

from core.bitacora_excel.procesador import (
    generar_bitacora,
    previsualizar_bitacora,
)


def abrir_archivo(ruta):
    try:
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.call(["open", ruta])
        else:
            subprocess.call(["xdg-open", ruta])
    except Exception as e:
        messagebox.showerror(
            "Error al abrir archivo",
            f"El archivo se generó, pero no se pudo abrir automáticamente.\n\n{e}"
        )


class BitacoraView(ctk.CTkFrame):
    def __init__(self, master, volver_callback):
        super().__init__(master)

        self.volver_callback = volver_callback
        self.archivo = None

        self.grid_columnconfigure(0, weight=1)

        self._crear_widgets()

    def _crear_widgets(self):
        self.titulo = ctk.CTkLabel(
            self,
            text="Procesador de Bitácora de Audiencias",
            font=("Arial", 20, "bold"),
        )
        self.titulo.grid(row=0, column=0, pady=(20, 10), padx=20, sticky="n")

        self.btn_seleccionar = ctk.CTkButton(
            self,
            text="Seleccionar archivo Excel",
            command=self.seleccionar_archivo,
        )
        self.btn_seleccionar.grid(row=1, column=0, pady=10, padx=20)

        self.label_archivo = ctk.CTkLabel(
            self,
            text="Ningún archivo seleccionado",
            wraplength=700,
            justify="center",
        )
        self.label_archivo.grid(row=2, column=0, pady=5, padx=20)

        self.btn_preview = ctk.CTkButton(
            self,
            text="Previsualizar",
            command=self.preview,
        )
        self.btn_preview.grid(row=3, column=0, pady=10, padx=20)

        self.btn_generar = ctk.CTkButton(
            self,
            text="Generar bitácora",
            command=self.generar_en_hilo,
        )
        self.btn_generar.grid(row=4, column=0, pady=10, padx=20)

        self.status_label = ctk.CTkLabel(
            self,
            text="Estado: esperando acción...",
        )
        self.status_label.grid(row=5, column=0, pady=10, padx=20)

        self.btn_volver = ctk.CTkButton(
            self,
            text="Volver",
            command=self.volver_callback,
        )
        self.btn_volver.grid(row=6, column=0, pady=(10, 20), padx=20)

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de bitácora",
            filetypes=[("Archivos Excel", "*.xls *.xlsx")],
        )

        if not ruta:
            return

        self.archivo = ruta
        self.label_archivo.configure(text=ruta)
        self.status_label.configure(text="Estado: archivo cargado")

    def preview(self):
        if not self.archivo:
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        self._set_estado_cargando(True, "Estado: generando previsualización...")

        hilo = threading.Thread(target=self._preview_worker, daemon=True)
        hilo.start()

    def _preview_worker(self):
        try:
            datos = previsualizar_bitacora(self.archivo)
            self.after(0, lambda datos=datos: self._on_preview_success(datos))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje: self._on_error(mensaje, en_preview=True))

    def _on_preview_success(self, datos):
        self._set_estado_cargando(False, "Estado: previsualización completada")

        cantidad = len(datos)

        if cantidad == 0:
            messagebox.showinfo(
                "Previsualización",
                "No se encontraron registros para mostrar.",
            )
            return

        primeros = datos[:5]

        lineas = []
        for i, reg in enumerate(primeros, start=1):
            correlativo = reg.get("CORRELATIVO", "")
            rit = reg.get("RIT", "")
            imputado = reg.get("IMPUTADO", "")
            sala = reg.get("SALA", "")
            lineas.append(
                f"{i}. Sala {sala} | N° {correlativo} | RIT {rit} | {imputado}"
            )

        texto = (
            f"Se encontraron {cantidad} registros en la previsualización.\n\n"
            + "\n".join(lineas)
        )

        messagebox.showinfo("Previsualización", texto)

    def generar_en_hilo(self):
        if not self.archivo:
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        self._set_estado_cargando(True, "Estado: procesando bitácora...")

        hilo = threading.Thread(target=self._generar_worker, daemon=True)
        hilo.start()

    def _generar_worker(self):
        try:
            salida = generar_bitacora(self.archivo)
            self.after(0, lambda salida=salida: self._on_generar_success(salida))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje: self._on_error(mensaje, en_preview=False))

    def _on_generar_success(self, salida):
        self._set_estado_cargando(False, "Estado: bitácora generada correctamente")

        abrir = messagebox.askyesno(
            "Éxito",
            f"Bitácora generada correctamente.\n\nArchivo:\n{salida}\n\n¿Deseas abrirlo?",
        )

        if abrir:
            abrir_archivo(salida)

    def _on_error(self, mensaje, en_preview=False):
        self._set_estado_cargando(False, "Estado: error")

        titulo = "Error en previsualización" if en_preview else "Error al generar bitácora"
        messagebox.showerror(titulo, mensaje)

    def _set_estado_cargando(self, cargando, texto):
        self.status_label.configure(text=texto)

        estado = "disabled" if cargando else "normal"

        self.btn_seleccionar.configure(state=estado)
        self.btn_preview.configure(state=estado)
        self.btn_generar.configure(state=estado)
        self.btn_volver.configure(state=estado)
