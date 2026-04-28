import os
import sys
import subprocess
import threading
import customtkinter as ctk

from tkinter import filedialog, messagebox

from core.carabineros_formulario.procesador import (
    ejecutar_carabineros,
    previsualizar_carabineros,
    generar_csv_cinj_desde_excel,
    limpiar_en_cinj_carabineros,
)
from core.carabineros_formulario.procesador_impresion import (
    leer_archivo_impresion,
    generar_csv_desde_impresion,
    previsualizar_impresion,
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
class DialogoHoraCodigo(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)

        self.title("Datos de certificación")
        self.geometry("360x260")
        self.resizable(False, False)

        self.resultado = None

        ctk.CTkLabel(self, text="Hora (ej: 1205)").pack(pady=(20, 5))
        self.entry_hora = ctk.CTkEntry(self, width=180)
        self.entry_hora.pack(pady=5)
        self.entry_hora.insert(0, "1205")

        ctk.CTkLabel(self, text="Código (ej: D2)").pack(pady=(10, 5))
        self.entry_codigo = ctk.CTkEntry(self, width=180)
        self.entry_codigo.pack(pady=5)
        self.entry_codigo.insert(0, "D2")

        self.btn_aceptar = ctk.CTkButton(
            self,
            text="Aceptar y continuar",
            command=self.aceptar
        )
        self.btn_aceptar.pack(pady=(20, 8))

        self.btn_cancelar = ctk.CTkButton(
            self,
            text="Cancelar",
            fg_color="gray",
            command=self.destroy
        )
        self.btn_cancelar.pack(pady=(0, 10))

        self.entry_hora.focus()
        self.bind("<Return>", lambda e: self.aceptar())

        self.grab_set()

    def aceptar(self):
        hora = self.entry_hora.get().strip()
        codigo = self.entry_codigo.get().strip().upper()

        if not hora.isdigit() or len(hora) != 4:
            messagebox.showerror("Error", "La hora debe tener formato HHMM, por ejemplo 1205.")
            return

        if not codigo:
            messagebox.showerror("Error", "Debes ingresar un código, por ejemplo D2.")
            return

        self.resultado = (hora, codigo)
        self.destroy()

class CarabinerosView(ctk.CTkFrame):
    def __init__(self, master, volver_callback):
        super().__init__(master)

        self.volver_callback = volver_callback
        self.archivo_terreno = None
        self.archivo_carabineros = None
        self.archivo_impresion = None
        self.tipo_actual = "carabineros"  # terreno o carabineros

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._crear_widgets()

    def _crear_widgets(self):
        # Encabezado con título y selector
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        header_frame.grid_columnconfigure(0, weight=1)

        titulo_frame = ctk.CTkFrame(header_frame)
        titulo_frame.pack(side="left")
        self.titulo = ctk.CTkLabel(
            titulo_frame,
            text="Automatización Formulario",
            font=("Arial", 20, "bold"),
        )
        self.titulo.pack()

        selector_frame = ctk.CTkFrame(header_frame)
        selector_frame.pack(side="right", padx=20)

        ctk.CTkLabel(
            selector_frame,
            text="Tipo:",
            font=("Arial", 11, "bold")
        ).pack(side="left", padx=(0, 10))

        self.tipo_selector = ctk.CTkSegmentedButton(
            selector_frame,
            values=["Terreno", "Carabineros"],
            command=self._cambiar_tipo,
            font=("Arial", 11),
        )
        self.tipo_selector.set("Carabineros")
        self.tipo_selector.pack(side="left")

        # Separador visual
        separator = ctk.CTkFrame(self, height=2, fg_color=("gray70", "gray30"))
        separator.grid(row=1, column=0, sticky="ew", pady=10)

        # Sección SOLO para Carabineros: Procesador de Impresión
        self.impresion_frame = ctk.CTkFrame(self)
        self.impresion_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(10, 10))
        self.impresion_frame.grid_columnconfigure((0, 1), weight=1)
        self._crear_widgets_procesador_impresion()

        # Frame contenedor para los tabs
        self.contenedor_tabs = ctk.CTkFrame(self)
        self.contenedor_tabs.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.contenedor_tabs.grid_columnconfigure(0, weight=1)
        self.contenedor_tabs.grid_rowconfigure(0, weight=1)

        # Frame Terreno
        self.terreno_frame = ctk.CTkFrame(self.contenedor_tabs)
        self.terreno_frame.grid(row=0, column=0, sticky="nsew")
        self.terreno_frame.grid_columnconfigure((0, 1), weight=1)
        self._crear_widgets_terreno()

        # Frame Carabineros
        self.carabineros_frame = ctk.CTkFrame(self.contenedor_tabs)
        self.carabineros_frame.grid(row=0, column=0, sticky="nsew")
        self.carabineros_frame.grid_columnconfigure((0, 1), weight=1)
        self._crear_widgets_carabineros()

        # Mostrar Carabineros por defecto
        self.carabineros_frame.tkraise()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

    def _cambiar_tipo(self, valor):
        """Cambia entre Terreno y Carabineros"""
        if valor == "Terreno":
            self.tipo_actual = "terreno"
            self.terreno_frame.tkraise()
            self.impresion_frame.grid_remove()  # Ocultar procesador de impresión
        else:
            self.tipo_actual = "carabineros"
            self.carabineros_frame.tkraise()
            self.impresion_frame.grid()  # Mostrar procesador de impresión

    def _crear_widgets_procesador_impresion(self):
        """Sección para procesar archivos de impresión de Carabineros"""
        # Sección: Procesador de Impresión (solo para Carabineros)
        impresion_section = ctk.CTkFrame(self.impresion_frame, fg_color=("gray85", "gray25"))
        impresion_section.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
        impresion_section.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            impresion_section,
            text="📋 PROCESADOR DE IMPRESIÓN",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=15, pady=(10, 8))

        self.btn_seleccionar_impresion = ctk.CTkButton(
            impresion_section,
            text="Seleccionar archivo",
            command=self.seleccionar_archivo_impresion,
            height=28,
            font=("Arial", 10)
        )
        self.btn_seleccionar_impresion.grid(row=1, column=0, sticky="ew", padx=(15, 5), pady=(0, 10))

        self.btn_preview_impresion = ctk.CTkButton(
            impresion_section,
            text="Previsualizar",
            command=self.previsualizar_impresion,
            height=28,
            font=("Arial", 10)
        )
        self.btn_preview_impresion.grid(row=1, column=1, sticky="ew", padx=5, pady=(0, 10))

        self.btn_procesar_impresion = ctk.CTkButton(
            impresion_section,
            text="Procesar",
            command=self.procesar_impresion_en_hilo,
            height=28,
            font=("Arial", 10)
        )
        self.btn_procesar_impresion.grid(row=1, column=2, sticky="ew", padx=(5, 15), pady=(0, 10))

        self.label_impresion = ctk.CTkLabel(
            impresion_section,
            text="Ningún archivo seleccionado",
            wraplength=900,
            justify="left",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        )
        self.label_impresion.grid(row=2, column=0, columnspan=3, sticky="ew", padx=15, pady=(0, 10))

        self.status_impresion = ctk.CTkLabel(
            impresion_section,
            text="Estado: esperando archivo...",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        )
        self.status_impresion.grid(row=3, column=0, columnspan=3, sticky="ew", padx=15, pady=(0, 10))

    def _crear_widgets_terreno(self):
        """Widgets para el flujo de Terreno - Layout mejorado con 2 columnas"""
        self.terreno_frame.grid_rowconfigure(0, weight=1)
        self.terreno_frame.grid_rowconfigure(1, weight=0)

        # Contenedor principal de las 2 columnas
        content_frame = ctk.CTkFrame(self.terreno_frame)
        content_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=15, pady=15)
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)

        # Columna izquierda: Datos de entrada
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(0, weight=0)
        left_frame.grid_rowconfigure(1, weight=0)

        # Sección: Seleccionar archivo
        file_section = ctk.CTkFrame(left_frame, fg_color=("gray90", "gray20"))
        file_section.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        file_section.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_section,
            text="📁 ARCHIVO",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 8))

        self.btn_seleccionar_terreno = ctk.CTkButton(
            file_section,
            text="Seleccionar CSV",
            command=lambda: self.seleccionar_archivo("terreno"),
            height=32
        )
        self.btn_seleccionar_terreno.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 10))

        self.label_archivo_terreno = ctk.CTkLabel(
            file_section,
            text="Ningún archivo seleccionado",
            wraplength=280,
            justify="left",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        )
        self.label_archivo_terreno.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Sección: Parámetros
        params_section = ctk.CTkFrame(left_frame, fg_color=("gray90", "gray20"))
        params_section.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        params_section.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            params_section,
            text="⚙️ PARÁMETROS",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 8))

        ctk.CTkLabel(
            params_section,
            text="Fecha (opcional)",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        ).grid(row=1, column=0, sticky="w", padx=15, pady=(0, 5))

        self.entry_fecha_terreno = ctk.CTkEntry(
            params_section,
            placeholder_text="ej: 01/01/2026 - 21/04/2026",
            height=28
        )
        self.entry_fecha_terreno.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Columna derecha: Acciones y estado
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(0, weight=0)
        right_frame.grid_rowconfigure(1, weight=1)

        # Sección: Acciones rápidas
        quick_section = ctk.CTkFrame(right_frame, fg_color=("gray90", "gray20"))
        quick_section.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        quick_section.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            quick_section,
            text="⚡ ACCIONES",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(10, 10))

        self.btn_preview_terreno = ctk.CTkButton(
            quick_section,
            text="👁️ Previsualizar",
            command=lambda: self.preview("terreno"),
            height=32
        )
        self.btn_preview_terreno.grid(row=1, column=0, sticky="ew", padx=(15, 7), pady=(0, 10))

        self.btn_generar_csv_terreno = ctk.CTkButton(
            quick_section,
            text="📊 CSV CINJ",
            command=lambda: self.generar_csv_cinj_en_hilo("terreno"),
            height=32
        )
        self.btn_generar_csv_terreno.grid(row=1, column=1, sticky="ew", padx=(7, 15), pady=(0, 10))

        self.btn_limpiar_terreno = ctk.CTkButton(
            quick_section,
            text="🗑️ Limpiar",
            command=lambda: self.limpiar_cinj_en_hilo("terreno"),
            height=32
        )
        self.btn_limpiar_terreno.grid(row=2, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 10))

        # Sección: Ejecución principal
        exec_section = ctk.CTkFrame(right_frame, fg_color=("gray85", "gray25"))
        exec_section.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 15))
        exec_section.grid_columnconfigure(0, weight=1)
        exec_section.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            exec_section,
            text="▶️ EJECUTAR",
            font=("Arial", 12, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 10))

        self.btn_ejecutar_terreno = ctk.CTkButton(
            exec_section,
            text="Ejecutar Automatización",
            command=lambda: self.ejecutar_en_hilo("terreno"),
            height=40,
            font=("Arial", 11, "bold")
        )
        self.btn_ejecutar_terreno.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        self.status_label_terreno = ctk.CTkLabel(
            exec_section,
            text="Estado: esperando acción...",
            justify="left",
            font=("Arial", 9)
        )
        self.status_label_terreno.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)

        # Pie: Botón volver
        self.btn_volver_terreno = ctk.CTkButton(
            self.terreno_frame,
            text="← Volver",
            command=self.volver_callback,
            height=32,
            fg_color=("gray60", "gray50"),
            hover_color=("gray70", "gray40")
        )
        self.btn_volver_terreno.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 15))

    def _crear_widgets_carabineros(self):
        """Widgets para el flujo de Carabineros - Layout mejorado con 2 columnas"""
        self.carabineros_frame.grid_rowconfigure(0, weight=1)
        self.carabineros_frame.grid_rowconfigure(1, weight=0)

        # Contenedor principal de las 2 columnas
        content_frame = ctk.CTkFrame(self.carabineros_frame)
        content_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=15, pady=15)
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)

        # Columna izquierda: Datos de entrada
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(0, weight=0)
        left_frame.grid_rowconfigure(1, weight=0)

        # Sección: Seleccionar archivo
        file_section = ctk.CTkFrame(left_frame, fg_color=("gray90", "gray20"))
        file_section.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        file_section.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_section,
            text="📁 ARCHIVO",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 8))

        self.btn_seleccionar_carabineros = ctk.CTkButton(
            file_section,
            text="Seleccionar CSV",
            command=lambda: self.seleccionar_archivo("carabineros"),
            height=32
        )
        self.btn_seleccionar_carabineros.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 10))

        self.label_archivo_carabineros = ctk.CTkLabel(
            file_section,
            text="Ningún archivo seleccionado",
            wraplength=280,
            justify="left",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        )
        self.label_archivo_carabineros.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Sección: Parámetros
        params_section = ctk.CTkFrame(left_frame, fg_color=("gray90", "gray20"))
        params_section.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        params_section.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            params_section,
            text="⚙️ PARÁMETROS",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 8))

        ctk.CTkLabel(
            params_section,
            text="Fecha (opcional)",
            font=("Arial", 9),
            text_color=("gray60", "gray40")
        ).grid(row=1, column=0, sticky="w", padx=15, pady=(0, 5))

        self.entry_fecha_carabineros = ctk.CTkEntry(
            params_section,
            placeholder_text="ej: 01/01/2026 - 21/04/2026",
            height=28
        )
        self.entry_fecha_carabineros.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Columna derecha: Acciones y estado
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(0, weight=0)
        right_frame.grid_rowconfigure(1, weight=1)

        # Sección: Acciones rápidas
        quick_section = ctk.CTkFrame(right_frame, fg_color=("gray90", "gray20"))
        quick_section.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        quick_section.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            quick_section,
            text="⚡ ACCIONES",
            font=("Arial", 11, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(10, 10))

        self.btn_preview_carabineros = ctk.CTkButton(
            quick_section,
            text="👁️ Previsualizar",
            command=lambda: self.preview("carabineros"),
            height=32
        )
        self.btn_preview_carabineros.grid(row=1, column=0, sticky="ew", padx=(15, 7), pady=(0, 10))

        self.btn_generar_csv_carabineros = ctk.CTkButton(
            quick_section,
            text="📊 CSV CINJ",
            command=lambda: self.generar_csv_cinj_en_hilo("carabineros"),
            height=32
        )
        self.btn_generar_csv_carabineros.grid(row=1, column=1, sticky="ew", padx=(7, 15), pady=(0, 10))

        self.btn_limpiar_carabineros = ctk.CTkButton(
            quick_section,
            text="🗑️ Limpiar",
            command=lambda: self.limpiar_cinj_en_hilo("carabineros"),
            height=32
        )
        self.btn_limpiar_carabineros.grid(row=2, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 10))

        # Sección: Ejecución principal
        exec_section = ctk.CTkFrame(right_frame, fg_color=("gray85", "gray25"))
        exec_section.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 15))
        exec_section.grid_columnconfigure(0, weight=1)
        exec_section.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            exec_section,
            text="▶️ EJECUTAR",
            font=("Arial", 12, "bold"),
            text_color=("gray20", "gray80")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 10))

        self.btn_ejecutar_carabineros = ctk.CTkButton(
            exec_section,
            text="Ejecutar Automatización",
            command=lambda: self.ejecutar_en_hilo("carabineros"),
            height=40,
            font=("Arial", 11, "bold")
        )
        self.btn_ejecutar_carabineros.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        self.status_label_carabineros = ctk.CTkLabel(
            exec_section,
            text="Estado: esperando acción...",
            justify="left",
            font=("Arial", 9)
        )
        self.status_label_carabineros.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)

        # Pie: Botón volver
        self.btn_volver_carabineros = ctk.CTkButton(
            self.carabineros_frame,
            text="← Volver",
            command=self.volver_callback,
            height=32,
            fg_color=("gray60", "gray50"),
            hover_color=("gray70", "gray40")
        )
        self.btn_volver_carabineros.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 15))

    def _get_archivo(self, tipo):
        """Obtiene el archivo del tipo especificado"""
        if tipo == "terreno":
            return self.archivo_terreno
        else:
            return self.archivo_carabineros

    def _set_archivo(self, tipo, valor):
        """Asigna el archivo del tipo especificado"""
        if tipo == "terreno":
            self.archivo_terreno = valor
        else:
            self.archivo_carabineros = valor

    def _get_status_label(self, tipo):
        """Obtiene el label de estado"""
        return self.status_label_terreno if tipo == "terreno" else self.status_label_carabineros

    def _get_label_archivo(self, tipo):
        """Obtiene el label del archivo"""
        return self.label_archivo_terreno if tipo == "terreno" else self.label_archivo_carabineros

    def _get_entry_fecha(self, tipo):
        """Obtiene el entry de fecha"""
        return self.entry_fecha_terreno if tipo == "terreno" else self.entry_fecha_carabineros

    def seleccionar_archivo(self, tipo):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo fuente",
            filetypes=[("Archivos compatibles", "*.xls *.xlsx *.csv")],
        )

        if not ruta:
            return

        self._set_archivo(tipo, ruta)
        self._get_label_archivo(tipo).configure(text=ruta)
        self._get_status_label(tipo).configure(text="Estado: archivo cargado")

    def preview(self, tipo):
        if not self._get_archivo(tipo):
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        self._set_estado_cargando(tipo, True, "Estado: generando previsualización...")
        hilo = threading.Thread(target=self._preview_worker, args=(tipo,), daemon=True)
        hilo.start()

    def _preview_worker(self, tipo):
        try:
            datos = previsualizar_carabineros(self._get_archivo(tipo))
            self.after(0, lambda datos=datos, tipo=tipo: self._on_preview_success(datos, tipo))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje, tipo=tipo: self._on_error(mensaje, tipo, en_preview=True))

    def _on_preview_success(self, datos, tipo):
        self._set_estado_cargando(tipo, False, "Estado: previsualización completada")

        if not datos:
            messagebox.showinfo("Previsualización", "No se encontraron registros.")
            return

        lineas = []
        for i, reg in enumerate(datos[:5], start=1):
            rit = getattr(reg, "rit", "")
            anio = getattr(reg, "anio", "")
            id_notificacion = getattr(reg, "id_notificacion", "")
            lineas.append(f"{i}. RIT {rit}-{anio} | ID {id_notificacion}")

        texto = "Se cargaron registros correctamente.\n\n" + "\n".join(lineas)
        messagebox.showinfo("Previsualización", texto)

    def ejecutar_en_hilo(self, tipo):
        if not self._get_archivo(tipo):
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        self._set_estado_cargando(tipo, True, "Estado: ejecutando automatización...")
        hilo = threading.Thread(target=self._ejecutar_worker, args=(tipo,), daemon=True)
        hilo.start()

    def _ejecutar_worker(self, tipo):
        try:
            fecha = self._get_entry_fecha(tipo).get().strip() or None
            salida = ejecutar_carabineros(self._get_archivo(tipo), fecha)
            self.after(0, lambda salida=salida, tipo=tipo: self._on_success(salida, tipo))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje, tipo=tipo: self._on_error(mensaje, tipo, en_preview=False))

    def _on_success(self, salida, tipo):
        self._set_estado_cargando(tipo, False, "Estado: automatización completada")

        abrir = messagebox.askyesno(
            "Éxito",
            f"Automatización completada.\n\nResultado:\n{salida}\n\n¿Deseas abrir el archivo?",
        )

        if abrir:
            abrir_archivo(salida)

    def _on_error(self, mensaje, tipo, en_preview=False):
        self._set_estado_cargando(tipo, False, "Estado: error")
        titulo = "Error en previsualización" if en_preview else "Error en automatización"
        messagebox.showerror(titulo, mensaje)

    def _set_estado_cargando(self, tipo, cargando, texto):
        self._get_status_label(tipo).configure(text=texto)
        estado = "disabled" if cargando else "normal"

        botones = [
            self.btn_seleccionar_terreno, self.btn_preview_terreno, self.btn_generar_csv_terreno,
            self.btn_limpiar_terreno, self.btn_ejecutar_terreno, self.btn_volver_terreno,
            self.btn_seleccionar_carabineros, self.btn_preview_carabineros, self.btn_generar_csv_carabineros,
            self.btn_limpiar_carabineros, self.btn_ejecutar_carabineros, self.btn_volver_carabineros,
        ]
        
        for btn in botones:
            btn.configure(state=estado)

    def generar_csv_cinj_en_hilo(self, tipo):
        if not self._get_archivo(tipo):
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        dialogo = DialogoHoraCodigo(self)
        self.wait_window(dialogo)

        if not dialogo.resultado:
            return

        hora, codigo = dialogo.resultado

        self._set_estado_cargando(tipo, True, "Estado: generando CSV para CINJ...")

        hilo = threading.Thread(
            target=self._generar_csv_cinj_worker,
            args=(tipo, hora, codigo),
            daemon=True
        )
        hilo.start()

    def _generar_csv_cinj_worker(self, tipo, hora, codigo):
        try:
            salida = generar_csv_cinj_desde_excel(self._get_archivo(tipo), hora, codigo)
            self.after(0, lambda salida=salida, tipo=tipo: self._on_generar_csv_cinj_success(salida, tipo))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje, tipo=tipo: self._on_error(mensaje, tipo, en_preview=False))

    def _on_generar_csv_cinj_success(self, salida, tipo):
        self._set_estado_cargando(tipo, False, "Estado: CSV para CINJ generado")
        self._set_archivo(tipo, salida)
        self._get_label_archivo(tipo).configure(text=salida)

        abrir = messagebox.askyesno(
            "Éxito",
            f"CSV generado correctamente.\n\nArchivo:\n{salida}\n\n"
            "La app ahora usará este archivo para el proceso.\n\n¿Deseas abrirlo?"
        )

        if abrir:
            abrir_archivo(salida)

    def limpiar_cinj_en_hilo(self, tipo):
        if not self._get_archivo(tipo):
            messagebox.showerror("Error", "Selecciona un archivo primero.")
            return

        self._set_estado_cargando(tipo, True, "Estado: limpiando registros en CINJ...")
        hilo = threading.Thread(target=self._limpiar_cinj_worker, args=(tipo,), daemon=True)
        hilo.start()

    def _limpiar_cinj_worker(self, tipo):
        try:
            salida = limpiar_en_cinj_carabineros(self._get_archivo(tipo))
            self.after(0, lambda salida=salida, tipo=tipo: self._on_limpiar_cinj_success(salida, tipo))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje, tipo=tipo: self._on_error(mensaje, tipo, en_preview=False))

    def _on_limpiar_cinj_success(self, salida, tipo):
        self._set_estado_cargando(tipo, False, "Estado: limpieza en CINJ completada")

        abrir = messagebox.askyesno(
            "Éxito",
            f"Limpieza en CINJ completada.\n\nResultado:\n{salida}\n\n¿Deseas abrir el archivo?"
        )

        if abrir:
            abrir_archivo(salida)

    # ========== Procesador de Impresiones ==========

    def seleccionar_archivo_impresion(self):
        """Selecciona archivo de impresión de Carabineros"""
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de impresión",
            filetypes=[("Archivos Excel", "*.xls *.xlsx"), ("Todos", "*.*")],
        )

        if not ruta:
            return

        self.archivo_impresion = ruta
        self.label_impresion.configure(text=ruta)
        self.status_impresion.configure(text="Estado: archivo de impresión cargado")

    def previsualizar_impresion(self):
        """Previsualiza registros extraídos del archivo de impresión"""
        if not self.archivo_impresion:
            messagebox.showerror("Error", "Selecciona un archivo de impresión primero.")
            return

        try:
            registros = previsualizar_impresion(self.archivo_impresion)

            if not registros:
                messagebox.showinfo("Previsualización", "No se encontraron IDs de notificación.")
                return

            lineas = []
            for i, reg in enumerate(registros, start=1):
                lineas.append(f"{i}. ID: {reg.id_notificacion}")

            texto = f"Se encontraron {len(registros)} registros.\n\nPrimeros:\n" + "\n".join(lineas)
            messagebox.showinfo("Previsualización", texto)
            
            self.status_impresion.configure(text=f"Estado: {len(registros)} registros encontrados")

        except Exception as e:
            messagebox.showerror("Error en previsualización", str(e))

    def procesar_impresion_en_hilo(self):
        """Procesa archivo de impresión: solicita código y hora, genera CSV"""
        if not self.archivo_impresion:
            messagebox.showerror("Error", "Selecciona un archivo de impresión primero.")
            return

        # Abrir diálogo para solicitar código y hora
        dialogo = DialogoHoraCodigo(self)
        self.wait_window(dialogo)

        if not dialogo.resultado:
            return

        hora, codigo = dialogo.resultado

        self.status_impresion.configure(text="Estado: procesando archivo de impresión...")

        hilo = threading.Thread(
            target=self._procesar_impresion_worker,
            args=(hora, codigo),
            daemon=True
        )
        hilo.start()

    def _procesar_impresion_worker(self, hora, codigo):
        try:
            salida = generar_csv_desde_impresion(self.archivo_impresion, codigo, hora)
            self.after(0, lambda salida=salida: self._on_procesar_impresion_success(salida))
        except Exception as e:
            mensaje = str(e)
            self.after(0, lambda mensaje=mensaje: self._on_procesar_impresion_error(mensaje))

    def _on_procesar_impresion_success(self, salida):
        self.status_impresion.configure(text="Estado: CSV generado")

        abrir = messagebox.askyesno(
            "Éxito",
            f"Archivo de impresión procesado correctamente.\n\n"
            f"Archivo:\n{salida}\n\n"
            f"Este CSV está listo para la automatización.\n\n"
            f"¿Deseas abrirlo?"
        )

        if abrir:
            abrir_archivo(salida)

        # Opcionalmente, cargar este archivo en el procesador de Carabineros
        cargar = messagebox.askyesno(
            "Cargar en procesador",
            "¿Deseas usar este archivo en el procesador de Carabineros?"
        )

        if cargar:
            self.archivo_carabineros = salida
            self.label_archivo_carabineros.configure(text=salida)
            self.carabineros_frame.tkraise()
            self.status_label_carabineros.configure(text="Estado: CSV procesado cargado")

    def _on_procesar_impresion_error(self, mensaje):
        self.status_impresion.configure(text="Estado: error")
        messagebox.showerror("Error procesando impresión", mensaje)
