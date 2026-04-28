import os
import sys
import subprocess
import traceback
from pathlib import Path
from typing import Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES
    DND_DISPONIBLE = True
except ImportError:
    DND_DISPONIBLE = False

from core.reporte_excel.procesador import procesar_archivo, previsualizar_archivo


def abrir_ruta(ruta: str | Path):
    ruta = str(ruta)

    if sys.platform.startswith("win"):
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta], check=False)
    else:
        subprocess.run(["xdg-open", ruta], check=False)


class DropZone(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.configure(
            corner_radius=12,
            border_width=2,
            border_color=("#d1d5db", "#4b5563"),
            fg_color=("#f9fafb", "#1f2937")
        )

        self.inner_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.inner_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.icon_label = ctk.CTkLabel(
            self.inner_frame,
            text="📄",
            font=ctk.CTkFont(size=48)
        )
        self.icon_label.pack(pady=(10, 5))

        self.main_label = ctk.CTkLabel(
            self.inner_frame,
            text="Arrastra aqui tu archivo Excel",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=("#374151", "#e5e7eb")
        )
        self.main_label.pack(pady=(5, 2))

        self.sub_label = ctk.CTkLabel(
            self.inner_frame,
            text="Formatos soportados: .xls, .xlsx",
            font=ctk.CTkFont(size=12),
            text_color=("#6b7280", "#9ca3af")
        )
        self.sub_label.pack()

    def set_hover(self, hovering: bool):
        if hovering:
            self.configure(
                border_color=("#3b82f6", "#60a5fa"),
                fg_color=("#eff6ff", "#1e3a5f")
            )
            self.icon_label.configure(text="📥")
            self.main_label.configure(text="Suelta el archivo aqui")
        else:
            self.reset()

    def set_success(self, filename: str):
        self.configure(
            border_color=("#22c55e", "#4ade80"),
            fg_color=("#f0fdf4", "#14532d")
        )
        self.icon_label.configure(text="✅")
        self.main_label.configure(text="Archivo seleccionado")
        self.sub_label.configure(text=filename)

    def reset(self):
        self.configure(
            border_color=("#d1d5db", "#4b5563"),
            fg_color=("#f9fafb", "#1f2937")
        )
        self.icon_label.configure(text="📄")
        self.main_label.configure(text="Arrastra aqui tu archivo Excel")
        self.sub_label.configure(text="Formatos soportados: .xls, .xlsx")


class StatusBar(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, corner_radius=8, **kwargs)

        self.configure(fg_color=("#e5e7eb", "#374151"))

        self.indicator = ctk.CTkLabel(
            self,
            text="●",
            font=ctk.CTkFont(size=14),
            text_color="#6b7280"
        )
        self.indicator.pack(side="left", padx=(15, 8), pady=10)

        self.status_label = ctk.CTkLabel(
            self,
            text="Esperando archivo...",
            font=ctk.CTkFont(size=13),
            text_color=("#4b5563", "#d1d5db"),
            anchor="w"
        )
        self.status_label.pack(side="left", fill="x", expand=True, pady=10)

    def set_status(self, text: str, status_type: str = "info"):
        colors = {
            "info": "#6b7280",
            "waiting": "#f59e0b",
            "processing": "#3b82f6",
            "success": "#22c55e",
            "error": "#ef4444"
        }
        color = colors.get(status_type, colors["info"])
        self.indicator.configure(text_color=color)
        self.status_label.configure(text=text)


class ReporteView(ctk.CTkFrame):
    def __init__(self, master, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.on_back = on_back
        self.ruta_archivo = ctk.StringVar()
        self.archivo_generado: Optional[Path] = None
        self.ultima_carpeta = Path.home()
        self.mostrar_todas_gestiones = ctk.BooleanVar(value=False)

        self._crear_interfaz()
        self._configurar_dnd()

    def _crear_interfaz(self):
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)

        top_bar = ctk.CTkFrame(self.main_container, fg_color="transparent")
        top_bar.pack(fill="x", pady=(0, 15))

        ctk.CTkButton(
            top_bar,
            text="← Volver",
            width=110,
            height=38,
            corner_radius=8,
            command=self.on_back
        ).pack(anchor="w")

        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            header_frame,
            text="Procesador de Reportes",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Procesa archivos Excel del flujo actual y genera el resultado final",
            font=ctk.CTkFont(size=13),
            text_color=("#6b7280", "#9ca3af")
        ).pack(anchor="w", pady=(4, 0))

        file_section = ctk.CTkFrame(self.main_container, fg_color="transparent")
        file_section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            file_section,
            text="Archivo de entrada",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 8))

        input_row = ctk.CTkFrame(file_section, fg_color="transparent")
        input_row.pack(fill="x")

        self.entry = ctk.CTkEntry(
            input_row,
            textvariable=self.ruta_archivo,
            placeholder_text="Selecciona o arrastra un archivo Excel...",
            height=42,
            font=ctk.CTkFont(size=13),
            corner_radius=8
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_examinar = ctk.CTkButton(
            input_row,
            text="Examinar",
            width=120,
            height=42,
            corner_radius=8,
            font=ctk.CTkFont(size=13),
            command=self.seleccionar_archivo
        )
        self.btn_examinar.pack(side="left", padx=(0, 10))

        texto_explorador = "Finder" if sys.platform == "darwin" else "Explorador"
        self.btn_finder = ctk.CTkButton(
            input_row,
            text=texto_explorador,
            width=120,
            height=42,
            corner_radius=8,
            font=ctk.CTkFont(size=13),
            fg_color=("#e5e7eb", "#4b5563"),
            hover_color=("#d1d5db", "#6b7280"),
            text_color=("#374151", "#f3f4f6"),
            command=self.abrir_explorador
        )
        self.btn_finder.pack(side="left")

        self.drop_zone = DropZone(self.main_container)
        self.drop_zone.pack(fill="x", pady=(10, 20), ipady=15)

        if not DND_DISPONIBLE:
            self.drop_zone.sub_label.configure(
                text="tkinterdnd2 no instalado. Usa Examinar."
            )

        opciones_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        opciones_frame.pack(fill="x", pady=(0, 15))

        self.chk_gestiones = ctk.CTkCheckBox(
            opciones_frame,
            text="Mostrar las 3 gestiones (codigo y hora). Si no, se mostrará solo la más actualizada.",
            variable=self.mostrar_todas_gestiones,
            onvalue=True,
            offvalue=False,
            command=self.refrescar_preview
        )
        self.chk_gestiones.pack(anchor="w")

        self.btn_procesar = ctk.CTkButton(
            self.main_container,
            text="Procesar Archivo",
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.procesar
        )
        self.btn_procesar.pack(fill="x", pady=(0, 15))

        self.preview_info = ctk.CTkLabel(
            self.main_container,
            text="Previsualizacion vacia. Selecciona un archivo.",
            anchor="w",
            justify="left",
            font=ctk.CTkFont(size=12),
            text_color=("#4b5563", "#d1d5db")
        )
        self.preview_info.pack(fill="x", pady=(0, 8))

        preview_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        preview_container.pack(fill="both", expand=True, pady=(0, 15))

        self.tree = ttk.Treeview(preview_container, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree_scroll_y = ttk.Scrollbar(preview_container, orient="vertical", command=self.tree.yview)
        self.tree_scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set)

        actions_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        actions_frame.pack(fill="x", pady=(0, 15))

        self.btn_abrir_archivo = ctk.CTkButton(
            actions_frame,
            text="Abrir archivo generado",
            height=40,
            corner_radius=8,
            font=ctk.CTkFont(size=13),
            fg_color=("#22c55e", "#16a34a"),
            hover_color=("#16a34a", "#15803d"),
            state="disabled",
            command=self.abrir_archivo_generado
        )
        self.btn_abrir_archivo.pack(side="left", expand=True, fill="x", padx=(0, 8))

        self.btn_abrir_carpeta = ctk.CTkButton(
            actions_frame,
            text="Abrir carpeta",
            height=40,
            corner_radius=8,
            font=ctk.CTkFont(size=13),
            fg_color=("#6366f1", "#4f46e5"),
            hover_color=("#4f46e5", "#4338ca"),
            state="disabled",
            command=self.abrir_carpeta_resultado
        )
        self.btn_abrir_carpeta.pack(side="left", expand=True, fill="x", padx=(8, 0))

        self.status_bar = StatusBar(self.main_container)
        self.status_bar.pack(fill="x", side="bottom")

    def _configurar_dnd(self):
        if not DND_DISPONIBLE:
            return

        try:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self.al_soltar_archivo)
            self.drop_zone.dnd_bind("<<DragEnter>>", lambda e: self.drop_zone.set_hover(True))
            self.drop_zone.dnd_bind("<<DragLeave>>", lambda e: self.drop_zone.set_hover(False))
        except Exception:
            self.drop_zone.sub_label.configure(
                text="Drag and drop no disponible en esta configuracion. Usa Examinar."
            )

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            initialdir=str(self.ultima_carpeta),
            filetypes=[
                ("Archivos Excel", "*.xls *.xlsx"),
                ("Todos los archivos", "*.*")
            ]
        )
        if ruta:
            self.establecer_archivo(ruta)

    def abrir_explorador(self):
        abrir_ruta(self.ultima_carpeta)

    def establecer_archivo(self, ruta: str):
        ruta = ruta.strip()

        if ruta.startswith("{") and ruta.endswith("}"):
            ruta = ruta[1:-1]

        extension = Path(ruta).suffix.lower()
        if extension not in [".xls", ".xlsx"]:
            messagebox.showwarning(
                "Archivo no valido",
                "Solo se permiten archivos .xls o .xlsx"
            )
            return

        ruta_path = Path(ruta)
        self.ultima_carpeta = ruta_path.parent

        self.ruta_archivo.set(ruta)
        self.status_bar.set_status("Archivo seleccionado y listo para procesar", "success")
        self.drop_zone.set_success(ruta_path.name)

        self.archivo_generado = None
        self.btn_abrir_archivo.configure(state="disabled")
        self.btn_abrir_carpeta.configure(state="disabled")

        self.refrescar_preview()

    def al_soltar_archivo(self, event):
        self.drop_zone.set_hover(False)
        data = event.data

        if not data:
            return

        try:
            rutas = self.tk.splitlist(data)
            if not rutas:
                return

            ruta = rutas[0]
            self.establecer_archivo(ruta)
        except Exception:
            messagebox.showerror(
                "Error",
                "No se pudo interpretar el archivo arrastrado."
            )

    def mostrar_preview(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")

        for _, row in df.iterrows():
            valores = ["" if v is None else str(v) for v in row.tolist()]
            self.tree.insert("", "end", values=valores)

        self.preview_info.configure(
            text=f"Mostrando {len(df)} registros en la previsualizacion."
        )

    def refrescar_preview(self):
        ruta = self.ruta_archivo.get().strip()
        if not ruta:
            return

        try:
            df_preview = previsualizar_archivo(
                ruta,
                mostrar_todas_gestiones=self.mostrar_todas_gestiones.get(),
                limite=50
            )
            self.mostrar_preview(df_preview)
            self.status_bar.set_status("Previsualizacion actualizada", "info")
        except Exception:
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = []
            self.preview_info.configure(text="No se pudo generar la previsualizacion.")
            self.status_bar.set_status("Error al generar previsualizacion", "error")
            print(traceback.format_exc())

    def procesar(self):
        ruta = self.ruta_archivo.get().strip()

        if not ruta:
            messagebox.showwarning(
                "Falta archivo",
                "Por favor, selecciona un archivo primero."
            )
            return

        try:
            self.status_bar.set_status("Procesando archivo... por favor espera", "processing")
            self.btn_procesar.configure(state="disabled", text="Procesando...")
            self.update_idletasks()

            archivo_salida = procesar_archivo(
                ruta,
                mostrar_todas_gestiones=self.mostrar_todas_gestiones.get()
            )
            self.archivo_generado = archivo_salida
            self.ultima_carpeta = archivo_salida.parent

            self.status_bar.set_status(
                f"Completado: {archivo_salida.name}",
                "success"
            )
            self.btn_abrir_archivo.configure(state="normal")
            self.btn_abrir_carpeta.configure(state="normal")
            self.btn_procesar.configure(state="normal", text="Procesar Archivo")

            abrir_ahora = messagebox.askyesno(
                "Proceso completado",
                f"Archivo generado correctamente:{archivo_salida} ¿Deseas abrirlo ahora?"
            )

            if abrir_ahora:
                abrir_ruta(archivo_salida)

        except Exception as e:
            self.archivo_generado = None
            self.btn_abrir_archivo.configure(state="disabled")
            self.btn_abrir_carpeta.configure(state="disabled")
            self.btn_procesar.configure(state="normal", text="Procesar Archivo")
            self.status_bar.set_status("Error al procesar el archivo", "error")

            messagebox.showerror(
                "Error",
                f"Ocurrio un error al procesar el archivo:{e}"
            )
            print(traceback.format_exc())

    def abrir_archivo_generado(self):
        if self.archivo_generado and self.archivo_generado.exists():
            abrir_ruta(self.archivo_generado)
        else:
            messagebox.showwarning(
                "Archivo no disponible",
                "No se encontro el archivo generado."
            )

    def abrir_carpeta_resultado(self):
        if self.archivo_generado and self.archivo_generado.exists():
            abrir_ruta(self.archivo_generado.parent)
        else:
            messagebox.showwarning(
                "Carpeta no disponible",
                "No se encontro la carpeta del resultado."
            )
