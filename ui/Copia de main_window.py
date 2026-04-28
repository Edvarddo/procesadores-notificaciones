"""
Procesador de Reportes - Interfaz Moderna
Con CustomTkinter + Drag and Drop
"""

import os
import sys
import subprocess
import traceback
from pathlib import Path
from typing import Optional

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("Error: faltan dependencias.")
    print("Instala con:")
    print("  pip install customtkinter tkinterdnd2")
    raise SystemExit(1)

from core.procesador import procesar_archivo


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


class MainWindow(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        ctk.CTk.__init__(self)
        self.TkdndVersion = TkinterDnD._require(self)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.title("Procesador de Reportes")
        self.geometry("800x520")
        self.minsize(700, 480)

        self.ruta_archivo = ctk.StringVar()
        self.archivo_generado: Optional[Path] = None

        self._crear_interfaz()
        self._configurar_dnd()

    def _crear_interfaz(self):
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=30, pady=25)

        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        title_label = ctk.CTkLabel(
            header_frame,
            text="Procesador de Reportes",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(anchor="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Procesa archivos Excel con seleccion manual o arrastrar y soltar",
            font=ctk.CTkFont(size=13),
            text_color=("#6b7280", "#9ca3af")
        )
        subtitle_label.pack(anchor="w", pady=(4, 0))

        theme_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        theme_frame.place(relx=1.0, rely=0.5, anchor="e")

        self.theme_switch = ctk.CTkSwitch(
            theme_frame,
            text="Modo oscuro",
            command=self._toggle_theme,
            font=ctk.CTkFont(size=12),
            width=40
        )
        self.theme_switch.pack()

        if ctk.get_appearance_mode() == "Dark":
            self.theme_switch.select()

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

        self.btn_procesar = ctk.CTkButton(
            self.main_container,
            text="Procesar Archivo",
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.procesar
        )
        self.btn_procesar.pack(fill="x", pady=(0, 15))

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
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self.al_soltar_archivo)
        self.drop_zone.dnd_bind("<<DragEnter>>", lambda e: self.drop_zone.set_hover(True))
        self.drop_zone.dnd_bind("<<DragLeave>>", lambda e: self.drop_zone.set_hover(False))

    def _toggle_theme(self):
        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xls *.xlsx"),
                ("Todos los archivos", "*.*")
            ]
        )
        if ruta:
            self.establecer_archivo(ruta)

    def abrir_explorador(self):
        abrir_ruta(Path.home())

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

        self.ruta_archivo.set(ruta)
        self.status_bar.set_status("Archivo seleccionado y listo para procesar", "success")
        self.drop_zone.set_success(Path(ruta).name)

        self.archivo_generado = None
        self.btn_abrir_archivo.configure(state="disabled")
        self.btn_abrir_carpeta.configure(state="disabled")

    def al_soltar_archivo(self, event):
        self.drop_zone.set_hover(False)
        data = event.data

        if not data:
            return

        rutas = self.tk.splitlist(data)
        if not rutas:
            return

        ruta = rutas[0]
        self.establecer_archivo(ruta)

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

            archivo_salida = procesar_archivo(ruta)
            self.archivo_generado = archivo_salida

            self.status_bar.set_status(
                f"Completado: {archivo_salida.name}",
                "success"
            )
            self.btn_abrir_archivo.configure(state="normal")
            self.btn_abrir_carpeta.configure(state="normal")
            self.btn_procesar.configure(state="normal", text="Procesar Archivo")

            abrir_ahora = messagebox.askyesno(
                "Proceso completado",
                f"Archivo generado correctamente:\n\n{archivo_salida}\n\n¿Deseas abrirlo ahora?"
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
                f"Ocurrio un error al procesar el archivo:\n\n{e}"
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


def main():
    app = MainWindow()
    app.mainloop()
