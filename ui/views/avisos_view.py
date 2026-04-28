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

from core.avisos_excel.procesador import (
    previsualizar_avisos,
    generar_avisos,
    obtener_ano_base,
)


def abrir_ruta(ruta: str | Path):
    ruta = str(ruta)

    if sys.platform.startswith("win"):
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta], check=False)
    else:
        subprocess.run(["xdg-open", ruta], check=False)


class DropZone(ctk.CTkFrame):
    def __init__(self, master, texto="Arrastra aqui tu archivo Excel", **kwargs):
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
            text=texto,
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

        self._default_text = texto

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
        self.main_label.configure(text=self._default_text)
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


class AvisosView(ctk.CTkFrame):
    def __init__(self, master, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.on_back = on_back
        self.ruta_archivo = ctk.StringVar()
        self.fecha_aviso = ctk.StringVar()
        self.nombre_manual = ctk.StringVar()
        self.rit_manual = ctk.StringVar()
        self.ano_manual = ctk.StringVar()

        self.archivo_generado: Optional[Path] = None
        self.ultima_carpeta = Path.home()

        self.df_base = None
        self.df_final = None
        self.ano_base = ""

        self._crear_interfaz()
        self._configurar_dnd()

    def _crear_interfaz(self):
        self.pack(fill="both", expand=True)

        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.main_container = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True)

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
            text="Generador de Avisos",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Filtra Personal/Art44 desde Detalle de Impresion y rellena la plantilla de avisos",
            font=ctk.CTkFont(size=13),
            text_color=("#6b7280", "#9ca3af")
        ).pack(anchor="w", pady=(4, 0))

        file_section = ctk.CTkFrame(self.main_container, fg_color="transparent")
        file_section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            file_section,
            text="Archivo Detalle de Impresion",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 8))

        input_row = ctk.CTkFrame(file_section, fg_color="transparent")
        input_row.pack(fill="x")

        self.entry = ctk.CTkEntry(
            input_row,
            textvariable=self.ruta_archivo,
            placeholder_text="Selecciona o arrastra el archivo Detalle de Impresion...",
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

        self.drop_zone = DropZone(
            self.main_container,
            texto="Arrastra aqui el Detalle de Impresion"
        )
        self.drop_zone.pack(fill="x", pady=(10, 20), ipady=15)

        if not DND_DISPONIBLE:
            self.drop_zone.sub_label.configure(
                text="tkinterdnd2 no instalado. Usa Examinar."
            )

        fecha_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        fecha_frame.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            fecha_frame,
            text="Fecha para todos los avisos",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 8))

        self.entry_fecha = ctk.CTkEntry(
            fecha_frame,
            textvariable=self.fecha_aviso,
            placeholder_text="Opcional. Ejemplo: 30-03-2026",
            height=42,
            font=ctk.CTkFont(size=13),
            corner_radius=8
        )
        self.entry_fecha.pack(fill="x")

        ctk.CTkLabel(
            fecha_frame,
            text="Si lo dejas vacio, los avisos saldran sin fecha para completarla a mano.",
            font=ctk.CTkFont(size=11),
            text_color=("#6b7280", "#9ca3af")
        ).pack(anchor="w", pady=(6, 0))

        acciones_lista = ctk.CTkFrame(self.main_container, fg_color="transparent")
        acciones_lista.pack(fill="x", pady=(0, 15))

        self.btn_quitar = ctk.CTkButton(
            acciones_lista,
            text="Quitar seleccionado",
            height=40,
            corner_radius=8,
            command=self.quitar_seleccionado
        )
        self.btn_quitar.pack(side="left", padx=(0, 10))

        self.lbl_resumen = ctk.CTkLabel(
            acciones_lista,
            text="Año base sugerido: - | Registros finales: 0",
            font=ctk.CTkFont(size=12),
            text_color=("#4b5563", "#d1d5db")
        )
        self.lbl_resumen.pack(side="left", fill="x", expand=True)

        manual_frame = ctk.CTkFrame(self.main_container)
        manual_frame.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            manual_frame,
            text="Agregar registro manual",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=12, pady=(10, 8))

        labels_row = ctk.CTkFrame(manual_frame, fg_color="transparent")
        labels_row.pack(fill="x", padx=12, pady=(0, 4))

        ctk.CTkLabel(
            labels_row,
            text="Nombre completo",
            anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkLabel(
            labels_row,
            text="RIT",
            width=160,
            anchor="w"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            labels_row,
            text="Año",
            width=120,
            anchor="w"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            labels_row,
            text="",
            width=130
        ).pack(side="left")

        form_row = ctk.CTkFrame(manual_frame, fg_color="transparent")
        form_row.pack(fill="x", padx=12, pady=(0, 10))

        self.entry_nombre_manual = ctk.CTkEntry(
            form_row,
            textvariable=self.nombre_manual,
            placeholder_text="Ej: JUAN PEREZ GONZALEZ",
            height=40
        )
        self.entry_nombre_manual.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.entry_rit_manual = ctk.CTkEntry(
            form_row,
            textvariable=self.rit_manual,
            placeholder_text="Ej: 1234",
            width=160,
            height=40
        )
        self.entry_rit_manual.pack(side="left", padx=(0, 10))

        self.entry_ano_manual = ctk.CTkEntry(
            form_row,
            textvariable=self.ano_manual,
            placeholder_text="Ej: 2026",
            width=120,
            height=40
        )
        self.entry_ano_manual.pack(side="left", padx=(0, 10))

        self.btn_agregar_manual = ctk.CTkButton(
            form_row,
            text="Agregar",
            width=130,
            height=40,
            command=self.agregar_manual
        )
        self.btn_agregar_manual.pack(side="left")

        self.btn_generar = ctk.CTkButton(
            self.main_container,
            text="Generar Avisos",
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.generar
        )
        self.btn_generar.pack(fill="x", pady=(0, 15))

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

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        if sys.platform.startswith("win"):
            tree_font = ("Segoe UI", 12)
            heading_font = ("Segoe UI", 12, "bold")
            row_height = 34
        else:
            tree_font = ("Arial", 13)
            heading_font = ("Arial", 13, "bold")
            row_height = 30

        style.configure(
            "Avisos.Treeview",
            background="#1f2937",
            foreground="#f9fafb",
            fieldbackground="#1f2937",
            rowheight=row_height,
            borderwidth=0,
            font=tree_font
        )

        style.configure(
            "Avisos.Treeview.Heading",
            background="#111827",
            foreground="#ffffff",
            relief="flat",
            font=heading_font
        )

        style.map(
            "Avisos.Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")]
        )

        table_frame = ctk.CTkFrame(preview_container, fg_color="transparent")
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            table_frame,
            show="headings",
            style="Avisos.Treeview"
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.tree_scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")

        self.tree_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")

        self.tree.configure(
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set
        )

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

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
        self.status_bar.set_status("Archivo seleccionado y listo para previsualizar", "success")
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

    def mostrar_preview(self):
        df = self.df_final if self.df_final is not None else None

        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.tree["columns"] = []
            self.preview_info.configure(text="No hay registros para mostrar.")
            self.lbl_resumen.configure(text=f"Año base sugerido: {self.ano_base or '-'} | Registros finales: 0")
            return

        mostrar = df.head(50).copy()
        self.tree["columns"] = list(mostrar.columns)

        for col in mostrar.columns:
            self.tree.heading(col, text=col)

            if col == "NOMBRE":
                ancho = 420
            elif col == "RIT-AÑO":
                ancho = 150
            else:
                ancho = 140

            self.tree.column(col, width=ancho, minwidth=110, anchor="center")

        for _, row in mostrar.iterrows():
            valores = ["" if v is None else str(v) for v in row.tolist()]
            self.tree.insert("", "end", values=valores)

        self.preview_info.configure(
            text=f"Mostrando {len(mostrar)} de {len(df)} registros finales."
        )
        self.lbl_resumen.configure(
            text=f"Año base sugerido: {self.ano_base or '-'} | Registros finales: {len(df)}"
        )

    def refrescar_preview(self):
        ruta = self.ruta_archivo.get().strip()
        if not ruta:
            return

        try:
            self.df_base = previsualizar_avisos(ruta, limite=None)
            self.df_final = self.df_base.copy()
            self.ano_base = obtener_ano_base(ruta)
            self.mostrar_preview()
            self.status_bar.set_status("Previsualizacion actualizada", "info")
        except Exception:
            self.df_base = None
            self.df_final = None
            self.ano_base = ""
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = []
            self.preview_info.configure(text="No se pudo generar la previsualizacion.")
            self.lbl_resumen.configure(text="Año base sugerido: - | Registros finales: 0")
            self.status_bar.set_status("Error al generar previsualizacion", "error")
            print(traceback.format_exc())

    def quitar_seleccionado(self):
        if self.df_final is None or self.df_final.empty:
            return

        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin seleccion", "Debes seleccionar una fila para quitar.")
            return

        item = seleccion[0]
        valores = self.tree.item(item, "values")
        if not valores:
            return

        nombre = str(valores[3]) if len(valores) > 3 else ""
        rit_ano = str(valores[4]) if len(valores) > 4 else ""

        mask = ~(
            (self.df_final["NOMBRE"].astype(str) == nombre) &
            (self.df_final["RIT-AÑO"].astype(str) == rit_ano)
        )
        self.df_final = self.df_final[mask].reset_index(drop=True)
        self.mostrar_preview()
        self.status_bar.set_status("Registro quitado de la seleccion final", "info")

    def agregar_manual(self):
        nombre = self.nombre_manual.get().strip()
        rit = self.rit_manual.get().strip()
        ano = self.ano_manual.get().strip()

        if not nombre or not rit or not ano:
            messagebox.showwarning(
                "Datos incompletos",
                "Debes ingresar NOMBRE, RIT y AÑO."
            )
            return

        nuevo = {
            "RUC": "",
            "RIT": rit,
            "AÑO": ano,
            "NOMBRE": nombre,
            "RIT-AÑO": f"{rit}-{ano}",
        }

        if self.df_final is None:
            import pandas as pd
            self.df_final = pd.DataFrame([nuevo])
        else:
            import pandas as pd
            self.df_final = pd.concat([self.df_final, pd.DataFrame([nuevo])], ignore_index=True)

        self.nombre_manual.set("")
        self.rit_manual.set("")
        self.ano_manual.set("")
        self.mostrar_preview()
        self.status_bar.set_status("Registro manual agregado", "success")

    def generar(self):
        ruta = self.ruta_archivo.get().strip()
        fecha = self.fecha_aviso.get().strip()

        if not ruta:
            messagebox.showwarning(
                "Falta archivo",
                "Por favor, selecciona un archivo primero."
            )
            return

        if self.df_final is None or self.df_final.empty:
            messagebox.showwarning(
                "Sin registros",
                "No hay registros finales para generar avisos."
            )
            return

        try:
            self.status_bar.set_status("Generando avisos... por favor espera", "processing")
            self.btn_generar.configure(state="disabled", text="Generando...")
            self.update_idletasks()

            archivo_salida = generar_avisos(
                ruta_archivo=ruta,
                fecha=fecha,
                df_final=self.df_final
            )

            self.archivo_generado = archivo_salida
            self.ultima_carpeta = archivo_salida.parent

            self.status_bar.set_status(
                f"Completado: {archivo_salida.name}",
                "success"
            )
            self.btn_abrir_archivo.configure(state="normal")
            self.btn_abrir_carpeta.configure(state="normal")
            self.btn_generar.configure(state="normal", text="Generar Avisos")

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
            self.btn_generar.configure(state="normal", text="Generar Avisos")
            self.status_bar.set_status("Error al generar avisos", "error")

            messagebox.showerror(
                "Error",
                f"Ocurrio un error al generar los avisos:\n\n{e}"
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
