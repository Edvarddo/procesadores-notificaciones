from pathlib import Path
from typing import Optional
import os
import sys
import subprocess
import traceback

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES
    DND_DISPONIBLE = True
except ImportError:
    DND_DISPONIBLE = False

from core.pdf_excel.service import procesar_pdf_excel


def abrir_ruta(ruta: str | Path):
    ruta = str(ruta)

    if sys.platform.startswith("win"):
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta], check=False)
    else:
        subprocess.run(["xdg-open", ruta], check=False)


def aplicar_formato_excel(df: pd.DataFrame, ruta_guardado: Path):
    df = df.copy()

    columnas_numericas = ["RIT_PDF", "ANO_PDF", "HORA"]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    with pd.ExcelWriter(ruta_guardado, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]

        # Formato 4 digitos para HORA
        if "HORA" in df.columns:
            col_idx = df.columns.get_loc("HORA") + 1
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                cell = row[col_idx - 1]
                if cell.value is not None:
                    cell.number_format = "0000"


class FileDropZone(ctk.CTkFrame):
    def __init__(self, master, titulo: str, subtitulo: str, **kwargs):
        super().__init__(master, **kwargs)

        self.titulo_original = titulo
        self.subtitulo_original = subtitulo

        self.configure(
            corner_radius=12,
            border_width=2,
            border_color=("#d1d5db", "#4b5563"),
            fg_color=("#f9fafb", "#1f2937")
        )

        self.inner = ctk.CTkFrame(self, fg_color="transparent")
        self.inner.pack(fill="both", expand=True, padx=20, pady=20)

        self.icon_label = ctk.CTkLabel(
            self.inner,
            text="📄",
            font=ctk.CTkFont(size=42)
        )
        self.icon_label.pack(pady=(5, 5))

        self.title_label = ctk.CTkLabel(
            self.inner,
            text=titulo,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.title_label.pack()

        self.subtitle_label = ctk.CTkLabel(
            self.inner,
            text=subtitulo,
            font=ctk.CTkFont(size=12),
            text_color=("#6b7280", "#9ca3af")
        )
        self.subtitle_label.pack(pady=(4, 0))

    def set_hover(self, hovering: bool):
        if hovering:
            self.configure(
                border_color=("#3b82f6", "#60a5fa"),
                fg_color=("#eff6ff", "#1e3a5f")
            )
            self.icon_label.configure(text="📥")
            self.title_label.configure(text="Suelta el archivo aqui")
        else:
            self.reset()

    def set_success(self, filename: str):
        self.configure(
            border_color=("#22c55e", "#4ade80"),
            fg_color=("#f0fdf4", "#14532d")
        )
        self.icon_label.configure(text="✅")
        self.title_label.configure(text="Archivo seleccionado")
        self.subtitle_label.configure(text=filename)

    def reset(self):
        self.configure(
            border_color=("#d1d5db", "#4b5563"),
            fg_color=("#f9fafb", "#1f2937")
        )
        self.icon_label.configure(text="📄")
        self.title_label.configure(text=self.titulo_original)
        self.subtitle_label.configure(text=self.subtitulo_original)


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
            text="Esperando archivos...",
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


class PdfExcelView(ctk.CTkFrame):
    def __init__(self, master, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.on_back = on_back

        self.ruta_pdf = ctk.StringVar()
        self.ruta_excel = ctk.StringVar()

        self.ultima_carpeta = Path.home()
        self.resultado: Optional[dict] = None

        self._crear_interfaz()
        self._configurar_dnd()

    def _crear_interfaz(self):
        self.main_container = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            corner_radius=0
        )
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
            text="Procesador PDF + Excel",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Extrae PDF, procesa Excel y une ambos resultados en un flujo unico",
            font=ctk.CTkFont(size=13),
            text_color=("#6b7280", "#9ca3af")
        ).pack(anchor="w", pady=(4, 0))

        pdf_section = ctk.CTkFrame(self.main_container, fg_color="transparent")
        pdf_section.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            pdf_section,
            text="Archivo PDF",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 8))

        pdf_row = ctk.CTkFrame(pdf_section, fg_color="transparent")
        pdf_row.pack(fill="x")

        self.entry_pdf = ctk.CTkEntry(
            pdf_row,
            textvariable=self.ruta_pdf,
            placeholder_text="Selecciona o arrastra el PDF...",
            height=42,
            font=ctk.CTkFont(size=13),
            corner_radius=8
        )
        self.entry_pdf.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_pdf = ctk.CTkButton(
            pdf_row,
            text="Buscar PDF",
            width=120,
            height=42,
            command=self.seleccionar_pdf
        )
        self.btn_pdf.pack(side="left")

        self.drop_pdf = FileDropZone(
            self.main_container,
            titulo="Arrastra aqui el archivo PDF",
            subtitulo="Formato soportado: .pdf"
        )
        self.drop_pdf.pack(fill="x", pady=(0, 16), ipady=10)

        excel_section = ctk.CTkFrame(self.main_container, fg_color="transparent")
        excel_section.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            excel_section,
            text="Archivo Excel",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 8))

        excel_row = ctk.CTkFrame(excel_section, fg_color="transparent")
        excel_row.pack(fill="x")

        self.entry_excel = ctk.CTkEntry(
            excel_row,
            textvariable=self.ruta_excel,
            placeholder_text="Selecciona o arrastra el Excel...",
            height=42,
            font=ctk.CTkFont(size=13),
            corner_radius=8
        )
        self.entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_excel = ctk.CTkButton(
            excel_row,
            text="Buscar Excel",
            width=120,
            height=42,
            command=self.seleccionar_excel
        )
        self.btn_excel.pack(side="left")

        self.drop_excel = FileDropZone(
            self.main_container,
            titulo="Arrastra aqui el archivo Excel",
            subtitulo="Formatos soportados: .xls, .xlsx"
        )
        self.drop_excel.pack(fill="x", pady=(0, 18), ipady=10)

        acciones_top = ctk.CTkFrame(self.main_container, fg_color="transparent")
        acciones_top.pack(fill="x", pady=(0, 15))

        texto_explorador = "Abrir Finder" if sys.platform == "darwin" else "Abrir Explorador"
        self.btn_explorador = ctk.CTkButton(
            acciones_top,
            text=texto_explorador,
            width=150,
            height=40,
            fg_color=("#e5e7eb", "#4b5563"),
            hover_color=("#d1d5db", "#6b7280"),
            text_color=("#374151", "#f3f4f6"),
            command=self.abrir_explorador
        )
        self.btn_explorador.pack(side="left")

        self.btn_procesar = ctk.CTkButton(
            acciones_top,
            text="Procesar PDF + Excel",
            height=46,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.procesar
        )
        self.btn_procesar.pack(side="left", fill="x", expand=True, padx=(12, 0))

        acciones_bottom = ctk.CTkFrame(self.main_container, fg_color="transparent")
        acciones_bottom.pack(fill="x", pady=(0, 15))

        self.btn_abrir_resultado = ctk.CTkButton(
            acciones_bottom,
            text="Abrir resultado final",
            state="disabled",
            command=self.abrir_resultado_final
        )
        self.btn_abrir_resultado.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_abrir_carpeta = ctk.CTkButton(
            acciones_bottom,
            text="Abrir carpeta de salida",
            state="disabled",
            command=self.abrir_carpeta_salida
        )
        self.btn_abrir_carpeta.pack(side="left", fill="x", expand=True, padx=(8, 8))

        self.btn_abrir_no_parseados = ctk.CTkButton(
            acciones_bottom,
            text="Abrir TXT no parseados",
            state="disabled",
            command=self.abrir_no_parseados
        )
        self.btn_abrir_no_parseados.pack(side="left", fill="x", expand=True, padx=(8, 0))

        self.status_bar = StatusBar(self.main_container)
        self.status_bar.pack(fill="x", pady=(10, 10))

    def _configurar_dnd(self):
        if not DND_DISPONIBLE:
            self.drop_pdf.subtitle_label.configure(
                text="tkinterdnd2 no instalado. Usa Buscar PDF."
            )
            self.drop_excel.subtitle_label.configure(
                text="tkinterdnd2 no instalado. Usa Buscar Excel."
            )
            return

        try:
            self.drop_pdf.drop_target_register(DND_FILES)
            self.drop_pdf.dnd_bind("<<Drop>>", self.al_soltar_pdf)
            self.drop_pdf.dnd_bind("<<DragEnter>>", lambda e: self.drop_pdf.set_hover(True))
            self.drop_pdf.dnd_bind("<<DragLeave>>", lambda e: self.drop_pdf.set_hover(False))

            self.drop_excel.drop_target_register(DND_FILES)
            self.drop_excel.dnd_bind("<<Drop>>", self.al_soltar_excel)
            self.drop_excel.dnd_bind("<<DragEnter>>", lambda e: self.drop_excel.set_hover(True))
            self.drop_excel.dnd_bind("<<DragLeave>>", lambda e: self.drop_excel.set_hover(False))
        except Exception:
            self.drop_pdf.subtitle_label.configure(
                text="Drag and drop no disponible. Usa Buscar PDF."
            )
            self.drop_excel.subtitle_label.configure(
                text="Drag and drop no disponible. Usa Buscar Excel."
            )

    def _normalizar_ruta_dnd(self, ruta: str) -> str:
        ruta = ruta.strip()
        if ruta.startswith("{") and ruta.endswith("}"):
            ruta = ruta[1:-1]
        return ruta

    def seleccionar_pdf(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            initialdir=str(self.ultima_carpeta),
            filetypes=[
                ("Archivos PDF", "*.pdf"),
                ("Todos los archivos", "*.*")
            ]
        )
        if ruta:
            self.establecer_pdf(ruta)

    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            initialdir=str(self.ultima_carpeta),
            filetypes=[
                ("Archivos Excel", "*.xls *.xlsx"),
                ("Todos los archivos", "*.*")
            ]
        )
        if ruta:
            self.establecer_excel(ruta)

    def abrir_explorador(self):
        abrir_ruta(self.ultima_carpeta)

    def _limpiar_resultado_ui(self):
        self.resultado = None
        self.btn_abrir_resultado.configure(state="disabled")
        self.btn_abrir_carpeta.configure(state="disabled")
        self.btn_abrir_no_parseados.configure(state="disabled")

    def establecer_pdf(self, ruta: str):
        ruta = self._normalizar_ruta_dnd(ruta)
        ruta_path = Path(ruta)

        if ruta_path.suffix.lower() != ".pdf":
            messagebox.showwarning("Archivo no valido", "Debes seleccionar un archivo PDF.")
            return

        self.ruta_pdf.set(str(ruta_path))
        self.ultima_carpeta = ruta_path.parent
        self.drop_pdf.set_success(ruta_path.name)
        self.status_bar.set_status("PDF seleccionado correctamente", "success")
        self._limpiar_resultado_ui()

    def establecer_excel(self, ruta: str):
        ruta = self._normalizar_ruta_dnd(ruta)
        ruta_path = Path(ruta)

        if ruta_path.suffix.lower() not in [".xls", ".xlsx"]:
            messagebox.showwarning("Archivo no valido", "Debes seleccionar un archivo Excel.")
            return

        self.ruta_excel.set(str(ruta_path))
        self.ultima_carpeta = ruta_path.parent
        self.drop_excel.set_success(ruta_path.name)
        self.status_bar.set_status("Excel seleccionado correctamente", "success")
        self._limpiar_resultado_ui()

    def al_soltar_pdf(self, event):
        self.drop_pdf.set_hover(False)

        try:
            rutas = self.tk.splitlist(event.data)
            if rutas:
                self.establecer_pdf(rutas[0])
        except Exception:
            messagebox.showerror("Error", "No se pudo interpretar el PDF arrastrado.")

    def al_soltar_excel(self, event):
        self.drop_excel.set_hover(False)

        try:
            rutas = self.tk.splitlist(event.data)
            if rutas:
                self.establecer_excel(rutas[0])
        except Exception:
            messagebox.showerror("Error", "No se pudo interpretar el Excel arrastrado.")

    def _guardar_bloques_no_parseados(self, bloques, ruta_txt: Path):
        with ruta_txt.open("w", encoding="utf-8") as f:
            for i, bloque in enumerate(bloques, 1):
                f.write(f"===== BLOQUE NO PARSEADO {i} =====\n")
                for linea in bloque:
                    f.write(str(linea) + "\n")
                f.write("\n")

    def procesar(self):
        ruta_pdf = self.ruta_pdf.get().strip()
        ruta_excel = self.ruta_excel.get().strip()

        if not ruta_pdf or not ruta_excel:
            messagebox.showwarning(
                "Faltan archivos",
                "Debes seleccionar un PDF y un Excel antes de procesar."
            )
            return

        try:
            self.status_bar.set_status("Procesando PDF + Excel... por favor espera", "processing")
            self.btn_procesar.configure(state="disabled", text="Procesando...")
            self.update_idletasks()

            resultado = procesar_pdf_excel(ruta_pdf, ruta_excel)

            claves_esperadas = ["df_final", "df_pdf", "df_excel", "no_parseados"]
            faltantes = [k for k in claves_esperadas if k not in resultado]
            if faltantes:
                raise ValueError(
                    f"El servicio no devolvio las claves esperadas: {faltantes}. "
                    f"Claves recibidas: {list(resultado.keys())}"
                )

            nombre_base = resultado.get("nombre_base_pdf", "resultado_final")
            ruta_guardado = filedialog.asksaveasfilename(
                title="Guardar resultado final",
                initialdir=str(self.ultima_carpeta),
                initialfile=f"{nombre_base}_resultado_final.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Archivo Excel", "*.xlsx")]
            )

            if not ruta_guardado:
                self.btn_procesar.configure(state="normal", text="Procesar PDF + Excel")
                self.status_bar.set_status("Guardado cancelado por el usuario", "waiting")
                return

            ruta_guardado = Path(ruta_guardado)
            carpeta_salida = ruta_guardado.parent
            self.ultima_carpeta = carpeta_salida

            df_final = resultado["df_final"]
            df_pdf = resultado["df_pdf"]
            df_excel = resultado["df_excel"]
            no_parseados = resultado["no_parseados"]

            # Guardado principal con formato Excel pro
            aplicar_formato_excel(df_final, ruta_guardado)

            ruta_pdf_csv = carpeta_salida / f"{ruta_guardado.stem}_pdf_extraido.csv"
            ruta_excel_csv = carpeta_salida / f"{ruta_guardado.stem}_excel_procesado.csv"
            ruta_txt = carpeta_salida / f"{ruta_guardado.stem}_no_parseados.txt"

            df_pdf.to_csv(ruta_pdf_csv, index=False, encoding="utf-8-sig")
            df_excel.to_csv(ruta_excel_csv, index=False, encoding="utf-8-sig")

            if no_parseados:
                self._guardar_bloques_no_parseados(no_parseados, ruta_txt)
                ruta_txt_final = ruta_txt
            else:
                ruta_txt_final = None

            self.resultado = {
                "archivo_final": str(ruta_guardado),
                "csv_pdf": str(ruta_pdf_csv),
                "csv_excel": str(ruta_excel_csv),
                "txt_no_parseados": str(ruta_txt_final) if ruta_txt_final else "",
                "estadisticas": resultado.get("estadisticas", {})
            }

            self.btn_procesar.configure(state="normal", text="Procesar PDF + Excel")
            self.btn_abrir_resultado.configure(state="normal")
            self.btn_abrir_carpeta.configure(state="normal")

            if ruta_txt_final and ruta_txt_final.exists():
                self.btn_abrir_no_parseados.configure(state="normal")
            else:
                self.btn_abrir_no_parseados.configure(state="disabled")

            self.status_bar.set_status(f"Completado: {ruta_guardado.name}", "success")

            abrir_ahora = messagebox.askyesno(
                "Proceso completado",
                f"Archivo generado correctamente:\n\n{ruta_guardado}\n\n¿Deseas abrirlo ahora?"
            )

            if abrir_ahora:
                self.abrir_resultado_final()

        except Exception as e:
            self.resultado = None
            self.btn_procesar.configure(state="normal", text="Procesar PDF + Excel")
            self.btn_abrir_resultado.configure(state="disabled")
            self.btn_abrir_carpeta.configure(state="disabled")
            self.btn_abrir_no_parseados.configure(state="disabled")
            self.status_bar.set_status("Error al procesar los archivos", "error")
            messagebox.showerror("Error", f"Ocurrio un error al procesar los archivos:\n\n{e}")
            print(traceback.format_exc())

    def abrir_resultado_final(self):
        if self.resultado and self.resultado.get("archivo_final"):
            ruta = Path(self.resultado["archivo_final"])
            if ruta.exists():
                abrir_ruta(ruta)
                return
        messagebox.showwarning("Archivo no disponible", "No se encontro el resultado final.")

    def abrir_carpeta_salida(self):
        if self.resultado and self.resultado.get("archivo_final"):
            ruta = Path(self.resultado["archivo_final"])
            if ruta.exists():
                abrir_ruta(ruta.parent)
                return
        messagebox.showwarning("Carpeta no disponible", "No se encontro la carpeta de salida.")

    def abrir_no_parseados(self):
        if self.resultado and self.resultado.get("txt_no_parseados"):
            ruta = Path(self.resultado["txt_no_parseados"])
            if ruta.exists():
                abrir_ruta(ruta)
                return
        messagebox.showwarning("Archivo no disponible", "No se encontro el TXT de no parseados.")
