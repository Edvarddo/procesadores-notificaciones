import customtkinter as ctk
from tkinterdnd2 import TkinterDnD

from ui.views.home_view import HomeView
from ui.views.reporte_view import ReporteView
from ui.views.pdf_excel_view import PdfExcelView
from ui.views.avisos_view import AvisosView
from ui.views.bitacora_view import BitacoraView
from ui.views.carabineros_view import CarabinerosView


class MainWindow(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        ctk.CTk.__init__(self)

        # Cargar soporte DnD en la ventana raíz
        self.TkdndVersion = TkinterDnD._require(self)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.title("Suite de Procesadores")
        self.geometry("1000x760")
        self.minsize(900, 650)

        self.current_view = None

        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        self.show_home()

    def clear_view(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_view()
        self.current_view = HomeView(
            self.container,
            on_open_reporte=self.show_reporte_view,
            on_open_pdf_excel=self.show_pdf_excel_view,
            on_open_avisos=self.show_avisos_view,
            on_open_bitacora=self.show_bitacora_view,
            on_open_carabineros=self.show_carabineros_view,
        )
        self.current_view.pack(fill="both", expand=True)

    def show_reporte_view(self):
        self.clear_view()
        self.current_view = ReporteView(
            self.container,
            on_back=self.show_home
        )
        self.current_view.pack(fill="both", expand=True)

    def show_pdf_excel_view(self):
        self.clear_view()
        self.current_view = PdfExcelView(
            self.container,
            on_back=self.show_home
        )
        self.current_view.pack(fill="both", expand=True)

    def show_avisos_view(self):
        self.clear_view()
        self.current_view = AvisosView(
            self.container,
            on_back=self.show_home
        )
        self.current_view.pack(fill="both", expand=True)

    def show_bitacora_view(self):
        self.clear_view()
        self.current_view = BitacoraView(
            self.container,
            volver_callback=self.show_home
        )
        self.current_view.pack(fill="both", expand=True)
    def show_carabineros_view(self):
        self.clear_view()
        self.current_view = CarabinerosView(
            self.container,
            volver_callback=self.show_home
        )
        self.current_view.pack(fill="both", expand=True)


def main():
    app = MainWindow()
    app.mainloop()
