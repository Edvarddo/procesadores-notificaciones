import customtkinter as ctk
from ui.widgets.processor_card import ProcessorCard


class HomeView(ctk.CTkFrame):
    def __init__(
            self,
            master,
            on_open_reporte,
            on_open_pdf_excel,
            on_open_avisos,
            on_open_bitacora,
            on_open_carabineros,
            **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.on_open_bitacora = on_open_bitacora
        self.on_open_carabineros = on_open_carabineros
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            header,
            text="Menu principal",
            font=ctk.CTkFont(size=28, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            header,
            text="Selecciona el procesador que quieres usar",
            font=ctk.CTkFont(size=14),
            text_color=("#6b7280", "#9ca3af")
        ).pack(anchor="w", pady=(5, 0))

        grid = ctk.CTkFrame(self, fg_color="transparent")
        grid.pack(fill="both", expand=True)

        grid.grid_rowconfigure(0, weight=1)
        grid.grid_rowconfigure(1, weight=1)
        grid.grid_rowconfigure(2, weight=1)

        card1 = ProcessorCard(
            grid,
            title="Procesador hojas de ruta",
            description="Procesa archivos Excel del reporte de notificaciones certificadas. Luego de haberlas ingresado por el CENTRO DE NOTIFICACIONES.",
            button_text="Abrir procesador",
            command=on_open_reporte
        )
        card1.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)

        card2 = ProcessorCard(
            grid,
            title="Procesador hojas de ruta",
            description="Compara las notificaciones que realmente salieron a terreno con las que fueron impresas. Pueden quedar algunos regitros/notificaciones vacios debido a que no todas las que se imprimen salen a terreno. Además, siempre se quitan o añaden un poco de notificaciones de otros dias.",
            button_text="Abrir modulo",
            command=on_open_pdf_excel
        )
        card2.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)

        card3 = ProcessorCard(
            grid,
            title="Procesador de avisos",
            description="Toma el archivo Detalle de Impresion, filtra registros Personal/Art44 y rellena la plantilla de avisos lista para imprimir.",
            button_text="Abrir modulo",
            command=on_open_avisos
        )
        card3.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=10)

        card4 = ProcessorCard(
            grid,
            title="Procesador de bitacora de audiencias",
            description="Genera la bitacora organizada por salas, con paginacion, sombreado de causas repetidas y formato listo para impresion.",
            button_text="Abrir modulo",
            command=self.on_open_bitacora
        )
        card4.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=10)
        
        card5 = ProcessorCard(
            grid,
            title="Automatización formulario - Carabineros",
            description="Carga un archivo CSV y ejecuta la automatización de ingreso en CINJ para el flujo de Carabineros.",
            button_text="Abrir modulo",
            command=self.on_open_carabineros
        )
        card5.grid(row=2, column=0, sticky="nsew", padx=(0, 10), pady=10)
