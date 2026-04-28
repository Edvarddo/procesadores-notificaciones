import customtkinter as ctk


class ProcessorCard(ctk.CTkFrame):
    def __init__(self, master, title: str, description: str, button_text: str, command, **kwargs):
        super().__init__(master, corner_radius=16, **kwargs)

        self.configure(fg_color=("#f8fafc", "#1f2937"))

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        self.title_label = ctk.CTkLabel(
            content,
            text=title,
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w"
        )
        self.title_label.pack(anchor="w", pady=(0, 10))

        self.description_label = ctk.CTkLabel(
            content,
            text=description,
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="w",
            wraplength=320
        )
        self.description_label.pack(anchor="w", fill="x", pady=(0, 20))

        self.open_button = ctk.CTkButton(
            content,
            text=button_text,
            command=command,
            height=40,
            corner_radius=10
        )
        self.open_button.pack(anchor="w")
