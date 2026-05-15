"""
main.py  –  Applikationens indgangspunkt

Start programmet med:   python main.py

For at tilføje et nyt værktøj:
  1. Opret en mappe under tools/ med en view.py
  2. Tilføj det i config/tools_registry.py  ← her er den eneste ændring nødvendig
"""

import customtkinter as ctk
from components.sidebar import Sidebar
from config.tools_registry import TOOLS

# ── Globale udseende-indstillinger ────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")


class App(ctk.CTk):
    """Hoved-vindue med sidebar (venstre) og indholdsområde (højre)."""

    def __init__(self):
        super().__init__()
        self.title("Dataværktøjer")
        self.geometry("1150x720")
        self.minsize(960, 620)

        self._setup_layout()
        self._load_tool(0)  # Åbn første værktøj ved opstart

    def _setup_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar (venstre panel)
        self.sidebar = Sidebar(self, tools=TOOLS, on_select=self._load_tool)
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        # Indholdsområde (højre panel)
        self.content_frame = ctk.CTkFrame(self, fg_color="#F7F7F5", corner_radius=0)
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self._current_view = None

    def _load_tool(self, index: int):
        """Erstat indholdet med det valgte værktøjs view."""
        if self._current_view:
            self._current_view.destroy()

        tool_class = TOOLS[index]["view"]
        self._current_view = tool_class(self.content_frame)
        self._current_view.grid(row=0, column=0, sticky="nsew", padx=36, pady=28)

        self.sidebar.set_active(index)


if __name__ == "__main__":
    app = App()
    app.mainloop()
