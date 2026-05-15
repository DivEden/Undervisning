import customtkinter as ctk
from theme.colors import (
    SIDEBAR_BG, SIDEBAR_HOVER, PRIMARY, TEXT_ON_DARK,
    TEXT_MUTED, TEXT_SIDEBAR, WHITE,
)


class Sidebar(ctk.CTkFrame):
    """
    Venstre navigationspanel (PowerToys-stil).

    For at ændre udseendet:
      - Bredde:       skift `width=` i __init__
      - Titeltekst:   skift `text=` i CTkLabel nedenfor
      - Farver:       se theme/colors.py
    """

    def __init__(self, parent, tools: list, on_select):
        super().__init__(parent, width=230, fg_color=SIDEBAR_BG, corner_radius=0)
        self.grid_propagate(False)
        self.grid_columnconfigure(0, weight=1)

        self._buttons: list[_NavButton] = []
        self._on_select = on_select

        # ── App-titel ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Dataværktøjer",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=TEXT_ON_DARK,
        ).grid(row=0, column=0, padx=20, pady=(28, 4), sticky="w")

        ctk.CTkLabel(
            self,
            text="Interne hjælpeværktøjer",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        # Skillelinje
        ctk.CTkFrame(self, height=1, fg_color=SIDEBAR_HOVER).grid(
            row=2, column=0, sticky="ew", padx=16, pady=(0, 12)
        )

        # ── Navknapper ─────────────────────────────────────────────────────
        for i, tool in enumerate(tools):
            btn = _NavButton(
                self,
                icon=tool.get("icon", "•"),
                label=tool["name"],
                description=tool.get("description", ""),
                command=lambda idx=i: on_select(idx),
            )
            btn.grid(row=i + 3, column=0, padx=12, pady=2, sticky="ew")
            self._buttons.append(btn)

        # Fyld-række (skubber footer ned)
        self.grid_rowconfigure(len(tools) + 3, weight=1)

        # Versionsnummer i bunden
        ctk.CTkLabel(
            self,
            text="v1.0.0",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=TEXT_MUTED,
        ).grid(row=len(tools) + 4, column=0, padx=20, pady=16, sticky="w")

    def set_active(self, index: int):
        """Marker én knap som aktiv og resten som inaktive."""
        for i, btn in enumerate(self._buttons):
            btn.set_active(i == index)


class _NavButton(ctk.CTkFrame):
    """Enkelt navigationselement med ikon, navn og beskrivelse."""

    def __init__(self, parent, icon: str, label: str, description: str, command):
        super().__init__(parent, fg_color="transparent", corner_radius=8, cursor="hand2")
        self.grid_columnconfigure(1, weight=1)
        self._command = command
        self._active = False

        # Ikon
        self._icon_lbl = ctk.CTkLabel(
            self,
            text=icon,
            font=ctk.CTkFont(size=18),
            text_color=TEXT_SIDEBAR,
            width=32,
        )
        self._icon_lbl.grid(row=0, column=0, rowspan=2, padx=(12, 6), pady=10)

        # Navn
        self._main_lbl = ctk.CTkLabel(
            self,
            text=label,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_SIDEBAR,
            anchor="w",
        )
        self._main_lbl.grid(row=0, column=1, padx=(0, 12), pady=(10, 0), sticky="w")

        # Beskrivelse
        self._desc_lbl = ctk.CTkLabel(
            self,
            text=description,
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self._desc_lbl.grid(row=1, column=1, padx=(0, 12), pady=(0, 10), sticky="w")

        # Bind events til alle børne-widgets
        for widget in (self, self._icon_lbl, self._main_lbl, self._desc_lbl):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

    def _on_click(self, _event=None):
        if self._command:
            self._command()

    def _on_enter(self, _event=None):
        if not self._active:
            self.configure(fg_color=SIDEBAR_HOVER)

    def _on_leave(self, _event=None):
        if not self._active:
            self.configure(fg_color="transparent")

    def set_active(self, active: bool):
        self._active = active
        self.configure(fg_color=PRIMARY if active else "transparent")
        color = WHITE if active else TEXT_SIDEBAR
        weight = "bold" if active else "normal"
        self._icon_lbl.configure(text_color=color)
        self._main_lbl.configure(
            text_color=color,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight=weight),
        )
