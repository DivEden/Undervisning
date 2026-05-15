"""
progress_panel.py  –  Fremdriftslinje + logoutput

Bruges i trin 3 (Kør import) til at vise:
  - Progressbar med procent
  - Rullende log med beskeder fra baggrundstråden
"""

import customtkinter as ctk
from theme.colors import CARD_BG, INPUT_BG, BORDER, PRIMARY, TEXT_PRIMARY, TEXT_SECONDARY


class ProgressPanel(ctk.CTkFrame):
    """
    Kombination af progressbar og scrollbar logboks.

    Brug:
        panel = ProgressPanel(parent, height=320)
        panel.set_progress(0.5, "Behandler fil 5/10")
        panel.log("Fil indlæst: rapport.docx")
        panel.clear()
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color=CARD_BG, corner_radius=10, **kwargs)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Progressbar
        self._bar = ctk.CTkProgressBar(
            self,
            fg_color=BORDER,
            progress_color=PRIMARY,
            height=8,
            corner_radius=4,
        )
        self._bar.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 4))
        self._bar.set(0)

        # Procentlabel
        self._pct_lbl = ctk.CTkLabel(
            self,
            text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_SECONDARY,
            anchor="e",
        )
        self._pct_lbl.grid(row=1, column=0, sticky="e", padx=16, pady=(0, 6))

        # Logboks (read-only)
        self._log_box = ctk.CTkTextbox(
            self,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=6,
            state="disabled",
            wrap="none",
        )
        self._log_box.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))

    def set_progress(self, value: float, text: str = ""):
        """
        Opdater fremdrift.
        value: float mellem 0.0 og 1.0
        text:  valgfri statusbesked vist ved siden af procenten
        """
        self._bar.set(value)
        pct = int(value * 100)
        self._pct_lbl.configure(text=f"{pct}%  {text}".strip())

    def log(self, message: str):
        """Tilføj en linje til logboksen og scroll til bunden."""
        self._log_box.configure(state="normal")
        self._log_box.insert("end", message + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def clear(self):
        """Nulstil progressbar og ryd logboksen."""
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")
        self._bar.set(0)
        self._pct_lbl.configure(text="0%")
