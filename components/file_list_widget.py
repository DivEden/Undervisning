"""
file_list_widget.py  –  Rullelistefelt med valgte filer

Viser en scrollbar liste af filer med:
  - Filtype-badge (DOCX = blå, PDF = orange)
  - Filnavn
  - Fjern-knap pr. fil
"""

import os
import customtkinter as ctk
from theme.colors import (
    CARD_BG, INPUT_BG, BORDER, TEXT_PRIMARY, TEXT_MUTED, WHITE,
    ACCENT,
)


class FileListWidget(ctk.CTkScrollableFrame):
    """
    Rullelistefelt til visning og styring af valgte filer.

    Brug:
        widget = FileListWidget(parent, height=300)
        widget.set_on_change(callback)    # callback(files: list[str])
        widget.add_files(["/sti/til/fil.docx"])
        widget.get_files()                # -> list[str]
        widget.clear()
    """

    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            fg_color=CARD_BG,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=TEXT_MUTED,
            corner_radius=8,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)
        self._files: list[str] = []
        self._rows: list[ctk.CTkFrame] = []
        self._on_change = None

    def set_on_change(self, callback):
        """Registrer en funktion der kaldes hver gang fillisten ændrer sig."""
        self._on_change = callback

    def add_files(self, paths: list[str]):
        """Tilføj filer til listen (ignorerer dubletter)."""
        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self._add_row(p)
        if self._on_change:
            self._on_change(self._files)

    def get_files(self) -> list[str]:
        """Returnér en kopi af den aktuelle filliste."""
        return list(self._files)

    def clear(self):
        """Fjern alle filer fra listen."""
        for row in self._rows:
            row.destroy()
        self._rows.clear()
        self._files.clear()
        if self._on_change:
            self._on_change(self._files)

    # ── Interne metoder ────────────────────────────────────────────────────

    def _add_row(self, path: str):
        row = ctk.CTkFrame(self, fg_color=INPUT_BG, corner_radius=6)
        row.grid(row=len(self._rows), column=0, sticky="ew", padx=6, pady=3)
        row.grid_columnconfigure(1, weight=1)

        # Filtype-badge
        ext = os.path.splitext(path)[1].upper().lstrip(".")
        badge_color = "#3B82F6" if ext == "DOCX" else ACCENT  # Blå=DOCX, Orange=PDF
        ctk.CTkLabel(
            row,
            text=ext or "?",
            fg_color=badge_color,
            text_color=WHITE,
            corner_radius=4,
            width=44,
            font=ctk.CTkFont(size=10, weight="bold"),
        ).grid(row=0, column=0, padx=(8, 6), pady=8)

        # Filnavn
        ctk.CTkLabel(
            row,
            text=os.path.basename(path),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=1, sticky="w", padx=4, pady=8)

        # Fjern-knap
        ctk.CTkButton(
            row,
            text="✕",
            width=28,
            height=28,
            fg_color="transparent",
            hover_color=BORDER,
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=12),
            command=lambda p=path, r=row: self._remove(p, r),
        ).grid(row=0, column=2, padx=(4, 8), pady=8)

        self._rows.append(row)

    def _remove(self, path: str, row: ctk.CTkFrame):
        row.destroy()
        idx = self._rows.index(row)
        self._rows.pop(idx)
        self._files.remove(path)
        # Genopdater grid-rækker
        for i, r in enumerate(self._rows):
            r.grid(row=i)
        if self._on_change:
            self._on_change(self._files)
