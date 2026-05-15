"""
step_header.py  –  Visuel 3-trins fremdriftsindikator

Viser hvilke trin der er gennemført (grøn), aktivt (teal) og kommende (grå).
Bruges øverst i Excel Import-wizarden.

For at ændre antal trin eller tekster:
    Redigér STEPS-listen i klassen.
"""

import customtkinter as ctk
from theme.colors import PRIMARY, SUCCESS, BORDER, TEXT_MUTED, WHITE


class StepHeader(ctk.CTkFrame):
    """
    Vandret trin-indikator med nummererede cirkler og forbindelseslinjer.

    Brug:
        header = StepHeader(parent)
        header.set_step(0)   # 0-indekseret
    """

    # Ændres her hvis antal trin eller navne skal opdateres
    STEPS = ["Vælg data", "Vælg Excel", "Kør import"]

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._step = 0
        self._dots: list[ctk.CTkLabel] = []
        self._labels: list[ctk.CTkLabel] = []
        self._build()

    def _build(self):
        for i, step_text in enumerate(self.STEPS):
            col = i * 2

            # Talcirkel
            dot = ctk.CTkLabel(
                self,
                text=str(i + 1),
                width=30,
                height=30,
                corner_radius=15,
                fg_color=BORDER,
                text_color=TEXT_MUTED,
                font=ctk.CTkFont(size=12, weight="bold"),
            )
            dot.grid(row=0, column=col, padx=6)
            self._dots.append(dot)

            # Trinnavn
            lbl = ctk.CTkLabel(
                self,
                text=step_text,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=TEXT_MUTED,
            )
            lbl.grid(row=1, column=col, padx=6)
            self._labels.append(lbl)

            # Forbindelseslinje
            if i < len(self.STEPS) - 1:
                ctk.CTkFrame(self, height=2, width=50, fg_color=BORDER).grid(
                    row=0, column=col + 1, padx=2
                )

        self._refresh()

    def set_step(self, step: int):
        """Opdater aktiv trin (0-indekseret)."""
        self._step = step
        self._refresh()

    def _refresh(self):
        for i, (dot, lbl) in enumerate(zip(self._dots, self._labels)):
            if i < self._step:
                # Gennemført
                dot.configure(fg_color=SUCCESS, text_color=WHITE)
                lbl.configure(
                    text_color=SUCCESS,
                    font=ctk.CTkFont(family="Segoe UI", size=11),
                )
            elif i == self._step:
                # Aktivt
                dot.configure(fg_color=PRIMARY, text_color=WHITE)
                lbl.configure(
                    text_color=PRIMARY,
                    font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                )
            else:
                # Kommende
                dot.configure(fg_color=BORDER, text_color=TEXT_MUTED)
                lbl.configure(
                    text_color=TEXT_MUTED,
                    font=ctk.CTkFont(family="Segoe UI", size=11),
                )
