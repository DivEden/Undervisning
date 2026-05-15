"""
view.py  –  Placeholder for kommende værktøj

Erstat dette med det rigtige værktøjs UI, når det skal bygges.

Fremgangsmåde:
  1. Omdøb denne mappe til noget sigende (fx tools/data_report/)
  2. Erstat denne view.py med det faktiske UI
  3. Opdater tools_registry.py med det nye navn og view-klasse
"""

import customtkinter as ctk
from theme.colors import CARD_BG, TEXT_PRIMARY, TEXT_SECONDARY


class PlaceholderView(ctk.CTkFrame):
    """Visningswidget for endnu ikke implementeret værktøj."""

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        card = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.grid(row=0, column=0)

        ctk.CTkLabel(
            inner,
            text="🔧",
            font=ctk.CTkFont(size=52),
        ).grid(row=0, column=0, pady=(0, 16))

        ctk.CTkLabel(
            inner,
            text="Kommende værktøj",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=1, column=0)

        ctk.CTkLabel(
            inner,
            text="Dette værktøj er endnu ikke bygget.\nKom tilbage snart!",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_SECONDARY,
            justify="center",
        ).grid(row=2, column=0, pady=(10, 0))
