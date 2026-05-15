"""
buttons.py  –  Genanvendelige knap-komponenter

Ændre knap-stil globalt ved at redigere klasserne herunder.
Brug disse klasser i stedet for rå ctk.CTkButton overalt i appen.
"""

import customtkinter as ctk
from theme.colors import (
    PRIMARY, PRIMARY_HOVER, WHITE,
    INPUT_BG, BORDER, TEXT_PRIMARY,
    ERROR,
)


class PrimaryButton(ctk.CTkButton):
    """Primær handlingsknap – fyldt teal. Bruges til 'Næste', 'Start import' osv."""

    def __init__(self, parent, text: str, command=None, **kwargs):
        super().__init__(
            parent,
            text=text,
            command=command,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color=WHITE,
            corner_radius=8,
            height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            **kwargs,
        )


class SecondaryButton(ctk.CTkButton):
    """Sekundær knap – let grå baggrund. Bruges til 'Tilbage', 'Gennemse' osv."""

    def __init__(self, parent, text: str, command=None, **kwargs):
        super().__init__(
            parent,
            text=text,
            command=command,
            fg_color=INPUT_BG,
            hover_color=BORDER,
            text_color=TEXT_PRIMARY,
            border_color=BORDER,
            border_width=1,
            corner_radius=8,
            height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            **kwargs,
        )


class DangerButton(ctk.CTkButton):
    """Destruktiv handlingsknap – rød. Bruges til slet/nulstil-handlinger."""

    def __init__(self, parent, text: str, command=None, **kwargs):
        super().__init__(
            parent,
            text=text,
            command=command,
            fg_color=ERROR,
            hover_color="#B91C1C",
            text_color=WHITE,
            corner_radius=8,
            height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            **kwargs,
        )
