import os
import threading
import tkinter.filedialog as fd
import tkinter.messagebox as mb

import customtkinter as ctk

from components.buttons import PrimaryButton, SecondaryButton
from theme.colors import BORDER, CARD_BG, INPUT_BG, TEXT_MUTED, TEXT_PRIMARY, TEXT_SECONDARY
from .logic import create_year_template, suggest_next_year


class TemplateMakerView(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._excel_path: str | None = None

        self._build_header()
        self._build_body()

    def _build_header(self):
        ctk.CTkLabel(
            self,
            text="Skabelon Maker",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            self,
            text="Opret nyt årsark og tilhørende Data-ark, fx 2027 + Data 2027.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_SECONDARY,
        ).grid(row=0, column=0, sticky="w", pady=(34, 0))

    def _build_body(self):
        card = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(4, weight=1)

        file_row = ctk.CTkFrame(card, fg_color=INPUT_BG, corner_radius=8)
        file_row.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        file_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            file_row,
            text="Excel-fil",
            width=120,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")

        self._excel_lbl = ctk.CTkLabel(
            file_row,
            text="Ingen fil valgt",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self._excel_lbl.grid(row=0, column=1, sticky="w")

        SecondaryButton(file_row, text="Gennemse…", command=self._pick_excel, width=120).grid(
            row=0, column=2, padx=(8, 12), pady=8
        )

        year_row = ctk.CTkFrame(card, fg_color=INPUT_BG, corner_radius=8)
        year_row.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

        ctk.CTkLabel(
            year_row,
            text="Nyt år",
            width=120,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")

        self._year_entry = ctk.CTkEntry(
            year_row,
            width=140,
            placeholder_text="fx 2027",
            fg_color=CARD_BG,
            border_color=BORDER,
        )
        self._year_entry.grid(row=0, column=1, sticky="w", pady=8)

        PrimaryButton(card, text="▶ Opret års-skabelon", command=self._run).grid(
            row=2, column=0, sticky="e", padx=20, pady=(4, 10)
        )

        ctk.CTkLabel(
            card,
            text=(
                "Resultat:\n"
                "- Nyt årsark med korrekt kolonneopsætning og tabel-stil\n"
                "- Nyt Data-ark klar til statistik\n"
                "- Eksisterende ark overskrives ikke"
            ),
            justify="left",
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_SECONDARY,
        ).grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._log = ctk.CTkTextbox(card, fg_color=INPUT_BG, border_color=BORDER, border_width=1, corner_radius=8)
        self._log.grid(row=4, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self._log.insert("1.0", "Klar. Vælg Excel-fil og skriv årstal.\n")
        self._log.configure(state="disabled")

    def _log_line(self, text: str):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _pick_excel(self):
        path = fd.askopenfilename(
            title="Vælg Excel-fil",
            filetypes=[("Excel", "*.xlsx"), ("Alle filer", "*.*")],
        )
        if not path:
            return

        self._excel_path = path
        self._excel_lbl.configure(text=os.path.basename(path), text_color=TEXT_PRIMARY)
        self._log_line(f"Valgt fil: {path}")

        try:
            nxt = suggest_next_year(path)
            self._year_entry.delete(0, "end")
            self._year_entry.insert(0, str(nxt))
            self._log_line(f"Foreslået næste år: {nxt}")
        except Exception as e:
            self._log_line(f"Kunne ikke foreslå år automatisk: {e}")

    def _run(self):
        if not self._excel_path:
            mb.showwarning("Mangler fil", "Vælg en Excel-fil først.")
            return

        raw = self._year_entry.get().strip()
        if not raw.isdigit():
            mb.showwarning("Ugyldigt år", "Skriv et årstal, fx 2027.")
            return

        year = int(raw)
        self._log_line(f"Starter oprettelse for år {year}...")

        thread = threading.Thread(target=self._worker, args=(self._excel_path, year), daemon=True)
        thread.start()

    def _worker(self, excel_path: str, year: int):
        try:
            year_sheet, data_sheet = create_year_template(excel_path, year)
            self.after(0, lambda: self._log_line(f"Færdig: {year_sheet} og {data_sheet} oprettet."))
            self.after(0, lambda: mb.showinfo("Færdig", f"Oprettet: {year_sheet}\nOprettet: {data_sheet}"))
        except Exception as e:
            self.after(0, lambda: self._log_line(f"Fejl: {e}"))
            self.after(0, lambda: mb.showerror("Fejl", str(e)))
