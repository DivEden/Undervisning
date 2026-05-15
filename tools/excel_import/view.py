"""
view.py  –  Excel Import-værktøjets UI (3-trins wizard)

Trin 0 – Vælg data:    Tilføj mappe eller individuelle filer
Trin 1 – Vælg Excel:   Vælg målfil og ark
Trin 2 – Kør import:   Se fremgang og log i realtid

For at tilpasse:
  - Filtyper der accepteres:   se _pick_folder() og _pick_files()
  - Import-logik:              se tools/excel_import/logic.py
  - Komponent-stil:            se components/ og theme/colors.py
"""

import os
import threading
import tkinter.filedialog as fd
import tkinter.messagebox as mb

import customtkinter as ctk

from components.buttons import PrimaryButton, SecondaryButton
from components.file_list_widget import FileListWidget
from components.step_header import StepHeader
from components.progress_panel import ProgressPanel
from theme.colors import (
    CARD_BG, INPUT_BG, BORDER, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_MUTED,
    INFO_BG, INFO_TEXT,
)
from .logic import extract_file, write_to_excel, list_importable_sheets


class ExcelImportView(ctk.CTkFrame):
    """Hoved-widget for Excel Import-værktøjet."""

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._step = 0
        self._excel_path: str | None = None
        self._sheet_name: str | None = None
        self._sheet_var = ctk.StringVar(value="")
        self._importable_sheets: list[str] = []

        self._build_header()
        self._build_step_panels()
        self._show_step(0)

    # ─── Øverste header med titel og trin-indikator ────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hdr,
            text="Excel Import",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            hdr,
            text="Importer data fra Word- og PDF-filer ind i et Excel-ark",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_SECONDARY,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        self._step_header = StepHeader(hdr)
        self._step_header.grid(row=0, column=1, rowspan=2, padx=(20, 0), sticky="e")

    # ─── Byg alle trin-paneler ─────────────────────────────────────────────

    def _build_step_panels(self):
        self._panels: list[ctk.CTkFrame] = [
            self._build_step0(),
            self._build_step1(),
            self._build_step2(),
        ]

    # ── Trin 0: Vælg datafiler ─────────────────────────────────────────────

    def _build_step0(self) -> ctk.CTkFrame:
        panel = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            panel,
            text="Vælg datafiler",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 4))

        ctk.CTkLabel(
            panel,
            text="Tilføj en hel mappe eller vælg individuelle .docx- og .pdf-filer.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_SECONDARY,
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))

        # Knaprækkefølge
        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 8))

        SecondaryButton(btn_row, text="📁  Vælg mappe", command=self._pick_folder).pack(
            side="left", padx=(0, 8)
        )
        SecondaryButton(btn_row, text="📄  Vælg filer", command=self._pick_files).pack(
            side="left", padx=(0, 8)
        )
        SecondaryButton(btn_row, text="Ryd liste", command=self._clear_files).pack(side="left")

        # Filliste
        self._file_list = FileListWidget(panel, height=280)
        self._file_list.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 8))
        self._file_list.set_on_change(self._on_files_changed)

        # Antal-label
        self._file_count_lbl = ctk.CTkLabel(
            panel,
            text="Ingen filer valgt",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MUTED,
        )
        self._file_count_lbl.grid(row=4, column=0, sticky="w", padx=20, pady=(0, 8))

        # Navigation
        nav = ctk.CTkFrame(panel, fg_color="transparent")
        nav.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))

        self._next_btn_0 = PrimaryButton(nav, text="Næste  →", command=lambda: self._go_to(1))
        self._next_btn_0.pack(side="right")
        self._next_btn_0.configure(state="disabled")

        return panel

    # ── Trin 1: Vælg Excel-fil og ark ─────────────────────────────────────

    def _build_step1(self) -> ctk.CTkFrame:
        panel = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            panel,
            text="Vælg Excel-mål",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 4))

        ctk.CTkLabel(
            panel,
            text="Vælg den Excel-fil og det ark du vil importere data ind i.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_SECONDARY,
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 20))

        ctk.CTkLabel(
            panel,
            text="Kun ark med dataskabelon vises (fx 2026). Pivot-/Data-ark skjules automatisk.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MUTED,
        ).grid(row=2, column=0, sticky="w", padx=20, pady=(0, 12))

        # Excel-filvælger
        excel_row = ctk.CTkFrame(panel, fg_color=INPUT_BG, corner_radius=8)
        excel_row.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 12))
        excel_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            excel_row,
            text="Excel-fil",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXT_PRIMARY,
            width=90,
        ).grid(row=0, column=0, padx=(16, 8), pady=14, sticky="w")

        self._excel_path_lbl = ctk.CTkLabel(
            excel_row,
            text="Ingen fil valgt",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self._excel_path_lbl.grid(row=0, column=1, sticky="w", pady=14)

        SecondaryButton(
            excel_row, text="Gennemse…", command=self._pick_excel, width=110
        ).grid(row=0, column=2, padx=(8, 16), pady=8)

        # Ark-dropdown
        sheet_row = ctk.CTkFrame(panel, fg_color=INPUT_BG, corner_radius=8)
        sheet_row.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        sheet_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            sheet_row,
            text="Ark",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXT_PRIMARY,
            width=90,
        ).grid(row=0, column=0, padx=(16, 8), pady=14, sticky="w")

        self._sheet_dropdown = ctk.CTkComboBox(
            sheet_row,
            variable=self._sheet_var,
            values=["Vælg Excel-fil først"],
            command=self._on_sheet_changed,
            state="disabled",
            width=260,
            fg_color=CARD_BG,
            border_color=BORDER,
            button_color="#0F766E",
            dropdown_hover_color=BORDER,
        )
        self._sheet_dropdown.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=10)

        # Opsummerings-kort
        self._summary_card = ctk.CTkFrame(panel, fg_color=INFO_BG, corner_radius=8)
        self._summary_card.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))
        self._summary_card.grid_columnconfigure(0, weight=1)

        self._summary_lbl = ctk.CTkLabel(
            self._summary_card,
            text="Vælg en Excel-fil og et ark for at fortsætte.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=INFO_TEXT,
            justify="left",
            anchor="w",
        )
        self._summary_lbl.grid(row=0, column=0, padx=16, pady=12, sticky="w")

        # Navigation
        nav = ctk.CTkFrame(panel, fg_color="transparent")
        nav.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 20))

        SecondaryButton(nav, text="← Tilbage", command=lambda: self._go_to(0)).pack(side="left")

        self._next_btn_1 = PrimaryButton(
            nav, text="Gå til import  →", command=lambda: self._go_to(2)
        )
        self._next_btn_1.pack(side="right")
        self._next_btn_1.configure(state="disabled")

        return panel

    # ── Trin 2: Kør import ─────────────────────────────────────────────────

    def _build_step2(self) -> ctk.CTkFrame:
        panel = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            panel,
            text="Kør import",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 4))

        ctk.CTkLabel(
            panel,
            text="Klik Start for at importere data. Fremgangen vises herunder.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXT_SECONDARY,
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))

        self._progress_panel = ProgressPanel(panel, height=300)
        self._progress_panel.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 12))

        # Navigation
        nav = ctk.CTkFrame(panel, fg_color="transparent")
        nav.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 20))

        self._back_btn_2 = SecondaryButton(nav, text="← Tilbage", command=lambda: self._go_to(1))
        self._back_btn_2.pack(side="left")

        self._run_btn = PrimaryButton(nav, text="▶  Start import", command=self._run_import)
        self._run_btn.pack(side="right")

        return panel

    # ─── Trin-navigation ──────────────────────────────────────────────────

    def _show_step(self, step: int):
        for i, panel in enumerate(self._panels):
            if i == step:
                panel.grid(row=2, column=0, sticky="nsew")
            else:
                panel.grid_remove()
        self._step_header.set_step(step)
        self._step = step

    def _go_to(self, step: int):
        if step == 2:
            # Nulstil import-panelet ved hvert besøg
            self._progress_panel.clear()
            self._run_btn.configure(state="normal", text="▶  Start import")
            self._back_btn_2.configure(state="normal")
        self._show_step(step)

    # ─── Filvalg ──────────────────────────────────────────────────────────

    def _pick_folder(self):
        folder = fd.askdirectory(title="Vælg mappe med datafiler")
        if not folder:
            return
        files = sorted([
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith((".docx", ".pdf"))
        ])
        self._file_list.add_files(files)

    def _pick_files(self):
        files = fd.askopenfilenames(
            title="Vælg datafiler",
            filetypes=[("Datafiler", "*.docx *.pdf"), ("Alle filer", "*.*")],
        )
        if files:
            self._file_list.add_files(list(files))

    def _clear_files(self):
        self._file_list.clear()

    def _on_files_changed(self, files: list[str]):
        count = len(files)
        if count == 0:
            self._file_count_lbl.configure(text="Ingen filer valgt")
            self._next_btn_0.configure(state="disabled")
        else:
            label = f"{count} fil{'er' if count > 1 else ''} valgt"
            self._file_count_lbl.configure(text=label)
            self._next_btn_0.configure(state="normal")

    # ─── Excel-valg ───────────────────────────────────────────────────────

    def _pick_excel(self):
        path = fd.askopenfilename(
            title="Vælg Excel-fil",
            filetypes=[("Excel-filer", "*.xlsx *.xlsm"), ("Alle filer", "*.*")],
        )
        if not path:
            return
        try:
            sheets = list_importable_sheets(path)
        except Exception as e:
            mb.showerror("Fejl", f"Kunne ikke åbne Excel-filen:\n{e}")
            return

        if not sheets:
            mb.showwarning(
                "Ingen data-ark fundet",
                "Filen indeholder ingen ark med korrekt datastruktur. "
                "Vælg en fil med årsark (fx 2026)."
            )
            return

        self._excel_path = path
        self._importable_sheets = sheets
        self._excel_path_lbl.configure(
            text=os.path.basename(path), text_color=TEXT_PRIMARY
        )
        self._sheet_dropdown.configure(values=sheets, state="normal")
        self._sheet_var.set(sheets[0] if sheets else "")
        self._on_sheet_changed(self._sheet_var.get())

    def _on_sheet_changed(self, value: str):
        self._sheet_name = value if value in self._importable_sheets else None
        self._update_summary()
        self._validate_step1()

    def _update_summary(self):
        n = len(self._file_list.get_files())
        xl = os.path.basename(self._excel_path) if self._excel_path else "–"
        sheet = self._sheet_name or "–"
        self._summary_lbl.configure(
            text=f"ℹ  Klar til at importere {n} fil{'er' if n != 1 else ''} → {xl}  ›  {sheet}"
        )

    def _validate_step1(self):
        ok = bool(self._excel_path and self._sheet_name)
        self._next_btn_1.configure(state="normal" if ok else "disabled")

    # ─── Import (kører i baggrundstråd) ───────────────────────────────────

    def _run_import(self):
        files = self._file_list.get_files()
        if not files:
            mb.showwarning("Ingen filer", "Tilføj mindst én datafil i trin 1.")
            return

        self._run_btn.configure(state="disabled", text="Behandler…")
        self._back_btn_2.configure(state="disabled")
        self._progress_panel.clear()

        threading.Thread(target=self._import_worker, daemon=True).start()

    def _import_worker(self):
        files = self._file_list.get_files()
        total_files = len(files)
        all_rows: list[dict] = []
        skipped: list[str] = []

        self._log(f"Starter import af {total_files} fil(er)…\n")

        for i, path in enumerate(files):
            fname = os.path.basename(path)
            self._log(f"Læser [{i + 1}/{total_files}]: {fname}")
            rows, error = extract_file(path)
            if error:
                self._log(f"  ⚠  {error}")
                skipped.append(fname)
            else:
                self._log(f"  ✓  {len(rows)} rækker udtrukket")
                all_rows.extend(rows)
            self._set_progress((i + 1) / total_files * 0.7, f"Fil {i + 1}/{total_files}")

        self._log(f"\nSamlet: {len(all_rows)} rækker fra {total_files - len(skipped)} filer.")

        if not all_rows:
            self._log("\n❌ Ingen data at skrive – import afbrudt.")
            self.after(0, lambda: self._run_btn.configure(state="normal", text="▶  Start import"))
            self.after(0, lambda: self._back_btn_2.configure(state="normal"))
            return

        self._log(f"\nSkriver til: {os.path.basename(self._excel_path)}  ›  {self._sheet_name}")

        try:
            written, warnings = write_to_excel(
                all_rows,
                self._excel_path,
                self._sheet_name,
                progress_callback=lambda v, t: self._set_progress(0.7 + v * 0.3, t),
            )
            for w in warnings:
                self._log(f"  ⚠  {w}")
            self._log(f"\n✅ Import fuldført! {written} rækker skrevet til '{self._sheet_name}'.")
        except Exception as e:
            self._log(f"\n❌ Fejl under skrivning til Excel:\n   {e}")

        if skipped:
            self._log(f"\n⚠  {len(skipped)} fil(er) sprunget over:")
            for f in skipped:
                self._log(f"   • {f}")

        self._set_progress(1.0, "Færdig")
        self.after(0, lambda: self._run_btn.configure(state="normal", text="▶  Kør igen"))
        self.after(0, lambda: self._back_btn_2.configure(state="normal"))

    # ─── Tråd-sikre UI-hjælpere ───────────────────────────────────────────

    def _log(self, message: str):
        self.after(0, lambda: self._progress_panel.log(message))

    def _set_progress(self, value: float, text: str = ""):
        self.after(0, lambda: self._progress_panel.set_progress(value, text))
