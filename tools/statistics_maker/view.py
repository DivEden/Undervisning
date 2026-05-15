import os
import threading
import tkinter.filedialog as fd
import tkinter.messagebox as mb

import customtkinter as ctk

from components.buttons import PrimaryButton, SecondaryButton
from theme.colors import CARD_BG, INPUT_BG, BORDER, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_MUTED
from .logic import (
    GROUP_OPTIONS,
    METRIC_OPTIONS,
    TEMPLATE_OPTIONS,
    list_data_sheets,
    list_target_sheets,
    generate_statistics,
    generate_template,
    preview_statistics,
    preview_template,
)


class StatisticsMakerView(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._excel_path: str | None = None
        self._source_var = ctk.StringVar(value="")
        self._target_var = ctk.StringVar(value="")
        self._group_var = ctk.StringVar(value="year")
        self._metric_var = ctk.StringVar(value="count_rows")
        self._template_var = ctk.StringVar(value="")

        self._source_map: dict[str, str] = {}
        self._target_map: dict[str, str] = {}
        self._group_map: dict[str, str] = {label: key for key, label in GROUP_OPTIONS.items()}
        self._metric_map: dict[str, str] = {label: key for key, label in METRIC_OPTIONS.items()}
        self._template_map: dict[str, str] = {"Ingen template (enkelt statistik)": "__NONE__"}
        for key, label in TEMPLATE_OPTIONS.items():
            self._template_map[label] = key

        self._build_header()
        self._build_body()

    def _build_header(self):
        ctk.CTkLabel(
            self,
            text="Statistik Maker",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            self,
            text="Generér statistik-tabeller direkte fra data-ark og indsæt dem i fx Data 2026.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_SECONDARY,
        ).grid(row=0, column=0, sticky="w", pady=(34, 0))

    def _build_body(self):
        card = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(9, weight=1)

        file_row = ctk.CTkFrame(card, fg_color=INPUT_BG, corner_radius=8)
        file_row.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        file_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            file_row,
            text="Excel-fil",
            width=110,
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

        self._source_cb = self._build_combo_row(card, 1, "Kilde", self._source_var, ["Vælg Excel-fil først"])
        self._target_cb = self._build_combo_row(card, 2, "Målark", self._target_var, ["Vælg Excel-fil først"])

        group_labels = list(GROUP_OPTIONS.values())
        metric_labels = list(METRIC_OPTIONS.values())

        self._group_cb = self._build_combo_row(card, 3, "Gruppér efter", self._group_var, group_labels, value=GROUP_OPTIONS["year"])
        self._metric_cb = self._build_combo_row(
            card,
            4,
            "Måling",
            self._metric_var,
            metric_labels,
            value=METRIC_OPTIONS["count_rows"],
            command=self._on_metric_changed,
        )

        template_labels = list(self._template_map.keys())
        self._template_cb = self._build_combo_row(
            card,
            5,
            "Template",
            self._template_var,
            template_labels,
            value=template_labels[0],
            command=self._on_template_changed,
        )

        hint = (
            "Eksempler:\n"
            "- Antal skoler per år: Gruppér efter = År, Måling = Antal unikke skoler\n"
            "- Antal specialklasser per måned: Gruppér efter = Måned, Måling = Antal specialklasse\n"
            "- Komplet månedlig årssammenligning: Måling = Månedlig sammenligning (Besøg + Elever pr. år)\n"
            "- Vælg en Template for at få færdige multi-kolonne tabeller\n"
            "- Alt der starter med 'Template EKSTRA' er nye, udvidede analyser (udover de klassiske)\n"
            "Hvis samme statistik køres igen, erstattes den gamle blok i målarket i stedet for at blive lagt oven i."
        )
        ctk.CTkLabel(
            card,
            text=hint,
            justify="left",
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_SECONDARY,
        ).grid(row=6, column=0, sticky="ew", padx=20, pady=(8, 8))

        action_row = ctk.CTkFrame(card, fg_color="transparent")
        action_row.grid(row=7, column=0, sticky="ew", padx=20, pady=(0, 10))

        self._preview_btn = SecondaryButton(action_row, text="Forhåndsvis", command=self._preview)
        self._preview_btn.pack(side="right", padx=(0, 8))

        self._run_btn = PrimaryButton(action_row, text="▶ Generér statistik", command=self._run)
        self._run_btn.pack(side="right")

        self._preview_area = ctk.CTkScrollableFrame(
            card,
            fg_color=INPUT_BG,
            border_color=BORDER,
            border_width=1,
            corner_radius=8,
            height=220,
        )
        self._preview_area.grid(row=8, column=0, sticky="nsew", padx=20, pady=(0, 10))
        self._preview_area.grid_columnconfigure(0, weight=1)

        self._preview_placeholder = ctk.CTkLabel(
            self._preview_area,
            text="Forhåndsvisning vises her, før du skriver til Excel.",
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._preview_placeholder.grid(row=0, column=0, sticky="w", padx=12, pady=12)

        self._log = ctk.CTkTextbox(card, fg_color=INPUT_BG, border_color=BORDER, border_width=1, corner_radius=8)
        self._log.grid(row=9, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self._log.insert("1.0", "Klar. Vælg Excel-fil for at starte.\n")
        self._log.configure(state="disabled")

    def _build_combo_row(self, parent, row, label, var, values, value=None, command=None):
        frame = ctk.CTkFrame(parent, fg_color=INPUT_BG, corner_radius=8)
        frame.grid(row=row, column=0, sticky="ew", padx=20, pady=(0, 10))
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame,
            text=label,
            width=110,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")

        cb = ctk.CTkComboBox(
            frame,
            variable=var,
            values=values,
            state="readonly",
            command=command,
            fg_color=CARD_BG,
            border_color=BORDER,
            button_color="#0F766E",
            dropdown_hover_color=BORDER,
            width=420,
        )
        cb.grid(row=0, column=1, sticky="w", pady=8, padx=(0, 12))
        if value is not None:
            cb.set(value)
        else:
            cb.set(values[0])
        return cb

    def _on_metric_changed(self, _value=None):
        metric_key = self._metric_map.get(self._metric_var.get())
        template_key = self._template_map.get(self._template_var.get(), "__NONE__")
        if template_key != "__NONE__":
            return

        if metric_key == "matrix_monthly_year_visits_students":
            self._group_var.set(GROUP_OPTIONS["month"])
            self._group_cb.configure(state="disabled")
        else:
            self._group_cb.configure(state="readonly")

    def _on_template_changed(self, _value=None):
        template_key = self._template_map.get(self._template_var.get(), "__NONE__")
        if template_key == "__NONE__":
            self._metric_cb.configure(state="readonly")
            self._on_metric_changed()
        else:
            self._group_cb.configure(state="disabled")
            self._metric_cb.configure(state="disabled")

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
        self._load_sheet_options()

    def _load_sheet_options(self):
        if not self._excel_path:
            return
        try:
            data_sheets = list_data_sheets(self._excel_path)
            target_sheets = list_target_sheets(self._excel_path)
        except Exception as e:
            mb.showerror("Fejl", f"Kunne ikke læse Excel-fil:\n{e}")
            return

        if not data_sheets:
            mb.showwarning("Ingen data-ark", "Fandt ingen data-ark med den forventede skabelon.")
            return

        self._source_map = {"Alle data-ark": "__ALL_DATA_SHEETS__"}
        for s in data_sheets:
            self._source_map[s] = s
        source_labels = list(self._source_map.keys())
        self._source_cb.configure(values=source_labels)
        self._source_cb.set(source_labels[0])

        self._target_map = {s: s for s in target_sheets}
        target_labels = list(self._target_map.keys())
        self._target_cb.configure(values=target_labels)

        preferred = "Data 2026" if "Data 2026" in self._target_map else target_labels[0]
        self._target_cb.set(preferred)

        self._log_line(f"Fundet {len(data_sheets)} data-ark og {len(target_sheets)} målark.")

    def _run(self):
        if not self._excel_path:
            mb.showwarning("Mangler fil", "Vælg en Excel-fil først.")
            return

        source_label = self._source_var.get()
        target_label = self._target_var.get()
        template_label = self._template_var.get()
        template_key = self._template_map.get(template_label, "__NONE__")
        group_label = self._group_var.get()
        metric_label = self._metric_var.get()

        if source_label not in self._source_map or target_label not in self._target_map:
            mb.showwarning("Mangler valg", "Vælg både kilde og målark.")
            return

        if template_key == "__NONE__":
            group_key = self._group_map.get(group_label)
            metric_key = self._metric_map.get(metric_label)
            if not group_key or not metric_key:
                mb.showwarning("Mangler valg", "Vælg både gruppering og måling.")
                return
        else:
            group_key = None
            metric_key = None

        self._run_btn.configure(state="disabled", text="Genererer...")
        if template_key == "__NONE__":
            self._log_line(f"Starter: {metric_label} pr. {group_label}")
            thread = threading.Thread(
                target=self._worker,
                args=(self._excel_path, self._source_map[source_label], self._target_map[target_label], group_key, metric_key),
                daemon=True,
            )
        else:
            self._log_line(f"Starter template: {template_label}")
            thread = threading.Thread(
                target=self._template_worker,
                args=(self._excel_path, self._source_map[source_label], self._target_map[target_label], template_key),
                daemon=True,
            )
        thread.start()

    def _clear_preview_area(self):
        for child in self._preview_area.winfo_children():
            child.destroy()

    def _render_preview_error(self, text: str):
        self._clear_preview_area()
        ctk.CTkLabel(
            self._preview_area,
            text=text,
            text_color="#DC2626",
            anchor="w",
            justify="left",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=12)

    def _render_preview_blocks(self, blocks: list[dict]):
        self._clear_preview_area()

        if not blocks:
            self._render_preview_error("Ingen data i forhåndsvisning.")
            return

        row_index = 0
        for block in blocks:
            title = block.get("title", "Forhåndsvisning")
            headers = block.get("headers", [])
            rows = block.get("rows", [])

            section = ctk.CTkFrame(self._preview_area, fg_color=CARD_BG, corner_radius=8)
            section.grid(row=row_index, column=0, sticky="ew", padx=10, pady=(10, 6))
            section.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                section,
                text=title,
                text_color=TEXT_PRIMARY,
                anchor="w",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 6))

            table = ctk.CTkFrame(section, fg_color="transparent")
            table.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

            max_cols = len(headers)
            for c in range(max_cols):
                table.grid_columnconfigure(c, weight=1)

            # Header row
            for c, h in enumerate(headers):
                ctk.CTkLabel(
                    table,
                    text=str(h),
                    fg_color="#0F766E",
                    text_color="#FFFFFF",
                    corner_radius=4,
                    anchor="w" if c == 0 else "e",
                    padx=8,
                    font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                ).grid(row=0, column=c, sticky="ew", padx=(0 if c == 0 else 4, 0), pady=(0, 4))

            # Data rows
            for r, row_vals in enumerate(rows, start=1):
                row_is_total = str(row_vals[0]).strip().upper() in {"TOTAL", "HOVEDTOTAL"}
                bg = "#E5E7EB" if row_is_total else ("#F8FAFC" if r % 2 == 0 else "#FFFFFF")
                for c, val in enumerate(row_vals):
                    txt = "" if val is None else str(val)
                    ctk.CTkLabel(
                        table,
                        text=txt,
                        fg_color=bg,
                        text_color=TEXT_PRIMARY,
                        corner_radius=4,
                        anchor="w" if c == 0 else "e",
                        padx=8,
                        font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold" if row_is_total else "normal"),
                    ).grid(row=r, column=c, sticky="ew", padx=(0 if c == 0 else 4, 0), pady=(0, 3))

            row_index += 1

    def _preview(self):
        if not self._excel_path:
            mb.showwarning("Mangler fil", "Vælg en Excel-fil først.")
            return

        source_label = self._source_var.get()
        template_label = self._template_var.get()
        template_key = self._template_map.get(template_label, "__NONE__")
        group_label = self._group_var.get()
        metric_label = self._metric_var.get()

        if source_label not in self._source_map:
            mb.showwarning("Mangler valg", "Vælg en kilde.")
            return

        if template_key == "__NONE__":
            group_key = self._group_map.get(group_label)
            metric_key = self._metric_map.get(metric_label)
            if not group_key or not metric_key:
                mb.showwarning("Mangler valg", "Vælg både gruppering og måling.")
                return
        else:
            group_key = None
            metric_key = None

        self._preview_btn.configure(state="disabled", text="Forhåndsviser...")
        if template_key == "__NONE__":
            self._log_line(f"Forhåndsviser: {metric_label} pr. {group_label}")
            thread = threading.Thread(
                target=self._preview_worker,
                args=(self._excel_path, self._source_map[source_label], group_key, metric_key),
                daemon=True,
            )
        else:
            self._log_line(f"Forhåndsviser template: {template_label}")
            thread = threading.Thread(
                target=self._preview_template_worker,
                args=(self._excel_path, self._source_map[source_label], template_key),
                daemon=True,
            )
        thread.start()

    def _preview_worker(self, excel_path, source_choice, group_by, metric):
        try:
            headers, rows = preview_statistics(
                excel_path=excel_path,
                source_choice=source_choice,
                group_by=group_by,
                metric=metric,
                max_rows=16,
            )
            blocks = [{"title": "Forhåndsvisning", "headers": headers, "rows": rows}]
            self.after(0, lambda: self._render_preview_blocks(blocks))
        except Exception as e:
            self.after(0, lambda: self._render_preview_error(f"Forhåndsvisning fejlede:\n{e}"))
            self.after(0, lambda: self._log_line(f"Forhåndsvisning fejl: {e}"))
        finally:
            self.after(0, lambda: self._preview_btn.configure(state="normal", text="Forhåndsvis"))

    def _preview_template_worker(self, excel_path, source_choice, template_key):
        try:
            blocks = preview_template(
                excel_path=excel_path,
                source_choice=source_choice,
                template_id=template_key,
                max_rows_per_block=6,
            )
            self.after(0, lambda: self._render_preview_blocks(blocks))
        except Exception as e:
            self.after(0, lambda: self._render_preview_error(f"Forhåndsvisning fejlede:\n{e}"))
            self.after(0, lambda: self._log_line(f"Forhåndsvisning fejl: {e}"))
        finally:
            self.after(0, lambda: self._preview_btn.configure(state="normal", text="Forhåndsvis"))

    def _template_worker(self, excel_path, source_choice, target_sheet, template_key):
        try:
            rows, title = generate_template(
                excel_path=excel_path,
                source_choice=source_choice,
                target_sheet=target_sheet,
                template_id=template_key,
            )
            self.after(0, lambda: self._log_line(f"Færdig template: {title} ({rows} rækker) -> {target_sheet}"))
            self.after(0, lambda: mb.showinfo("Færdig", f"Template oprettet i ark: {target_sheet}"))
        except Exception as e:
            self.after(0, lambda: self._log_line(f"Fejl: {e}"))
            self.after(0, lambda: mb.showerror("Fejl", str(e)))
        finally:
            self.after(0, lambda: self._run_btn.configure(state="normal", text="▶ Generér statistik"))

    def _worker(self, excel_path, source_choice, target_sheet, group_by, metric):
        try:
            rows, title = generate_statistics(
                excel_path=excel_path,
                source_choice=source_choice,
                target_sheet=target_sheet,
                group_by=group_by,
                metric=metric,
            )
            self.after(0, lambda: self._log_line(f"Færdig: {title} ({rows} rækker) -> {target_sheet}"))
            self.after(0, lambda: mb.showinfo("Færdig", f"Statistik oprettet i ark: {target_sheet}"))
        except Exception as e:
            self.after(0, lambda: self._log_line(f"Fejl: {e}"))
            self.after(0, lambda: mb.showerror("Fejl", str(e)))
        finally:
            self.after(0, lambda: self._run_btn.configure(state="normal", text="▶ Generér statistik"))
