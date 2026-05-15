# ─────────────────────────────────────────────────────────────────────────────
# tools_registry.py  –  Registrer nye værktøjer her
#
# Hvert element i TOOLS-listen skal have:
#   name         – Visningsnavn i sidebar
#   description  – Kort undertekst vist under navnet
#   icon         – Unicode-emoji eller symbol vist i sidebar
#   view         – En CTkFrame-underklasse der udgør værktøjets UI
#
# For at tilføje et nyt værktøj:
#   1. Opret en ny mappe under tools/
#   2. Lav en view.py med en CTkFrame-underklasse
#   3. Tilføj en ny dict her i TOOLS-listen
# ─────────────────────────────────────────────────────────────────────────────

from tools.excel_import.view import ExcelImportView
from tools.statistics_maker.view import StatisticsMakerView
from tools.template_maker.view import TemplateMakerView

TOOLS = [
    {
        "name": "Excel Import",
        "description": "Importer data til Excel-ark",
        "icon": "📥",
        "view": ExcelImportView,
    },
    {
        "name": "Statistik Maker",
        "description": "Byg statistik-tabeller i Excel",
        "icon": "🔧",
        "view": StatisticsMakerView,
    },
    {
        "name": "Skabelon Maker",
        "description": "Opret nyt årsark + Data-ark",
        "icon": "🧩",
        "view": TemplateMakerView,
    },
]
