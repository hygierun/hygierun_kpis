"""Contrôles du loader mensuel : une seule ligne d'en-têtes par feuille, colonnes numériques saines.

Lancer avec : python -m pytest tests  (ou python tests/test_monthly_loader.py)
"""

import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.monthly_loader import MonthlyDataLoader  # noqa: E402

REQUIRED = {"Clients": ["Désignation", "Qualification", "Solde cpta", "Vtes 2026"]}
HEADER = ["Désignation", "Qualification", "Solde cpta", "Vtes 2026"]


def _load(rows):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "input.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Clients"
        for row in rows:
            ws.append(row)
        wb.save(path)
        loader = MonthlyDataLoader(str(path))
        ok = loader._load(REQUIRED)
        return ok, loader


def test_single_header_is_accepted():
    ok, loader = _load([HEADER, ["A", "Client", 10.0, 100.0], ["B", "Client", 5.0, 0.0]])
    assert ok, loader.errors
    assert loader.dfs["Clients"]["Vtes 2026"].sum() == 100.0


def test_stacked_headers_are_reported_with_the_sheet_name():
    ok, loader = _load([HEADER, ["Désignation", "Référence", "Qualification", "Solde cpta", "Vtes 2026"],
                        ["A", 1, "Client", 10.0, 100.0]])
    assert not ok
    assert any("Feuille 'Clients'" in e and "plusieurs lignes d'en-têtes" in e and "1 et 2" in e
               for e in loader.errors), loader.errors


def test_text_in_numeric_column_is_reported_with_the_sheet_name():
    rows = [HEADER] + [["A", "Client", "texte", "texte"] for _ in range(10)]
    ok, loader = _load(rows)
    assert not ok
    assert any("Feuille 'Clients'" in e and "contient du texte" in e for e in loader.errors), loader.errors


def test_isolated_bad_cell_is_tolerated():
    rows = [HEADER] + [["A", "Client", 1.0, 10.0] for _ in range(20)] + [["B", "Client", 1.0, "x"]]
    ok, loader = _load(rows)
    assert ok, loader.errors
    assert loader.dfs["Clients"]["Vtes 2026"].sum() == 200.0


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
