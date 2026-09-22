"""Excel de sortie de la section SAV : synthèse chiffrée + feuilles de détail (4 blocs implémentés)."""

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook

from .excel_helpers import select_columns, write_frame, write_summary
from .monthly_loader import MONTHS_FR
from .monthly_sav import SavCalculator


def summary_rows(result: dict) -> List[Optional[list]]:
    """Lignes de la feuille Synthèse : [libellé, valeur, N-1, évol N-1, Mois-1, évol Mois-1]."""
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"], d["n1"], d["m1"]

    def line(label: str, block: str, field: str, evo_key: Optional[str] = None) -> list:
        return [label, cur[block][field], n1[block][field], e.get(f"{evo_key}_n1") if evo_key else None,
                m1[block][field], e.get(f"{evo_key}_m1") if evo_key else None]

    return [
        ["Devis SAV (Salarié Jimmy DALLEAU + Joël DANVIN)"],
        line("Nb devis", "devis", "nb", "devis_nb"),
        line("CA HT devis (€)", "devis", "ca"),
        None,
        ["Facturation vs Docs Nuls (Fact_Arch)"],
        line("Nb factures total", "docs_nuls", "total", "docs_nuls_total"),
        line("  dont Docs Nuls (DN)", "docs_nuls", "dn", "docs_nuls_dn"),
        line("  dont hors Docs Nuls", "docs_nuls", "non_dn", "docs_nuls_non_dn"),
        line("CA HT (€)", "docs_nuls", "ca", "docs_nuls_ca"),
        None,
        ["Main d'œuvre (feuille MO+Depl, heures)"],
        line("Total heures", "main_oeuvre", "total", "main_oeuvre_total"),
        line("  dont EBC", "main_oeuvre", "ebc", "main_oeuvre_ebc"),
        line("  dont SAV", "main_oeuvre", "sav", "main_oeuvre_sav"),
        None,
        ["Déplacement (feuille MO+Depl, unités)"],
        line("Total", "deplacement", "total", "deplacement_total"),
        line("  dont EBC", "deplacement", "ebc", "deplacement_ebc"),
        line("  dont SAV", "deplacement", "sav", "deplacement_sav"),
    ]


def generate_sav_excel(calc: SavCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"SAV - {MONTHS_FR[month - 1]} {year}", summary_rows(result))

    write_frame(wb, "Devis", select_columns(calc.devis_rows(year, month), ["Date", "N°", "Client", "Salarié", "Total HT"]))
    write_frame(wb, "Docs_Nuls", select_columns(calc.docs_nuls_rows(year, month),
                ["Catégorie", "Date", "N°", "Client", "Salarié", "Représentant", "PosteEtats", "Type Doc Nul", "Total HT"]))
    write_frame(wb, "Main_oeuvre", select_columns(calc.main_oeuvre_rows(year, month), ["Représentant", "Article réf", "Article", "Catégorie", "Heures"]))
    write_frame(wb, "Deplacement", select_columns(calc.deplacement_rows(year, month), ["Représentant", "Article réf", "Article", "Catégorie", "Heures"]))

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
