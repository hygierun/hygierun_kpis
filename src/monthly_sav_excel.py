"""Excel de sortie de la section SAV : synthèse chiffrée + feuilles de détail (4 blocs implémentés)."""

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook

from .excel_helpers import select_columns, write_frame, write_summary
from .monthly_loader import MONTHS_FR
from .monthly_sav import SavCalculator


def sav_rows(result: dict) -> List[Optional[list]]:
    """Lignes Synthèse propres au SAV (devis, facturation vs docs nuls, main d'œuvre, déplacement,
    nombre d'interventions, productivité)."""
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"], d["n1"], d["m1"]
    m1_interventions = (result.get("mois1_reference") or {}).get("nb_interventions", {})

    def line(label: str, block: str, field: str, evo_key: Optional[str] = None) -> list:
        return [label, cur[block][field], n1[block][field], e.get(f"{evo_key}_n1") if evo_key else None,
                m1[block][field], e.get(f"{evo_key}_m1") if evo_key else None]

    def line_simple(label: str, value) -> list:
        return [label, value, None, None, None, None]

    def line_interventions(label: str, champ: str) -> list:
        return [label, cur["nb_interventions"][champ], None, None,
                m1_interventions.get(champ), e.get(f"nb_interventions_{champ}_m1")]

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
        None,
        ["Nombre d'interventions (feuille Bilan_Fiches, fiches papier)"],
        line_interventions("Total", "total"),
        line_interventions("  dont EBC", "ebc"),
        line_interventions("  dont SAV", "sav"),
        None,
        ["Productivité (= heures en intervention / heures travaillées)"],
        line_simple("  Heures travaillées EBC", cur["productivite"]["heures_travaillees"]["ebc"]),
        line_simple("  Heures travaillées SAV", cur["productivite"]["heures_travaillees"]["sav"]),
        line_simple("  Heures intervention EBC", cur["productivite"]["heures_intervention"]["ebc"]),
        line_simple("  Heures intervention SAV", cur["productivite"]["heures_intervention"]["sav"]),
        line_simple("  Productivité EBC (%)", cur["productivite"]["pct"]["ebc"]),
        line_simple("  Productivité SAV (%)", cur["productivite"]["pct"]["sav"]),
        line_simple("  Productivité totale (%)", cur["productivite"]["pct"]["total"]),
    ]


def achat_rows(result: dict) -> List[Optional[list]]:
    """Lignes Synthèse propres à l'Achat/Appro (valorisation du stock, articles à épuisement)."""
    cur = result["data"]["cur"]

    def line_simple(label: str, value) -> list:
        return [label, value, None, None, None, None]

    return [
        ["Valorisation du stock (feuille Stock_Invent, Achat/Appro)"],
        *[line_simple(f"  {row['famille']}", row["total"]) for row in cur["valorisation_stock"]["rows"]],
        line_simple("  TOTAL", cur["valorisation_stock"]["total"]),
        None,
        ["Articles à épuisement (feuille Ach_Arti, Réappro = 'A épuis.', Achat/Appro)"],
        line_simple("Total", cur["articles_epuisement"]["total"]),
        line_simple("  dont vendus sur l'année (%)", cur["articles_epuisement"]["pct_ventes"]),
        line_simple("  dont stockés / dispo compta (%)", cur["articles_epuisement"]["pct_dispo"]),
    ]


def summary_rows(result: dict) -> List[Optional[list]]:
    """Lignes de la feuille Synthèse : [libellé, valeur, N-1, évol N-1, Mois-1, évol Mois-1]."""
    return sav_rows(result) + [None] + achat_rows(result)


def write_sav_sheets(wb: Workbook, calc: SavCalculator, year: int, month: int):
    write_frame(wb, "Devis", select_columns(calc.devis_rows(year, month), ["Date", "N°", "Client", "Salarié", "Total HT"]))
    write_frame(wb, "Facture SAV", select_columns(calc.docs_nuls_rows(year, month),
                ["Catégorie", "Date", "N°", "Client", "Salarié", "Représentant", "PosteEtats", "Type Doc Nul", "Total HT"]))
    write_frame(wb, "Main_oeuvre", select_columns(calc.main_oeuvre_rows(year, month), ["Représentant", "Article réf", "Article", "Catégorie", "Heures"]))
    write_frame(wb, "Deplacement", select_columns(calc.deplacement_rows(year, month), ["Représentant", "Article réf", "Article", "Catégorie", "Heures"]))


def write_achat_sheets(wb: Workbook, calc: SavCalculator, year: int, month: int):
    stock_rows = calc.valorisation_stock_rows(year, month)
    col_t = next(c for c in stock_rows.columns if isinstance(c, str) and c.startswith("T "))
    write_frame(wb, "Valorisation_Stock", select_columns(stock_rows, ["Dépot", "Désignation", "Référence", "Famille", col_t]))

    write_frame(wb, "Articles_Epuisement", select_columns(calc.articles_epuisement_rows(year, month),
                ["Fournisseur", "Désignation", "Référence", "Famille", "Réappro", "Dispo", f"Vtes {year}"]))


def generate_sav_excel(calc: SavCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"SAV - {MONTHS_FR[month - 1]} {year}", summary_rows(result))

    write_sav_sheets(wb, calc, year, month)
    write_achat_sheets(wb, calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
