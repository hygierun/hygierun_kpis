"""Orchestration du Bilan Mensuel consolidé : un seul fichier Input, un seul PowerPoint de sortie.

Réutilise tel quel le chargement (MonthlyDataLoader.load_bilan_total) et les 4 calculators de section
existants - ce module ne fait qu'assembler, aucune règle de calcul n'est dupliquée ou modifiée.

Mois-1 : lu depuis le PowerPoint Bilan Mensuel consolidé du mois précédent (optionnel). Sans lui, les
évolutions Mois-1 qui en dépendent (clients bloqués, factures dues, GPS, nombre d'interventions,
productivité, valorisation du stock, articles à épuisement) restent 'n/a', sans crash.
"""

import tempfile
from pathlib import Path
from typing import Dict, Optional

from . import monthly_commerce_excel, monthly_livraison_excel, monthly_preparation_excel, monthly_sav_excel
from .monthly_bilan_mois1 import read_bilan_mois1_reference
from .monthly_bilan_pptx import generate_bilan_pptx
from .monthly_commerce import CommerceCalculator
from .monthly_livraison import LivraisonComptaCalculator
from .monthly_loader import MonthlyDataLoader
from .monthly_preparation import PreparationCalculator
from .monthly_sav import SavCalculator

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
BILAN_TEMPLATE = TEMPLATES_DIR / "KPI_Mensuel.pptx"


def build_bilan_report(input_path: str, year: int, month: int, template_path: Path = BILAN_TEMPLATE,
                       mois1_pptx_path: Optional[str] = None) -> Dict:
    loader = MonthlyDataLoader(input_path)
    if not loader.load_bilan_total(year, month):
        raise ValueError("\n".join(loader.errors))

    mois1 = read_bilan_mois1_reference(mois1_pptx_path) if mois1_pptx_path else {"livraison": None, "sav": None}

    results = {
        "commerce": CommerceCalculator(loader.dfs).compute(year, month),
        "preparation": PreparationCalculator(loader.dfs).compute(year, month),
        "livraison": LivraisonComptaCalculator(loader.dfs).compute(year, month, mois1_reference=mois1["livraison"]),
        "sav": SavCalculator(loader.dfs).compute(year, month, mois1_reference=mois1["sav"]),
    }

    summaries = {
        "commerce": monthly_commerce_excel.summary_rows(results["commerce"]),
        "preparation": monthly_preparation_excel.summary_rows(results["preparation"]),
        "livraison": monthly_livraison_excel.summary_rows(results["livraison"]),
        "sav": monthly_sav_excel.summary_rows(results["sav"]),
    }

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = generate_bilan_pptx(results, str(template_path), str(Path(tmp) / "bilan.pptx"))
        return {"results": results, "summaries": summaries, "pptx": Path(pptx_path).read_bytes()}
