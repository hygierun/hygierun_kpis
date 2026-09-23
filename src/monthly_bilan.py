"""Orchestration du Bilan Mensuel consolidé : un seul fichier Input, un zip de sortie (1 PowerPoint +
4 Excel regroupés : Commerce / Préparation+Livraison / Compta+Achats / SAV).

Réutilise tel quel le chargement (MonthlyDataLoader.load_bilan_total) et les 4 calculators de section
existants - ce module ne fait qu'assembler, aucune règle de calcul n'est dupliquée ou modifiée.

Mois-1 : lu depuis le PowerPoint Bilan Mensuel consolidé du mois précédent (optionnel). Sans lui, les
évolutions Mois-1 qui en dépendent (clients bloqués, factures dues, GPS, nombre d'interventions,
productivité, valorisation du stock, articles à épuisement) restent 'n/a', sans crash.
"""

import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Dict, Optional

from . import monthly_commerce_excel, monthly_livraison_excel, monthly_preparation_excel, monthly_sav_excel
from .monthly_bilan_excel import generate_compta_achats_excel, generate_preparation_livraison_excel, generate_sav_only_excel
from .monthly_bilan_mois1 import read_bilan_mois1_reference
from .monthly_bilan_pptx import generate_bilan_pptx
from .monthly_commerce import CommerceCalculator
from .monthly_livraison import LivraisonComptaCalculator
from .monthly_loader import MONTHS_FR, MonthlyDataLoader
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

    commerce_calc = CommerceCalculator(loader.dfs)
    preparation_calc = PreparationCalculator(loader.dfs)
    livraison_calc = LivraisonComptaCalculator(loader.dfs)
    sav_calc = SavCalculator(loader.dfs)

    results = {
        "commerce": commerce_calc.compute(year, month),
        "preparation": preparation_calc.compute(year, month),
        "livraison": livraison_calc.compute(year, month, mois1_reference=mois1["livraison"]),
        "sav": sav_calc.compute(year, month, mois1_reference=mois1["sav"]),
    }

    summaries = {
        "commerce": monthly_commerce_excel.summary_rows(results["commerce"]),
        "preparation": monthly_preparation_excel.summary_rows(results["preparation"]),
        "livraison": monthly_livraison_excel.summary_rows(results["livraison"]),
        "sav": monthly_sav_excel.summary_rows(results["sav"]),
    }

    suffix = f"{MONTHS_FR[month - 1]}{year}"
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pptx_path = generate_bilan_pptx(results, str(template_path), str(tmp / "bilan.pptx"))
        commerce_xlsx = monthly_commerce_excel.generate_commerce_excel(commerce_calc, results["commerce"], str(tmp / "commerce.xlsx"))
        prep_livraison_xlsx = generate_preparation_livraison_excel(
            preparation_calc, results["preparation"], livraison_calc, results["livraison"], str(tmp / "prep_livraison.xlsx"))
        compta_achats_xlsx = generate_compta_achats_excel(
            livraison_calc, results["livraison"], sav_calc, results["sav"], str(tmp / "compta_achats.xlsx"))
        sav_xlsx = generate_sav_only_excel(sav_calc, results["sav"], str(tmp / "sav.xlsx"))

        excels = {
            "commerce": Path(commerce_xlsx).read_bytes(),
            "preparation_livraison": Path(prep_livraison_xlsx).read_bytes(),
            "compta_achats": Path(compta_achats_xlsx).read_bytes(),
            "sav": Path(sav_xlsx).read_bytes(),
        }
        pptx_bytes = Path(pptx_path).read_bytes()

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"Bilan_Mensuel_{suffix}.pptx", pptx_bytes)
            zf.writestr(f"Commerce_{suffix}.xlsx", excels["commerce"])
            zf.writestr(f"Preparation_Livraison_{suffix}.xlsx", excels["preparation_livraison"])
            zf.writestr(f"Compta_Achats_{suffix}.xlsx", excels["compta_achats"])
            zf.writestr(f"SAV_{suffix}.xlsx", excels["sav"])

        return {"results": results, "summaries": summaries, "pptx": pptx_bytes, "excels": excels,
                "zip": zip_buffer.getvalue()}
