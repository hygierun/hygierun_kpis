"""Orchestration du Bilan Mensuel : chargement -> calcul -> génération des livrables."""

import tempfile
from pathlib import Path
from typing import Dict

from .monthly_commerce import CommerceCalculator
from .monthly_commerce_excel import generate_commerce_excel
from .monthly_commerce_pptx import generate_commerce_pptx
from .monthly_loader import MonthlyDataLoader

COMMERCE_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "KPI_Commerce.pptx"


def build_commerce_report(input_path: str, year: int, month: int, template_path: Path = COMMERCE_TEMPLATE) -> Dict:
    """Retourne {'result': KPI calculés, 'excel': bytes, 'pptx': bytes} pour la section Commerce."""
    loader = MonthlyDataLoader(input_path)
    if not loader.load_commerce(year, month):
        raise ValueError("\n".join(loader.errors))

    calc = CommerceCalculator(loader.dfs)
    result = calc.compute(year, month)

    with tempfile.TemporaryDirectory() as tmp:
        excel_path = generate_commerce_excel(calc, result, str(Path(tmp) / "commerce.xlsx"))
        pptx_path = generate_commerce_pptx(result, str(template_path), str(Path(tmp) / "commerce.pptx"))
        return {
            "result": result,
            "excel": Path(excel_path).read_bytes(),
            "pptx": Path(pptx_path).read_bytes(),
        }
