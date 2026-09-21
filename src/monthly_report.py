"""Orchestration du Bilan Mensuel : chargement -> calcul -> génération des livrables, section par section."""

import tempfile
from pathlib import Path
from typing import Dict

from . import monthly_commerce_excel, monthly_preparation_excel
from .monthly_commerce import CommerceCalculator
from .monthly_commerce_excel import generate_commerce_excel
from .monthly_commerce_pptx import generate_commerce_pptx
from .monthly_loader import MonthlyDataLoader, commerce_required_columns, preparation_required_columns
from .monthly_preparation import PreparationCalculator
from .monthly_preparation_excel import generate_preparation_excel
from .monthly_preparation_pptx import generate_preparation_pptx

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
COMMERCE_TEMPLATE = TEMPLATES_DIR / "KPI_Commerce.pptx"
PREPARATION_TEMPLATE = TEMPLATES_DIR / "KPI_Preparation.pptx"


def _build_report(loader_method: str, calculator_cls, excel_module, generate_excel, generate_pptx,
                  template_path: Path, input_path: str, year: int, month: int) -> Dict:
    loader = MonthlyDataLoader(input_path)
    if not getattr(loader, loader_method)(year, month):
        raise ValueError("\n".join(loader.errors))

    calc = calculator_cls(loader.dfs)
    result = calc.compute(year, month)

    with tempfile.TemporaryDirectory() as tmp:
        excel_path = generate_excel(calc, result, str(Path(tmp) / "section.xlsx"))
        pptx_path = generate_pptx(result, str(template_path), str(Path(tmp) / "section.pptx"))
        return {
            "result": result,
            "summary": excel_module.summary_rows(result),
            "excel": Path(excel_path).read_bytes(),
            "pptx": Path(pptx_path).read_bytes(),
        }


def build_commerce_report(input_path: str, year: int, month: int, template_path: Path = COMMERCE_TEMPLATE) -> Dict:
    return _build_report("load_commerce", CommerceCalculator, monthly_commerce_excel, generate_commerce_excel,
                         generate_commerce_pptx, template_path, input_path, year, month)


def build_preparation_report(input_path: str, year: int, month: int, template_path: Path = PREPARATION_TEMPLATE) -> Dict:
    return _build_report("load_preparation", PreparationCalculator, monthly_preparation_excel, generate_preparation_excel,
                         generate_preparation_pptx, template_path, input_path, year, month)


MONTHLY_SECTIONS = {
    "commerce": {
        "label": "🛒 Commerce", "name": "Commerce",
        "input_help": "Input_Mensuel_Commerce.xlsx : feuilles Cdes_ALivr, Cdes_Arch, Fact_Arch, Livr_Arch, Stats_NewClients, top10",
        "required": commerce_required_columns, "build": build_commerce_report,
    },
    "preparation": {
        "label": "📦 Préparation", "name": "Préparation",
        "input_help": "Input_Mensuel_Preparation.xlsx : feuilles Livr_Arch, Ach_Recep",
        "required": preparation_required_columns, "build": build_preparation_report,
    },
}
