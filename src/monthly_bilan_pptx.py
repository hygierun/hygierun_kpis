"""Remplit le PowerPoint Bilan Mensuel consolidé (11 diapos) en réutilisant, telles quelles, les
fonctions de remplissage de chaque section déjà validées. Aucune règle de calcul ni logique de
remplissage n'est dupliquée ou modifiée ici : ce module ne fait qu'aiguiller vers le bon index de
diapo dans le deck fusionné.

Ordre des diapos (copié-collé par Antoine depuis les 4 templates section par section) :
0-1 : garde (titre + transition), aucun champ à remplir
2-4 : Commerce (3 diapos)
5   : Préparation (1 diapo)
6-8 : Livraison & Compta (3 diapos)
9-10: SAV / Achat (2 diapos)
"""

from pathlib import Path

from pptx import Presentation

from .monthly_commerce_pptx import _fill_slide1 as _fill_commerce_1
from .monthly_commerce_pptx import _fill_slide2 as _fill_commerce_2
from .monthly_commerce_pptx import _fill_slide3 as _fill_commerce_3
from .monthly_livraison_pptx import _fill_slide1 as _fill_livraison_1
from .monthly_livraison_pptx import _fill_slide2 as _fill_livraison_2
from .monthly_livraison_pptx import _fill_slide3 as _fill_livraison_3
from .monthly_preparation_pptx import _fill_slide1 as _fill_preparation_1
from .monthly_sav_pptx import (_fill_articles_epuisement, _fill_deplacement, _fill_devis, _fill_docs_nuls,
                               _fill_main_oeuvre, _fill_nb_interventions, _fill_productivite,
                               _fill_valorisation_stock)
from .pptx_helpers import find_shape, set_subtitle

SLIDES_A_REMPLIR = range(2, 11)
# Diapo 10 (Achat/Appro) : le sous-titre n'a qu'un seul run (pas le schéma standard à 3 runs de
# set_subtitle), déjà le cas dans le template SAV/Achat autonome - on la laisse telle quelle.
SLIDES_SOUS_TITRE = [i for i in SLIDES_A_REMPLIR if i != 10]


def generate_bilan_pptx(results: dict, template_path: str, output_path: str) -> str:
    """results : {"commerce": ..., "preparation": ..., "livraison": ..., "sav": ...} - chaque valeur est
    exactement le dict renvoyé par le .compute() de la section correspondante, sans transformation."""
    prs = Presentation(template_path)
    slides = prs.slides

    year, month = results["commerce"]["year"], results["commerce"]["month"]
    for idx in SLIDES_SOUS_TITRE:
        set_subtitle(find_shape(slides[idx], 5), year, month)

    _fill_commerce_1(slides[2], results["commerce"])
    _fill_commerce_2(slides[3], results["commerce"])
    _fill_commerce_3(slides[4], results["commerce"])

    _fill_preparation_1(slides[5], results["preparation"])

    _fill_livraison_1(slides[6], results["livraison"])
    _fill_livraison_2(slides[7], results["livraison"])
    _fill_livraison_3(slides[8], results["livraison"])

    _fill_devis(slides[9], results["sav"])
    _fill_docs_nuls(slides[9], results["sav"])
    _fill_main_oeuvre(slides[9], results["sav"])
    _fill_deplacement(slides[9], results["sav"])
    _fill_productivite(slides[9], results["sav"])
    _fill_nb_interventions(slides[9], results["sav"])

    _fill_valorisation_stock(slides[10], results["sav"])
    _fill_articles_epuisement(slides[10], results["sav"])

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
