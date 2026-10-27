"""Ajoute au PDF d'un deck les liens internes des boutons de navigation (formes « Nav … » munies d'une action
« aller à la diapositive »), aux positions exactes des formes. Sert quand l'export PDF ne conserve pas ces actions.

Usage : python3 pdf_liens_navigation.py <deck.pptx> <deck.pdf> <sortie.pdf>
"""
import sys

from pptx import Presentation
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DictionaryObject, FloatObject, NameObject, NumberObject

PPTX, PDF, SORTIE = sys.argv[1:4]
prs = Presentation(PPTX)
lecteur = PdfReader(PDF)
ecrivain = PdfWriter()
ecrivain.append(lecteur, import_outline=False)
assert len(ecrivain.pages) == len(prs.slides), "le PDF doit avoir une page par diapositive"
index = {sl.slide_id: k for k, sl in enumerate(prs.slides)}
L, H = prs.slide_width, prs.slide_height
n = 0
for k, sl in enumerate(prs.slides):
    page = ecrivain.pages[k]
    pw, ph = float(page.mediabox.width), float(page.mediabox.height)
    for sh in sl.shapes:
        cible = sh.click_action.target_slide if sh.click_action.action else None
        if cible is None:
            continue
        x1 = sh.left / L * pw
        x2 = (sh.left + sh.width) / L * pw
        y1 = ph - (sh.top + sh.height) / H * ph
        y2 = ph - sh.top / H * ph
        dest = ArrayObject([ecrivain.pages[index[cible.slide_id]].indirect_reference, NameObject("/Fit")])
        lien = DictionaryObject({NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Link"),
                                 NameObject("/Rect"): ArrayObject([FloatObject(x1), FloatObject(y1), FloatObject(x2), FloatObject(y2)]),
                                 NameObject("/Border"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(0)]),
                                 NameObject("/Dest"): dest})
        page.annotations = page.annotations or ArrayObject()
        page.annotations.append(ecrivain._add_object(lien))
        n += 1
# signets du lecteur PDF : un par diapositive titrée, les activités rangées sous leur groupe
parent = None
for k, sl in enumerate(prs.slides):
    sep = next((sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "Rectangle 9" and sh.text_frame.text.strip()), None)
    titre = next((sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"), None)
    if sep:
        parent = ecrivain.add_outline_item(sep, k)
    elif titre:
        ecrivain.add_outline_item(titre, k, parent=parent)
ecrivain.page_mode = "/UseOutlines"
ecrivain.write(SORTIE)
print(SORTIE, "liens ajoutés :", n)
