# -*- coding: utf-8 -*-
"""
FPPN - Harmonisation du format des PBC après les ajouts d'items (charte du modèle PBC)
- mises en forme conditionnelles de la priorité (E) et du statut (M) étendues à tous les items ;
- zone d'impression étendue à la dernière ligne ;
- hauteur des lignes ajustée au texte (colonnes C et K), sans jamais réduire une hauteur existante ;
- onglets « Points à clarifier » et « Synthèse du pointage » : retour à la ligne, police et hauteurs alignés.
- impression des onglets annexes : paysage, une page en largeur, en-tête répété.
Utilisation : python harmoniser_pbc_fppn.py <pbc1.xlsx> [<pbc2.xlsx> ...]   (fichiers modifiés en place)
"""
import math
import sys
from copy import copy

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList


def lignes(texte, largeur):
    if not texte:
        return 1
    cpl = max(1.0, largeur * 0.95)
    return sum(max(1, math.ceil(len(morceau) / cpl)) for morceau in str(texte).split("\n"))


def largeur(ws, col):
    for dim in ws.column_dimensions.values():
        if dim.min and dim.max and dim.min <= col <= dim.max and dim.width:
            return dim.width
    return 8.43


def hauteur(ws, r, cols, base=15.0, marge=12.0, mini=30.0):
    n = max(lignes(ws.cell(r, c).value, largeur(ws, c)) for c in cols)
    besoin = max(mini, n * base + marge)
    if (ws.row_dimensions[r].height or 15) < besoin:
        ws.row_dimensions[r].height = besoin


def feuille_pbc(ws):
    items = [r for r in range(11, ws.max_row + 1) if str(ws.cell(r, 2).value or "").startswith("=")]
    dern = items[-1]
    # Mises en forme conditionnelles : mêmes règles, plage étendue
    anciennes = [(str(cf.sqref), cf.rules) for cf in ws.conditional_formatting]
    ws.conditional_formatting = ConditionalFormattingList()
    for plage, regles in anciennes:
        col = plage.split(":")[0].rstrip("0123456789")
        for regle in regles:
            ws.conditional_formatting.add(f"{col}11:{col}{dern}", regle)
    # Zone d'impression
    ws.print_area = f"A1:N{dern + 1}"
    # Hauteurs et alignements des items ; style des cellules vides des lignes d'items
    ref = items[0]
    for r in items:
        for c in range(2, 14):
            if not ws.cell(r, c).has_style:
                ws.cell(r, c)._style = copy(ws.cell(ref, c)._style)
        hauteur(ws, r, (3, 11))
    return len(items), dern


def points(ws):
    for r in range(6, ws.max_row + 1):
        if ws.cell(r, 2).value is None:
            continue
        for c in range(2, 6):
            x = ws.cell(r, c)
            x.font = copy(ws.cell(6, c).font)
            x.alignment = copy(ws.cell(6, c).alignment)
            x.border = copy(ws.cell(6, c).border)
        n = ws.cell(r, 2)
        if str(n.value).isdigit():
            n.value = int(n.value)
        n.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="top")
        hauteur(ws, r, (3, 4, 5), mini=31.5)


def mise_en_page(ws, titres="5:5"):
    """Impression : paysage, une page en largeur, en-tête de tableau répété."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = titres
    ws.print_area = f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}{ws.max_row}"


def synthese(ws):
    for r in range(6, 10):
        hauteur(ws, r, (4,), mini=30)


for f in sys.argv[1:]:
    wb = openpyxl.load_workbook(f)
    ws = next(w for w in wb.worksheets if w.title.startswith("PBC"))
    n, d = feuille_pbc(ws)
    if "Points à clarifier" in wb.sheetnames:
        points(wb["Points à clarifier"])
        mise_en_page(wb["Points à clarifier"])
    if "Synthèse du pointage" in wb.sheetnames:
        synthese(wb["Synthèse du pointage"])
        mise_en_page(wb["Synthèse du pointage"])
    wb.save(f)
    print(f"{f} : {n} items harmonisés, dernière ligne {d}")
