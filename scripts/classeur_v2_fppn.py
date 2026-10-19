# -*- coding: utf-8 -*-
"""
FPPN - Classeur des calculs v2 (rapport + compléments + traçabilité des chiffres)
================================================================================

Regroupe les onglets des deux classeurs de calculs et ajoute l'onglet « Traçabilité » :
chaque chiffre écrit dans le PowerPoint, sa diapositive, sa valeur exacte et le résultat du contrôle automatique.
Utilisation : python classeur_v2_fppn.py <calculs_rapport.xlsx> <calculs_complements.xlsx> <rapport_index.json> <sortie.xlsx> [calculs_ratios.xlsx]
Avec le classeur des ratios (version 3), ses onglets « 10_ » sont ajoutés, ainsi qu'un tableau de bord mis en forme.
Version 5 : arguments facultatifs <calculs_provisions.xlsx> <formules.json> : onglets « 12_ » des provisions et onglet
« 11_formules » (registre complet des formules, « Non communiquée » en rouge).
Version 8 : argument facultatif <calculs_sensibilites.xlsx> : onglets « 13 » à formules vivantes (sensibilités et leviers),
copiés avec leurs formules, commentaires, cellules fusionnées et hauteurs de lignes ; recalcul complet à l'ouverture.
"""
import json
import sys
from copy import copy

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

RAPPORT, COMPLEMENTS, INDEX, SORTIE = sys.argv[1:5]
RATIOS = sys.argv[5] if len(sys.argv) > 5 else None
PROV = sys.argv[6] if len(sys.argv) > 6 else None
FORMULES = sys.argv[7] if len(sys.argv) > 7 else None
SENSIB = sys.argv[8] if len(sys.argv) > 8 else None
ENTETE = PatternFill("solid", fgColor="112A40")
_v = __import__("re").search(r"_v(\d+)\.xlsx$", SORTIE)
VER = _v.group(1) if _v else "2"


def copier(src, dst):
    for row in src.iter_rows():
        for c in row:
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.number_format, d.alignment = c.number_format, copy(c.alignment)
            if c.comment:
                d.comment = copy(c.comment)
    for k, dim in src.column_dimensions.items():
        dst.column_dimensions[k].width = dim.width
    for k, dim in src.row_dimensions.items():
        if dim.height:
            dst.row_dimensions[k].height = dim.height
    for m in src.merged_cells.ranges:
        dst.merge_cells(str(m))
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.freeze_panes = src.freeze_panes


def entete(ws, titres, largeurs):
    for j, (t, w) in enumerate(zip(titres, largeurs), 1):
        c = ws.cell(row=1, column=j, value=t)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), ENTETE
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


wb = openpyxl.Workbook()
lm = wb.active
lm.title = "0_lisez_moi"
idx = json.load(open(INDEX, encoding="utf-8"))
trace = idx["trace"]
ok = sum(t["retrouve"] for t in trace)
lignes = [
    f"Classeur des calculs du rapport actuariel et statistique du FPPN, version {VER} (rapport d'étape au 15/10/2026).",
    "Onglets 1 à 7 : calculs de la version 1 (données, effectifs, sinistralité, bilan, Grenier, annexes, placements).",
    "Onglets 8 : compléments v2 (inflation et valeurs réelles, TCAM, tendances, sensibilités, provisions, tables de mortalité).",
    f"Onglet 9_tracabilite : chaque chiffre écrit dans le PowerPoint v{VER}, avec sa diapositive et sa valeur exacte.",
    f"Contrôle automatique : {ok} chiffres retrouvés sur {len(trace)} enregistrés, {idx['diapositives']} diapositives.",
    "Formats : pourcentages à deux décimales, montants exacts en FCFA sans décimales.",
    "Sources : fichiers reçus du FPPN le 15/10/2026, Banque mondiale (FP.CPI.TOTL.ZG), tables CIMA validées le 18/10/2026.",
]
for i, t in enumerate(lignes, 1):
    lm.cell(row=i, column=1, value=t)
lm["A1"].font = Font(bold=True, size=12)
lm.column_dimensions["A"].width = 120

for fichier in [f for f in (RAPPORT, COMPLEMENTS, RATIOS, PROV) if f]:
    src = openpyxl.load_workbook(fichier)
    for ws in src.worksheets:
        copier(ws, wb.create_sheet(ws.title))

tr = wb.create_sheet("9_tracabilite")
entete(tr, ["Diapositive", "Indicateur", "Valeur exacte", "Texte affiché", "Retrouvé"], [12, 48, 22, 30, 11])
for i, t in enumerate(trace, 2):
    tr.cell(row=i, column=1, value=t["diapositive"])
    tr.cell(row=i, column=2, value=t["indicateur"])
    tr.cell(row=i, column=3, value=t["valeur_exacte"])
    tr.cell(row=i, column=4, value=t["texte"])
    tr.cell(row=i, column=5, value="Oui" if t["retrouve"] else "Non").alignment = Alignment(horizontal="center")

fg = wb.create_sheet("9_figures_tableaux")
entete(fg, ["Type", "Numéro", "Titre", "Diapositive"], [10, 10, 80, 12])
r = 2
for typ, liste in (("Figure", idx["figures"]), ("Tableau", idx["tableaux"])):
    for n, titre, dia in liste:
        for j, v in enumerate((typ, n, titre, dia), 1):
            fg.cell(row=r, column=j, value=v)
        r += 1

if RATIOS:
    src = openpyxl.load_workbook(RATIOS)["10_ratios"]
    lignes_r = list(src.iter_rows(values_only=True))
    ent = {v: i for i, v in enumerate(lignes_r[0])}
    tb = wb.create_sheet("10_tableau_de_bord", 1)
    tb.sheet_view.showGridLines = False
    tb["B2"] = "Tableau de bord des ratios prudentiels du FPPN (Décision n° 05/2022/COM/UEMOA)"
    tb["B2"].font = Font(bold=True, size=14, color="112A40")
    tb["B3"] = ("Statut : « Conforme » si la valeur respecte la norme du texte ; « Non conforme » sinon ; « En attente » si une donnée manque ; "
                "« À apprécier » si la norme n'est pas chiffrée. Statut global jugé sur 2022 à 2025 ; 2020, 2021 et 2026 indicatifs.")
    tb["B3"].font = Font(italic=True, size=9, color="53565A")
    cols = ["Priorité", "Famille", "Ratio", "Norme", "2022", "2023", "2024", "2025", "2026", "Statut global (2022 à 2025)", "Statut retraité", "Source"]
    larg = [8, 22, 46, 26, 11, 11, 11, 11, 11, 18, 14, 50]
    fin = Side(style="thin", color="BBBCBC")
    for j, (c, w) in enumerate(zip(cols, larg), 2):
        x = tb.cell(row=5, column=j, value=c.replace("Statut global (2022 à 2025)", "Statut 2022-2025"))
        x.font, x.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="060606")
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        tb.column_dimensions[x.column_letter].width = w
    for i, l in enumerate(lignes_r[1:], 6):
        pct = "%" in str(l[ent["Norme"]]) and "mois" not in str(l[ent["Norme"]])
        for j, c in enumerate(cols, 2):
            v = l[ent[c]] if c in ent else None
            x = tb.cell(row=i, column=j, value=v)
            x.border = Border(bottom=fin)
            x.alignment = Alignment(vertical="center", wrap_text=c in ("Ratio", "Norme", "Source"))
            if c in ("2022", "2023", "2024", "2025", "2026") and isinstance(v, (int, float)):
                x.number_format = "0.00%" if pct else "# ##0.00"
    der = 5 + len(lignes_r) - 1
    for colonne in ("K", "L"):
        plage = f"{colonne}6:{colonne}{der}"
        tb.conditional_formatting.add(plage, FormulaRule(formula=[f'LEFT({colonne}6,4)="Conf"'], fill=PatternFill("solid", fgColor="112A40"), font=Font(color="FFFFFF", bold=True)))
        tb.conditional_formatting.add(plage, FormulaRule(formula=[f'LEFT({colonne}6,3)="Non"'], fill=PatternFill("solid", fgColor="C00000"), font=Font(color="FFFFFF", bold=True)))
        tb.conditional_formatting.add(plage, FormulaRule(formula=[f'OR(LEFT({colonne}6,3)="En ",LEFT({colonne}6,1)="À")'], fill=PatternFill("solid", fgColor="F2F2F2"), font=Font(color="53565A")))
    tb.freeze_panes = "E6"
    lm.cell(row=len(lignes) + 1, column=1, value="Onglets 10 : ratios prudentiels, classement des comptes, marge de solvabilité, excédent par risque, contrôle du PCMS, adossement par échéance ; onglet 10_tableau_de_bord mis en forme.")

if FORMULES:
    import sys as _s
    import os as _o
    _s.path.insert(0, _o.path.dirname(_o.path.abspath(__file__)))
    from formules_fppn import COLONNES
    reg = json.load(open(FORMULES, encoding="utf-8"))
    fo = wb.create_sheet("11_formules", 1)
    entete(fo, [t for _, t, _ in COLONNES], [w for _, _, w in COLONNES])
    for i, x in enumerate(reg, 2):
        for j, (k, _, _) in enumerate(COLONNES, 1):
            c = fo.cell(row=i, column=j, value=x.get(k, ""))
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if str(c.value).startswith("Non communiquée"):
                c.font = Font(color="C00000", bold=True)
    fo.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(COLONNES))}{len(reg) + 1}"
    fo.freeze_panes = "D2"
    lm.cell(row=len(lignes) + 2, column=1, value="Onglet 11_formules : registre complet des formules (méthode du FPPN, méthode réglementaire, écart, risques, source). Onglets 12_ : provisions et variantes de sinistralité (version 5).")
if SENSIB:
    src = openpyxl.load_workbook(SENSIB)
    for ws in src.worksheets:
        copier(ws, wb.create_sheet(ws.title))
    wb.calculation.fullCalcOnLoad = True
    sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
    import charte_excel_fppn
    charte_excel_fppn.appliquer(wb)
    lm.cell(row=len(lignes) + 4, column=1, value="Charte graphique Adenka : Helvetica, titre au-dessus et source sous chaque tableau, en-tête noir, première colonne grise, lignes alternées, cellules modifiables en orange, statuts conformes en bleu nuit et non conformes en rouge.")
    lm.cell(row=len(lignes) + 3, column=1, value="Onglets 13 (version 8) : analyses de sensibilité et leviers de conformité, à formules vivantes. Méthode et références (13a), "
            "paramètres modifiables en orange (13b), données (13c), statuts et montants à combler (13d), facteurs récurrents (13e), montants d'équilibre des leviers (13f), "
            "inflation médicale ANStat (13g), choc de 100 000 000 FCFA et faisabilité conjointe (13h), rendement et taux du marché (13i). Un commentaire explique chaque cellule clé.")
wb.save(SORTIE)
print(f"{len(wb.sheetnames)} onglets ; {ok}/{len(trace)} chiffres retrouvés")
