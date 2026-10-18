# -*- coding: utf-8 -*-
"""
FPPN - Classeur des calculs v2 (rapport + compléments + traçabilité des chiffres)
================================================================================

Regroupe les onglets des deux classeurs de calculs et ajoute l'onglet « Traçabilité » :
chaque chiffre écrit dans le PowerPoint v2, sa diapositive, sa valeur exacte et le résultat du contrôle automatique.
Utilisation : python classeur_v2_fppn.py <calculs_rapport.xlsx> <calculs_complements.xlsx> <rapport_v2_index.json> <sortie.xlsx>
"""
import json
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

RAPPORT, COMPLEMENTS, INDEX, SORTIE = sys.argv[1:5]
ENTETE = PatternFill("solid", fgColor="112A40")


def copier(src, dst):
    for row in src.iter_rows():
        for c in row:
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.number_format, d.alignment = c.number_format, copy(c.alignment)
    for k, dim in src.column_dimensions.items():
        dst.column_dimensions[k].width = dim.width
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
    "Classeur des calculs du rapport actuariel et statistique du FPPN, version 2 (rapport d'étape au 15/10/2026).",
    "Onglets 1 à 7 : calculs de la version 1 (données, effectifs, sinistralité, bilan, Grenier, annexes, placements).",
    "Onglets 8 : compléments v2 (inflation et valeurs réelles, TCAM, tendances, sensibilités, provisions, tables de mortalité).",
    "Onglet 9_tracabilite : chaque chiffre écrit dans le PowerPoint v2, avec sa diapositive et sa valeur exacte.",
    f"Contrôle automatique : {ok} chiffres retrouvés sur {len(trace)} enregistrés, {idx['diapositives']} diapositives.",
    "Formats : pourcentages à deux décimales, montants exacts en FCFA sans décimales.",
    "Sources : fichiers reçus du FPPN le 15/10/2026, Banque mondiale (FP.CPI.TOTL.ZG), tables CIMA validées le 18/10/2026.",
]
for i, t in enumerate(lignes, 1):
    lm.cell(row=i, column=1, value=t)
lm["A1"].font = Font(bold=True, size=12)
lm.column_dimensions["A"].width = 120

for fichier in (RAPPORT, COMPLEMENTS):
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

wb.save(SORTIE)
print(f"{len(wb.sheetnames)} onglets ; {ok}/{len(trace)} chiffres retrouvés")
