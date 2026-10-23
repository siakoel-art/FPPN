# -*- coding: utf-8 -*-
"""
FPPN - Classeur des calculs, compléments de la version 11 du rapport (vérification de cohérence du 23/10/2026)
=============================================================================================================
Ajoute au classeur existant, avec des formules :
  D9_placements_gl         écritures des grands livres 2021 au 31/08/2026 (comptes 5221, 5330, 77), solde cumulé et jours de détention ;
  26_placements_gl         encours moyen pondéré par les jours, produits financiers et rendement, par exercice ;
  27_pidr_plafond          provision du personnel avec le plafond de la convention collective (25 fois le SMIG annuel) ;
  28_ratios_garantie       sinistralité, frais affectés, ratio combiné et résultat technique par garantie ;
  29_qualite_bases         complétude et cohérence des bases (calcul Python sur 4 454 416 lignes : valeurs) ;
et remplace dans 7_rendement_placements la moyenne des soldes de clôture par l'encours moyen des grands livres.
Chaque nouvel onglet reçoit la barre de navigation et le bloc d'explications des autres onglets.
Utilisation : python classeur_v11_fppn.py <classeur.xlsx> <placements_gl.json> <qualite.json> <sortie.xlsx>
"""
import datetime as dt
import json
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

SRC, PGL, QJ, SORTIE = sys.argv[1:5]
IDX = json.load(open(sys.argv[5], encoding="utf-8"))  # index du rapport du 23/10/2026 (traçabilité, figures, tableaux)
G = json.load(open(PGL, encoding="utf-8"))
QB = json.load(open(QJ, encoding="utf-8"))
wb = openpyxl.load_workbook(SRC)
MODELE = wb["7_rendement_placements"]
NOIR, BLEU, ORANGE, GRIS, PY = "000000", "112A40", "EFAC62", "F2F2F2", "DCE6F1"
F9 = Font(name="Helvetica", size=9)
FMT = "# ##0"


def barre(ws, sources):
    """Barre de navigation (ligne 1) copiée de l'onglet modèle."""
    for j in range(1, 5):
        c, m = ws.cell(1, j), MODELE.cell(1, j)
        c.value, c._style = m.value, copy(m._style)
        if m.hyperlink is not None:
            c.hyperlink = copy(m.hyperlink)
    ws.cell(1, 5).value, ws.cell(1, 5)._style = "Chiffres tirés de :", copy(MODELE.cell(1, 5)._style)
    for k, src in enumerate(sources, 6):
        c = ws.cell(1, k, src)
        c._style = copy(MODELE.cell(1, 6)._style)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{src}'!A1")
    ws.freeze_panes = "A3"


def bloc_explications(ws, titre, rubriques):
    ws.cell(2, 1, titre)._style = copy(MODELE.cell(2, 1)._style)
    r = 3
    for lab, txt in rubriques:
        a, b_ = ws.cell(r, 1, lab), ws.cell(r, 2, txt)
        a._style, b_._style = copy(MODELE.cell(3, 1)._style), copy(MODELE.cell(3, 2)._style)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        r += 1
    return r + 1


def entete(ws, r, cols, largeurs):
    for j, (h, w) in enumerate(zip(cols, largeurs), 1):
        c = ws.cell(r, j, h)
        c._style = copy(MODELE.cell(11, 1)._style)
        ws.column_dimensions[c.column_letter].width = w


def cellule(ws, r, j, v, fmt=None, fond=None):
    c = ws.cell(r, j, v)
    c.font = F9
    if fmt:
        c.number_format = fmt
    if fond:
        c.fill = PatternFill("solid", fgColor=fond)
    return c


def bandeau_python(ws, r, texte):
    c = ws.cell(r, 1, texte)
    c.font = Font(name="Helvetica", size=9, bold=True)
    c.fill = PatternFill("solid", fgColor=PY)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 30


# ------------------------------------------------------------------ D9 : écritures des grands livres
d9 = wb.create_sheet("D9_placements_gl")
barre(d9, ["26_placements_gl"])
r0 = bloc_explications(d9, "Données : écritures des grands livres sur les placements (2021 au 31/08/2026)", [
    ("À quoi sert cet onglet", "Lister chaque écriture des comptes de dépôts à terme (5221), du FCP (5330) et des produits financiers (77), pour mesurer l'encours réellement placé pendant l'année."),
    ("D'où viennent les chiffres", "« GRAND LIVRE 2021.xlsx », « GRAND LIVRE DES COMPTES 2022.xlsx », grand livre des comptes 2023, « GRAND LIVRE FPPN 2024.xlsx », grands livres des comptes 2025 et du 01/01 au 31/08/2026, lus par placements_gl_fppn.py."),
    ("Comment le calcul est fait", "Colonne G : solde des placements après l'écriture (somme des débits moins crédits des comptes 5221 et 5330 depuis le début de l'année). Colonne H : jours pendant lesquels ce solde reste placé, jusqu'à l'écriture suivante ou à la fin de la période. Colonne I : solde × jours."),
    ("Exemple chiffré", "Un dépôt de 40 000 000 FCFA placé 8 jours compte pour 320 000 000 FCFA·jours ; la somme de la colonne I divisée par le nombre de jours de l'année donne l'encours moyen."),
    ("Limites", "2026 : période du 01/01 au 31/08/2026 (243 jours)."),
])
entete(d9, r0, ["Exercice", "Date", "Compte", "Placement ?", "Débit", "Crédit", "Solde des placements après l'écriture", "Jours détenus", "Solde × jours"], [9, 12, 12, 11, 16, 16, 20, 11, 20])
lignes = []
for a in sorted(G, key=int):
    for x in G[a]["detail"]:
        lignes.append((int(a), dt.date.fromisoformat(x["date"]), x["compte"], x["debit"], x["credit"]))
lignes.sort(key=lambda t: (t[0], t[1], 0 if t[2].startswith(("5221", "5330")) else 1))
d0 = r0 + 1
for k, (a, d, cpt, de, cr) in enumerate(lignes):
    r = d0 + k
    cellule(d9, r, 1, a)
    cellule(d9, r, 2, d, "dd/mm/yyyy")
    cellule(d9, r, 3, cpt)
    cellule(d9, r, 4, f'=IF(OR(LEFT(C{r},4)="5221",LEFT(C{r},4)="5330"),"Oui","Non")')
    cellule(d9, r, 5, de, FMT)
    cellule(d9, r, 6, cr, FMT)
    cellule(d9, r, 7, f'=SUMIFS($E${d0}:E{r},$A${d0}:A{r},A{r},$D${d0}:D{r},"Oui")-SUMIFS($F${d0}:F{r},$A${d0}:A{r},A{r},$D${d0}:D{r},"Oui")', FMT)
    cellule(d9, r, 8, None, "0")
    cellule(d9, r, 9, f'=IF(D{r}="Oui",G{r}*H{r},0)', FMT)
DF = d0 + len(lignes) - 1
# colonne auxiliaire « dernière écriture du jour » : le solde retenu est celui après la dernière écriture de la journée
d9.cell(r0, 10, "Dernière écriture du jour ?")._style = copy(MODELE.cell(11, 1)._style)
d9.column_dimensions["J"].width = 14
for r in range(d0, DF + 1):
    # dernière écriture « placement » du jour : seule elle porte les jours de détention
    cellule(d9, r, 10, f'=IF(D{r}="Non","",IF(COUNTIFS(A{r + 1}:$A${DF + 1},A{r},B{r + 1}:$B${DF + 1},B{r},D{r + 1}:$D${DF + 1},"Oui")>0,"Non","Oui"))' if r < DF else f'=IF(D{r}="Non","","Oui")')
    fin = "DATE(2026,8,31)" if d9.cell(r, 1).value == 2026 else f"DATE(A{r},12,31)"
    d9.cell(r, 8).value = (f'=IF(J{r}<>"Oui",0,IF(COUNTIFS($A${d0}:$A${DF},A{r},$D${d0}:$D${DF},"Oui",$B${d0}:$B${DF},">"&B{r})>0,'
                           f'_xlfn.MINIFS($B${d0}:$B${DF},$A${d0}:$A${DF},A{r},$D${d0}:$D${DF},"Oui",$B${d0}:$B${DF},">"&B{r})-B{r},{fin}+1-B{r}))')
d9.cell(DF + 2, 1, "Source : grands livres des comptes du FPPN de 2021 au 31/08/2026 ; script placements_gl_fppn.py").font = Font(name="Helvetica", size=8, italic=True)

# ------------------------------------------------------------------ 26 : encours moyen et rendement
p26 = wb.create_sheet("26_placements_gl")
barre(p26, ["D9_placements_gl", "7_rendement_placements"])
r1 = bloc_explications(p26, "Tableau : encours moyen des placements et rendement, lus dans les grands livres", [
    ("À quoi sert cet onglet", "Mesurer l'argent réellement placé pendant l'année, et non le seul solde de clôture, puis le rendement des placements."),
    ("D'où viennent les chiffres", "Onglet D9_placements_gl (écritures des grands livres 2021 au 31/08/2026)."),
    ("Comment le calcul est fait", "Encours moyen = somme (solde × jours détenus) / jours de la période. Rendement = produits financiers (comptes 77) / encours moyen."),
    ("Exemple chiffré", "2025 : 117 912 488 FCFA de produits financiers / 763 752 135 FCFA d'encours moyen = 15,44 %. De 2021 à 2025 : 173 881 316 FCFA / 3 545 662 961 FCFA = 4,90 % par an."),
    ("Comment lire le résultat", "Le DAT EFC est renouvelé chaque mois et soldé en décembre : son solde de clôture est nul, mais il est placé onze mois sur douze."),
    ("Valable pour", "2021 au 31/08/2026 (2026 : 243 jours)."),
    ("Limites", "Les intérêts du DAT PEP, constatés chaque année puis annulés, sont encaissés en une fois à sa clôture le 30/10/2025 : le rendement annuel est irrégulier, le rendement cumulé de 2021 à 2025 ne l'est pas."),
])
entete(p26, r1, ["Exercice", "Jours de la période", "Somme solde × jours", "Encours moyen placé", "Solde à la clôture", "Produits financiers (77)", "Rendement"], [10, 12, 22, 20, 18, 20, 12])
for k, a in enumerate(range(2021, 2027)):
    r = r1 + 1 + k
    cellule(p26, r, 1, a)
    cellule(p26, r, 2, "=DATE(2026,8,31)-DATE(2026,1,1)+1" if a == 2026 else f"=DATE({a},12,31)-DATE({a},1,1)+1")
    cellule(p26, r, 3, f"=SUMIFS('D9_placements_gl'!$I${d0}:$I${DF},'D9_placements_gl'!$A${d0}:$A${DF},A{r})", FMT)
    cellule(p26, r, 4, f"=C{r}/B{r}", FMT)
    cellule(p26, r, 5, f"=SUMIFS('D9_placements_gl'!$E${d0}:$E${DF},'D9_placements_gl'!$A${d0}:$A${DF},A{r},'D9_placements_gl'!$D${d0}:$D${DF},\"Oui\")-SUMIFS('D9_placements_gl'!$F${d0}:$F${DF},'D9_placements_gl'!$A${d0}:$A${DF},A{r},'D9_placements_gl'!$D${d0}:$D${DF},\"Oui\")", FMT)
    cellule(p26, r, 6, f"=SUMIFS('D9_placements_gl'!$F${d0}:$F${DF},'D9_placements_gl'!$A${d0}:$A${DF},A{r},'D9_placements_gl'!$C${d0}:$C${DF},\"77*\")-SUMIFS('D9_placements_gl'!$E${d0}:$E${DF},'D9_placements_gl'!$A${d0}:$A${DF},A{r},'D9_placements_gl'!$C${d0}:$C${DF},\"77*\")", FMT)
    cellule(p26, r, 7, f"=F{r}/D{r}", "0.00%")
rt = r1 + 8
cellule(p26, rt, 1, "2021 à 2025").font = Font(name="Helvetica", size=9, bold=True)
cellule(p26, rt, 3, f"=SUM(C{r1 + 1}:C{r1 + 5})", FMT)
cellule(p26, rt, 4, f"=SUM(D{r1 + 1}:D{r1 + 5})", FMT)
cellule(p26, rt, 6, f"=SUM(F{r1 + 1}:F{r1 + 5})", FMT)
cellule(p26, rt, 7, f"=F{rt}/D{rt}", "0.00%")
p26.cell(rt + 1, 1, "Rendement moyen 2021 à 2025 = somme des produits / somme des encours moyens annuels.").font = Font(name="Helvetica", size=8, italic=True)
P26 = {a: r1 + 1 + k for k, a in enumerate(range(2021, 2027))}

# ------------------------------------------------------------------ 7 : rendement sur l'encours des grands livres
m = MODELE
for r in range(12, 17):
    a = int(m.cell(r, 1).value)
    m.cell(r, 3).value = f"='26_placements_gl'!$D${P26[a]}"
for r in range(3, 10):
    lab = m.cell(r, 1).value
    if lab == "D'où viennent les chiffres":
        m.cell(r, 2).value = "Produits financiers : onglet 10_agregats (comptes 77). Placements moyens : onglet 26_placements_gl (encours moyen lu dans les grands livres)."
    elif lab == "Comment le calcul est fait":
        m.cell(r, 2).value = "Rendement = produits financiers (classe 77) / encours moyen des placements (DAT et FCP), pondéré par les jours de détention dans l'année."
    elif lab == "Exemple chiffré":
        m.cell(r, 2).value = "En 2025 : 117 912 488 FCFA de produits financiers / 763 752 135 FCFA d'encours moyen = 15,44 %."
    elif lab == "Limites":
        m.cell(r, 2).value = "Rendement irrégulier : les intérêts du DAT PEP (132 663 262 FCFA) sont encaissés en une fois à sa clôture le 30/10/2025."
m.cell(18, 1).value = "Source : onglet 10_agregats (comptes 77) ; onglet 26_placements_gl (grands livres 2021 à 2025)"
m.cell(1, 7).value, m.cell(1, 7)._style = "26_placements_gl", copy(m.cell(1, 6)._style)
m.cell(1, 7).hyperlink = Hyperlink(ref="G1", location="'26_placements_gl'!A1")

# ------------------------------------------------------------------ 27 : provision du personnel plafonnée
p27 = wb.create_sheet("27_pidr_plafond")
barre(p27, ["14b_pidr_individuel", "14c_pidr_synthese"])
r2 = bloc_explications(p27, "Tableau : provision du personnel avec le plafond de la convention collective", [
    ("À quoi sert cet onglet", "Mesurer l'effet du plafond de l'indemnité de départ à la retraite prévu par la convention collective."),
    ("D'où viennent les chiffres", "Onglet 14b_pidr_individuel (calcul agent par agent). SMIG : décret n° 2022-986 du 21/12/2022 (75 000 FCFA par mois depuis le 01/01/2023)."),
    ("Comment le calcul est fait", "Plafond = 25 × SMIG annuel (2 080 heures, soit 12 mois). Provision plafonnée d'un agent = provision × MIN(1 ; plafond / indemnité projetée)."),
    ("Exemple chiffré", "Plafond = 25 × 75 000 × 12 = 22 500 000 FCFA. Un agent dont l'indemnité projetée vaut 45 000 000 FCFA voit sa provision divisée par deux."),
    ("Valable pour", "Agents de la DISA 2025, évaluation au 31/12/2025."),
    ("Limites", "Le décret n° 2017-210 (art. 4 et 5) ne prévoit pas de plafond : le calcul central du rapport ne l'applique pas. Le SMIG est gardé constant jusqu'au départ."),
])
cellule(p27, r2, 1, "SMIG mensuel (FCFA)")
cellule(p27, r2, 2, 75000, FMT, ORANGE)
cellule(p27, r2 + 1, 1, "Plafond (25 × SMIG annuel)")
cellule(p27, r2 + 1, 2, f"=25*B{r2}*12", FMT)
h27 = r2 + 3
entete(p27, h27, ["Identifiant", "Indemnité projetée", "Provision sans plafond", "Coefficient de plafond", "Provision avec plafond", "Plafond dépassé ?"], [12, 18, 18, 14, 18, 14])
for k in range(79):
    r, src = h27 + 1 + k, 6 + k
    cellule(p27, r, 1, f"='14b_pidr_individuel'!A{src}")
    cellule(p27, r, 2, f"='14b_pidr_individuel'!O{src}", FMT)
    cellule(p27, r, 3, f"='14b_pidr_individuel'!T{src}", FMT)
    cellule(p27, r, 4, f"=IF(B{r}>0,MIN(1,$B${r2 + 1}/B{r}),1)", "0.0000")
    cellule(p27, r, 5, f"=C{r}*D{r}", FMT)
    cellule(p27, r, 6, f'=IF(B{r}>$B${r2 + 1},"Oui","Non")')
rT = h27 + 80
cellule(p27, rT, 1, "Total").font = Font(name="Helvetica", size=9, bold=True)
cellule(p27, rT, 3, f"=SUM(C{h27 + 1}:C{rT - 1})", FMT)
cellule(p27, rT, 5, f"=SUM(E{h27 + 1}:E{rT - 1})", FMT)
cellule(p27, rT, 6, f'=COUNTIF(F{h27 + 1}:F{rT - 1},"Oui")')

# ------------------------------------------------------------------ 28 : ratios par garantie
p28 = wb.create_sheet("28_ratios_garantie")
barre(p28, ["10_excedent_par_risque"])
r3 = bloc_explications(p28, "Tableau : ratios par garantie et résultat technique", [
    ("À quoi sert cet onglet", "Donner, pour la maladie et le décès, la sinistralité, la part des frais, le ratio combiné et le résultat technique."),
    ("D'où viennent les chiffres", "Onglet 10_excedent_par_risque (cotisations, prestations, frais affectés au prorata des cotisations)."),
    ("Comment le calcul est fait", "Sinistralité = prestations / cotisations. Taux de frais = frais affectés / cotisations. Ratio combiné = (prestations + frais) / cotisations. Résultat technique = cotisations - prestations - frais."),
    ("Exemple chiffré", "Maladie 2025 : (1 771 635 361 + 875 420 236) / 2 409 281 030 = 109,87 %, soit un résultat technique de -237 774 567 FCFA."),
    ("Comment lire le résultat", "Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte, frais compris."),
    ("Valable pour", "Exercices 2020 à 2025 et du 01/01 au 31/08/2026."),
    ("Limites", "Frais répartis au prorata des cotisations, faute de clé fournie par le FPPN. Hors produits financiers et éléments exceptionnels."),
])
entete(p28, r3, ["Exercice", "Garantie", "Cotisations", "Prestations", "Frais affectés", "Sinistralité", "Taux de frais", "Ratio combiné", "Résultat technique"], [10, 12, 18, 18, 18, 12, 12, 12, 18])
r = r3 + 1
for k in range(7):
    src = 11 + k
    for gar, cc, pp, ff in (("Maladie", "B", "C", "D"), ("Décès", "F", "G", "H")):
        cellule(p28, r, 1, f"='10_excedent_par_risque'!A{src}")
        cellule(p28, r, 2, gar)
        cellule(p28, r, 3, f"='10_excedent_par_risque'!{cc}{src}", FMT)
        cellule(p28, r, 4, f"='10_excedent_par_risque'!{pp}{src}", FMT)
        cellule(p28, r, 5, f"='10_excedent_par_risque'!{ff}{src}", FMT)
        cellule(p28, r, 6, f"=D{r}/C{r}", "0.00%")
        cellule(p28, r, 7, f"=E{r}/C{r}", "0.00%")
        cellule(p28, r, 8, f"=(D{r}+E{r})/C{r}", "0.00%")
        cellule(p28, r, 9, f"=C{r}-D{r}-E{r}", FMT)
        r += 1

# ------------------------------------------------------------------ 29 : qualité des bases (calcul Python)
p29 = wb.create_sheet("29_qualite_bases")
barre(p29, ["D3_precomptes_mensuels"])
r4 = bloc_explications(p29, "Tableau : complétude et cohérence des bases de données", [
    ("À quoi sert cet onglet", "Contrôler, variable par variable, les valeurs manquantes et les incohérences des bases de précomptes et des prêts du Grenier."),
    ("D'où viennent les chiffres", "Base de travail des fichiers de précomptes (lignes retenues de janvier 2020 à août 2026) et balance âgée du Grenier au 31/08/2026."),
    ("Comment le calcul est fait", "Comptage des valeurs manquantes et des lignes en anomalie par le script qualite_bases_fppn.py ; voir le rapport, diapositive « Qualité des bases, variable par variable »."),
    ("Limites", "Les lignes multiples d'un même matricule dans un mois sont normales pour le PEP (plusieurs plans) et le Grenier (plusieurs prêts)."),
])
bandeau_python(p29, r4 - 1, "Calcul Python : les 4 454 416 lignes détaillées ne tiennent pas dans Excel. Script qualite_bases_fppn.py (base_fppn.py pour la base) ; résultats repris en valeurs, voir le rapport.")
entete(p29, r4, ["Base", "Variable ou contrôle", "Lignes", "Manquantes ou anomalies", "Part"], [32, 40, 12, 18, 10])
r = r4 + 1
for x in QB["completude"]:
    for j, v in enumerate((x["base"], "Complétude : " + x["variable"], x["lignes"], x["manquantes"]), 1):
        cellule(p29, r, j, v, FMT if j in (3, 4) else None, PY if j in (3, 4) else None)
    cellule(p29, r, 5, f"=D{r}/C{r}", "0.00%")
    r += 1
for x in QB["coherence"]:
    for j, v in enumerate((x["base"], "Cohérence : " + x["controle"], x["lignes"], x["anomalies"]), 1):
        cellule(p29, r, j, v, FMT if j in (3, 4) else None, PY if j in (3, 4) else None)
    cellule(p29, r, 5, f"=D{r}/C{r}", "0.00%")
    r += 1

# ------------------------------------------------------------------ ordre des onglets, guide
for nom in ("26_placements_gl", "27_pidr_plafond", "28_ratios_garantie", "29_qualite_bases"):
    ws = wb[nom]
    wb.move_sheet(ws, offset=wb.sheetnames.index("13a_methode") - wb.sheetnames.index(nom))
wb.move_sheet(wb["D9_placements_gl"], offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("D9_placements_gl"))
gd = wb["0_guide"]
rg = gd.max_row + 2
gd.cell(rg, 1, "Compléments du 23/10/2026").font = Font(name="Helvetica", size=11, bold=True)
for k, (nom, txt) in enumerate((("26_placements_gl", "Encours moyen des placements et rendement, lus dans les grands livres (remplace la moyenne des soldes de clôture)."),
                                 ("27_pidr_plafond", "Provision du personnel avec le plafond de 25 fois le SMIG annuel (convention collective, art. 40)."),
                                 ("28_ratios_garantie", "Sinistralité, frais, ratio combiné et résultat technique par garantie."),
                                 ("29_qualite_bases", "Complétude et cohérence des bases (calcul Python, valeurs)."),
                                 ("D9_placements_gl", "Écritures des grands livres sur les dépôts à terme, le FCP et les produits financiers.")), rg + 1):
    c = gd.cell(k, 1, nom)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{nom}'!A1")
    c.font = Font(name="Helvetica", size=9, color="0563C1", underline="single")
    gd.cell(k, 2, txt).font = F9
# traçabilité, figures et tableaux : numéros de diapositive du rapport livré avec ce classeur
def _reecrire(ws, lignes):
    r0 = next(r for r in range(1, 30) if ws.cell(r, 1).value in ("Diapositive", "Type"))
    nc = max(len(x) for x in lignes)
    st = [copy(ws.cell(r0 + 1, j)._style) for j in range(1, nc + 1)]
    for r in range(r0 + 1, ws.max_row + 1):
        for j in range(1, nc + 1):
            ws.cell(r, j).value = None
    for i, x in enumerate(lignes, r0 + 1):
        for j, v in enumerate(x, 1):
            ws.cell(i, j, v)._style = copy(st[j - 1])
    for t in ws.tables.values():
        t.ref = f"A{r0}:{get_column_letter(nc)}{r0 + len(lignes)}"


_ok = sum(1 for t in IDX["trace"] if t["retrouve"])
wb["0_lisez_moi"]["A6"].value = f"Contrôle automatique : {_ok} chiffres retrouvés sur {len(IDX['trace'])} enregistrés, {IDX['diapositives']} diapositives."
_lm = wb["0_lisez_moi"]
_lm["A8"].value = "Sources : fichiers reçus du FPPN du 15/10/2026 au 22/10/2026, Banque mondiale (FP.CPI.TOTL.ZG), tables CIMA validées le 18/10/2026, décret n° 2022-986 du 21/12/2022 (SMIG)."
_der = max(r for r in range(1, _lm.max_row + 1) if _lm.cell(r, 1).value)
if not str(_lm.cell(_der, 1).value).startswith("Onglets 26 à 29"):
    _lm.cell(_der + 1, 1, "Onglets 26 à 29 et D9 : encours et rendement des placements lus dans les grands livres (26, D9), provision du personnel avec le plafond de la convention collective (27), ratios par garantie (28), qualité des bases variable par variable (29).")._style = copy(_lm.cell(_der, 1)._style)
_reecrire(wb["9_tracabilite"], [(t["diapositive"], t["indicateur"], t["valeur_exacte"], t["texte"], "Oui" if t["retrouve"] else "Non") for t in IDX["trace"]])
_reecrire(wb["9_figures_tableaux"], [("Figure", *x) for x in IDX["figures"]] + [("Tableau", *x) for x in IDX["tableaux"]])

# valeurs publiées du contrôle : rendement lu dans les grands livres (rapport v11), à la place de l'encours moyen des balances
_rdt = {a: (G[str(a)]["placements"]["moyenne"], G[str(a)]["produits_77"] / G[str(a)]["placements"]["moyenne"]) for a in range(2021, 2026)}
_pub = {}
for k, a in enumerate(range(2021, 2026)):
    _pub[("7_rendement_placements", f"C{11 + k}")] = _rdt[a][0]
    _pub[("7_rendement_placements", f"D{11 + k}")] = _rdt[a][1]
    _pub[("10_ratios", "KMOQS"[k] + "26")] = _rdt[a][1]
for row in wb["0_controle"].iter_rows():
    cle = (row[0].value, row[1].value)
    if cle in _pub:
        row[2].value = _pub.pop(cle)
assert not _pub, _pub
wb.save(SORTIE)
print(SORTIE, "écritures D9 :", len(lignes), "| lignes", d0, "à", DF)
