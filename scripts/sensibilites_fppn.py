# -*- coding: utf-8 -*-
"""
FPPN - Analyses de sensibilité et leviers de conformité, 2020 au 31/08/2026 (rapport v8)
=====================================================================================

Classeur à formules vivantes (onglets « 13 ») : toutes les valeurs calculées sont des formules Excel qui partent des
données de l'onglet 13c et des paramètres de l'onglet 13b ; un commentaire explique chaque cellule clé.

Méthode (détaillée dans l'onglet 13a, avec ses références) :
1. reprise des numérateurs et dénominateurs des ratios de la Décision n° 05/2022/COM/UEMOA (calcul v3 à v7) ;
2. statut de chaque ratio chaque année, puis qualification « structurel » (non conforme toutes les années mesurées) ;
3. montant à combler : inversion de la formule du ratio (écart à la borne la plus proche) ;
4. facteurs récurrents : regroupement des ratios par comptes en cause ;
5. leviers tirés des textes : montants d'équilibre par ratio et par année, effet d'un choc de 100 000 000 FCFA
   enregistré en partie double, une hypothèse à la fois (IAS 19 § 145, IFRS 17 § 128) ;
6. faisabilité conjointe des normes de flux (démonstration algébrique) ;
7. inflation médicale : sous-indices IHPC de l'ANStat pondérés par les prestations du FPPN.

Utilisation : python sensibilites_fppn.py <ratios.json> <resultats.json> <dossier_csv_ihpc> <sortie.xlsx> [<dossier_csv_investing>]
Le dossier facultatif contient les rendements quotidiens des obligations de l'État de Côte d'Ivoire à 5 et 10 ans
exportés d'Investing.com (« Rendement de l'Obligation Cote d'Ivoire 5-Year / 10-Year - Données Historiques »).
Le classeur produit est recalculé par LibreOffice (contrôle) ; les valeurs relues alimentent le rapport.
"""
import glob
import json
import os
import sys

import openpyxl
import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

RATIOS, RESULTATS, IHPC, SORTIE = sys.argv[1:5]
INVEST = sys.argv[5] if len(sys.argv) > 5 else None
RJ = json.load(open(RATIOS, encoding="utf-8"))
RS = json.load(open(RESULTATS, encoding="utf-8"))
EX = list(range(2020, 2027))
C0 = 3  # colonne de 2020
COL = {ex: L(C0 + i) for i, ex in enumerate(EX)}
NOIR, BLEU, GRIS = "060606", "112A40", "F2F2F2"
H = PatternFill("solid", fgColor=BLEU)
HS = PatternFill("solid", fgColor=GRIS)
FIN = Border(bottom=Side(style="thin", color="BBBCBC"))
AUT = "Adenka"
FMT_M, FMT_P, FMT_R = "#,##0", "0.00%", "0.00"

wb = openpyxl.Workbook()
wb.remove(wb.active)


def com(cell, texte, larg=420, haut=160):
    c = Comment(texte, AUT)
    c.width, c.height = larg, haut
    cell.comment = c


def titre(ws, t, sous=None):
    ws.sheet_view.showGridLines = False
    ws["A1"] = t
    ws["A1"].font = Font(bold=True, size=13, color=BLEU)
    if sous:
        ws["A2"] = sous
        ws["A2"].font = Font(italic=True, size=9, color="53565A")
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells("A2:L2")
        ws.row_dimensions[2].height = 42


def entete(ws, r, valeurs, fill=H):
    for j, v in enumerate(valeurs, 1):
        c = ws.cell(r, j, v)
        c.font = Font(bold=True, color="FFFFFF" if fill is H else NOIR)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def largeurs(ws, ls):
    for j, w in enumerate(ls, 1):
        ws.column_dimensions[L(j)].width = w


def cel(ws, r, c, v, fmt=None, gras=False, wrap=False, couleur=None):
    x = ws.cell(r, c, v)
    if fmt:
        x.number_format = fmt
    if gras or couleur:
        x.font = Font(bold=gras, color=couleur)
    x.alignment = Alignment(wrap_text=wrap, vertical="top")
    x.border = FIN
    return x


ANS = ["Code", "Libellé"] + [str(e) for e in EX]

# =============================================================================================== 13a méthode
REFS = [
    ("R1", "Décision n° 05/2022/COM/UEMOA (annexe, tableaux 1 et 2, § 1.2 et section II)", "Formules et normes des ratios prudentiels et de gestion.", "Statut de chaque ratio ; montant à combler obtenu en inversant la formule du texte."),
    ("R2", "Règlement n° 07/2009/CM/UEMOA du 26/06/2009, art. 83 et 84", "Art. 83 : sanctions pour non-respect des règles prudentielles durant un exercice. Art. 84 : retrait d'agrément pour non-respect des règles prudentielles sur deux exercices.", "Qualification des écarts : un écart répété sur deux exercices ou plus est « récurrent » ; présent toutes les années mesurées, il est « structurel »."),
    ("R3", "Arrêté n° 00115/MEPS/CAB du 12/12/2023 (Côte d'Ivoire), grille des sanctions", "Ligne 13 : non-respect des règles prudentielles sur deux exercices.", "Même qualification, au niveau national."),
    ("R4", "Règlement d'exécution n° 004/2022/COM/UEMOA, art. 16, 17, 19, 21 et 36", "Art. 16 : réassurance sur décision de l'AG. Art. 17 : placements (sécurité, rendement, liquidité, diversification). Art. 19 al. 3 : réserve de trésorerie d'un trimestre. Art. 21 : fonds propres et prestations des risques longs. Art. 36 : plan de redressement.", "Leviers L5 (placements) et L7 (fonds propres) ; ratios RL4 et RL5."),
    ("R5", "Statuts du FPPN du 15/06/2023, art. 17", "Cotisation obligatoire de 5 % de la solde de base : 3,5 % maladie-maternité et 1,5 % invalidité-décès.", "Base du levier L1 (taux de cotisation)."),
    ("R6", "Statuts, art. 93 ; règlement intérieur, art. 86", "En cas de déséquilibre financier : modification du taux ou du montant des cotisations, ou du régime des prestations, et toute mesure de redressement.", "Fondement des leviers L1 (cotisations) et L2 (prestations)."),
    ("R7", "Statuts, art. 90 ; règlement intérieur, art. 83", "Réserve légale d'au moins 20 % à 50 % des résultats excédentaires nets de l'exercice précédent.", "Levier L4 (dotation à la réserve légale), bornes de 20 % et 50 %."),
    ("R8", "Statuts, art. 79, 81 et 85 ; règlement intérieur, art. 78", "Ressources (dont dons et legs) ; placements autorisés par le Conseil d'Administration, aux conditions fixées par le Ministre.", "Leviers L5 (réallocation des placements) et L7 (apport en fonds propres)."),
    ("R9", "Règlement intérieur, art. 41 et 45", "Couverture maladie limitée aux soins liés à une hospitalisation ; exclusions (optique, prothèses dentaires, maladies de longue durée...).", "Facteur F5 et levier L2 (élargissement de la couverture)."),
    ("R10", "Association Actuarielle Internationale, ISAP 1 « General Actuarial Practice », adoptée le 1er décembre 2018, § 2.7.7", "L'actuaire examine la sensibilité de la méthode aux variations des hypothèses clés, lorsque c'est approprié.", "Obligation de l'analyse de sensibilité dans la démarche actuarielle."),
    ("R11", "IAS 19 révisée (2011), § 145", "Analyse de sensibilité pour chaque hypothèse significative, aux changements raisonnablement possibles, avec description de la méthode, des hypothèses et de leurs limites.", "Méthode « une hypothèse à la fois », toutes choses égales par ailleurs, et présentation des limites."),
    ("R12", "IFRS 17, § 128 a) et § 129 ; IFRS 7, § 40", "Sensibilité du résultat et des capitaux propres aux variables de risque raisonnablement possibles ; méthode et hypothèses décrites.", "Effet du choc sur le résultat et les fonds propres (écritures en partie double)."),
    ("R13", "ANStat, IHPC par poste de consommation (national), fonction Santé, extraction du 19/10/2026", "Indices mensuels de janvier 2017 à octobre 2024 pour huit postes de la santé.", "Inflation médicale (onglet 13g)."),
    ("R15", "Investing.com, « Rendement de l'Obligation Cote d'Ivoire 5-Year / 10-Year - Données Historiques », exports du 19/10/2026", "Cotations quotidiennes des obligations de l'État de Côte d'Ivoire à 5 ans (mars 2021 au 15/10/2026) et à 10 ans (mars 2021 au 19/10/2026). Source secondaire (agrégateur du marché secondaire).", "Lecture complémentaire du ratio de rentabilité des placements (onglet 13i)."),
    ("R14", "BCEAO, Rapport sur la Politique Monétaire dans l'UMOA, mars 2026, tableau 12", "Inflation de la santé dans l'UEMOA : -1,3 % (T2 2025), -2,4 % (T3 2025), -1,9 % (T4 2025).", "Repli pour 2025, faute de série ANStat après octobre 2024."),
]
ws = wb.create_sheet("13a_methode")
titre(ws, "Analyses de sensibilité et leviers de conformité : méthode et références",
      "Les normes R10 à R12 sont des références de bonne pratique actuarielle et comptable, retenues pour la méthode ; elles ne sont pas opposables au FPPN. "
      "Les textes R1 à R9 sont opposables. Tous les calculs des onglets 13b à 13h sont des formules ; modifier un paramètre de l'onglet 13b met à jour l'ensemble.")
ETAPES = [
    ("1. Données", "Numérateurs et dénominateurs des ratios de la Décision n° 05/2022, recalculés depuis les balances (onglets 10_ et 13c).", "R1"),
    ("2. Statut", "Conforme si la valeur respecte la borne du texte. Une valeur négative d'un ratio plafonné (excédent, équilibre stable) est non conforme.", "R1"),
    ("3. Qualification", "Structurel : non conforme toutes les années mesurées de 2020 au 31/08/2026. Récurrent : au moins deux exercices de jugement (2022 à 2025).", "R2, R3"),
    ("4. Montant à combler", "Inversion de la formule : montant qui ramène le ratio sur la borne la plus proche, les autres termes restant inchangés.", "R1, R11"),
    ("5. Facteurs récurrents", "Regroupement des ratios non conformes selon les comptes qui expliquent l'écart chaque année.", "R1"),
    ("6. Leviers", "Leviers prévus par les textes. Pour chacun : montant d'équilibre par ratio et par année, et effet d'un choc de 100 000 000 FCFA enregistré en partie double, une hypothèse à la fois.", "R4 à R9, R11, R12"),
    ("7. Faisabilité conjointe", "Démonstration algébrique : les normes de flux (sinistralité, coût par unité prestée, excédent) peuvent-elles être respectées ensemble ?", "R1"),
    ("8. Inflation médicale", "Sous-indices IHPC pondérés par les prestations maladie du FPPN de 2025 ; seuil d'inflation qui ferait dépasser 80 % de sinistralité.", "R13, R14"),
    ("Limites", "Choc de 100 000 000 FCFA : convention de présentation pour comparer les leviers, pas une prévision. Effets croisés et réactions des adhérents non modélisés. 2026 : huit mois, balance provisoire, indicatif. 2020 et 2021 : antérieurs à la Décision du 1er juin 2022, indicatifs.", "R11, R12"),
]
entete(ws, 4, ["Étape", "Contenu", "Références"])
for i, (a, b, c) in enumerate(ETAPES, 5):
    cel(ws, i, 1, a, gras=True)
    cel(ws, i, 2, b, wrap=True)
    cel(ws, i, 3, c)
    ws.row_dimensions[i].height = 44
r0 = 5 + len(ETAPES) + 1
entete(ws, r0, ["Réf.", "Texte ou norme", "Disposition", "Usage dans l'analyse"])
for i, ref in enumerate(REFS, r0 + 1):
    for j, v in enumerate(ref, 1):
        cel(ws, i, j, v, wrap=True, gras=j == 1)
    ws.row_dimensions[i].height = 58
largeurs(ws, [22, 70, 70, 60])

# =============================================================================================== 13c données
R = {x["code"]: x for x in RJ["ratios"]}
AG = RJ["agregats"]
MS = {m["exercice"]: m for m in RJ["marge_solvabilite"]}


def serie(code, cle):
    d = R[code][cle] or {}
    return {ex: d.get(str(ex), d.get(ex)) for ex in EX}


VARS = [  # (clé, libellé, série, source)
    ("COT", "Cotisations acquises (70 + 7906 - 6906)", serie("SIN", "den"), "Dénominateur de SIN ; onglet 10_ratios"),
    ("PRE", "Charges de prestations (60 + 6901 - 7901)", serie("SIN", "num"), "Numérateur de SIN"),
    ("CF", "Charges de fonctionnement (612, 613, 618, 62 à 68)", serie("CHF", "num"), "Numérateur de CHF et de CUP"),
    ("P6518", "dont primes d'assurance (6518, garantie décès)", {ex: AG[str(ex)]["prime_assurance_6518"] for ex in EX}, "Balance, compte 6518 ; PCMS : 651 Primes d'assurances, classe 65"),
    ("CFB", "Charges de fonctionnement brut (61 à 68, 691, 697)", serie("CHB", "num"), "Numérateur de CHB"),
    ("PROD", "Total des produits (liste de l'annexe)", serie("CHB", "den"), "Dénominateur de CHB"),
    ("RN", "Excédent net de l'exercice", serie("EXN", "num"), "Numérateur de EXN"),
    ("RES", "Réserve obligatoire (111)", serie("ROB", "num"), "Numérateur de ROB"),
    ("FP", "Fonds propres et assimilés (10 à 15, 190)", serie("FEP", "num"), "Numérateur de FEP"),
    ("DET", "Dettes, y compris provisions techniques", serie("FEP", "den"), "Dénominateur de FEP"),
    ("TRE", "Trésorerie actif (classe 5 débitrice)", serie("LIQ", "num"), "Numérateur de LIQ et de RL3"),
    ("DCT", "Dettes à court terme (classes 4 et 5 créditrices, variation des provisions techniques)", serie("LIQ", "den"), "Dénominateur de LIQ, RL2 et RL3"),
    ("RL2N", "Valeurs réalisables et disponibles", serie("RL2", "num"), "Numérateur de RL2"),
    ("EMP", "Emplois moyens et longs (classe 2 nette)", serie("EQS", "num"), "Numérateur de EQS"),
    ("RESL", "Ressources moyennes et longues (classes 10 à 19)", serie("EQS", "den"), "Dénominateur de EQS"),
    ("ACTEP", "Actifs admis en couverture de l'épargne (banques et DAT dédiés, FCP)", serie("RL1", "num"), "Numérateur de RL1"),
    ("EPA", "Épargne des adhérents (41195100 PEP + 41195200 EFC)", serie("RL1", "den"), "Dénominateur de RL1"),
    ("RL4N", "Réserve de trésorerie des risques longs constituée", serie("RL4", "num"), "Aucun compte identifié"),
    ("RL4D", "Un trimestre de sorties d'épargne et de frais au prorata", serie("RL4", "den"), "Dénominateur de RL4"),
    ("FP2", "Fonds propres (10 à 14 + résultat)", serie("RL5", "num"), "Numérateur de RL5"),
    ("SORT2", "Sorties d'épargne des exercices n et n-1", serie("RL5", "den"), "Dénominateur de RL5"),
    ("P129", "Pertes reportées (129, solde débiteur)", {ex: -MS[ex]["pertes_reportees_129"] for ex in EX}, "Balance, compte 129 ; onglet 10_marge_solvabilite"),
    ("COTM", "Cotisations maladie (7010)", {ex: AG[str(ex)]["cot_maladie"] for ex in EX}, "Balance, compte 7010"),
    ("PREM", "Prestations maladie", {ex: AG[str(ex)]["prest_maladie"] for ex in EX}, "Balance, comptes 601, 6013, 6018, 60800007"),
    ("LIQr", "LIQ retraité (épargne en ressource longue)", serie("LIQ", "valeurs_retraitees"), "Onglet 10_ratios, valeurs retraitées"),
    ("RL2r", "RL2 retraité", serie("RL2", "valeurs_retraitees"), "Idem"),
    ("RL3r", "RL3 retraité", serie("RL3", "valeurs_retraitees"), "Idem"),
    ("EQSr", "EQS retraité (épargne et avances en emplois et ressources longs)", serie("EQS", "valeurs_retraitees"), "Idem"),
]
ws = wb.create_sheet("13c_donnees")
titre(ws, "Données des ratios, 2020 au 31/08/2026 (FCFA)",
      "Valeurs reprises du calcul des ratios (onglets 10_), lui-même fait à partir des balances reçues le 15/10/2026. 2026 : huit mois, balance provisoire.")
entete(ws, 4, ANS + ["Source"])
VR = {}
for i, (k, lib, s, src) in enumerate(VARS, 5):
    VR[k] = i
    cel(ws, i, 1, k, gras=True)
    cel(ws, i, 2, lib, wrap=True)
    for ex in EX:
        v = s.get(ex)
        fmt = FMT_R if k.endswith("r") else FMT_M
        cel(ws, i, C0 + EX.index(ex), None if v is None else (round(v, 6) if k.endswith("r") else round(v)), fmt)
    cel(ws, i, C0 + len(EX), src, wrap=True)
largeurs(ws, [9, 52] + [17] * 7 + [50])
ws.freeze_panes = "C5"


def V(k, ex):
    return f"'13c_donnees'!{COL[ex]}{VR[k]}"


# =============================================================================================== 13b paramètres
NORMES = [  # code, nom, borne basse, borne haute, strict, unité, source
    ("SIN", "Taux de sinistralité", 0.70, 0.80, False, "%", "Décision 05/2022, tableau 1, ligne 5"),
    ("ROB", "Ratio de réserve obligatoire (mois)", 6, 9, False, "mois", "Décision 05/2022, tableau 1, ligne 10"),
    ("FEP", "Ratio des fonds propres et assimilés", 1, None, False, "x", "Décision 05/2022, tableau 2, ligne 9"),
    ("LIQ", "Ratio de liquidité", 1, None, False, "x", "Décision 05/2022, tableau 2, ligne 8"),
    ("EQS", "Ratio d'équilibre stable", 0, 1, False, "x", "Décision 05/2022, tableau 1, ligne 11 (borne basse 0 : ressources négatives)"),
    ("EXN", "Taux d'excédent net", 0, 0.05, False, "%", "Décision 05/2022, tableau 2, ligne 10 (borne basse 0 : déficit)"),
    ("RL1", "Solvabilité de la branche épargne", 1, None, False, "x", "Décision 05/2022, section II, § 2.1"),
    ("RL2", "Trésorerie à échéance (risques longs)", 1, None, True, "x", "Décision 05/2022, section II, § 2.2 (strictement supérieur)"),
    ("RL3", "Trésorerie immédiate (risques longs)", 1, None, False, "x", "Décision 05/2022, section II, § 2.3"),
    ("RL4", "Réserve de trésorerie (trimestres)", 1, None, False, "x", "Règlement d'exécution 004/2022, art. 19 al. 3"),
    ("RL5", "Fonds propres / prestations longues", 1, None, False, "x", "Règlement d'exécution 004/2022, art. 21"),
    ("CHF", "Taux des charges de fonctionnement", None, 0.20, False, "%", "Décision 05/2022, tableau 1, ligne 2"),
    ("CHB", "Taux des charges de fonctionnement brut", None, 0.15, False, "%", "Décision 05/2022, tableau 1, ligne 3"),
    ("CUP", "Coût par unité monétaire prestée", None, 0.15, False, "%", "Décision 05/2022, tableau 2, ligne 5"),
]
ws = wb.create_sheet("13b_parametres", 1)
titre(ws, "Paramètres : normes du texte et leviers (cellules modifiables en orange)",
      "Borne absente : -1E+15 ou 1E+15 (pas de borne). Les cellules orange peuvent être modifiées ; tous les onglets 13 se recalculent.")
entete(ws, 4, ["Code", "Ratio", "Borne basse", "Borne haute", "Strict", "Unité", "Source"])
JAUNE = PatternFill("solid", fgColor="FFF2CC")
NR, P = {}, {}
for i, (code, nom, lo, hi, st, u, src) in enumerate(NORMES, 5):
    NR[code] = i
    cel(ws, i, 1, code, gras=True)
    cel(ws, i, 2, nom)
    f = FMT_P if u == "%" else FMT_R
    f = '[>=100000000000000]"pas de borne";[<=-100000000000000]"pas de borne";' + f
    cel(ws, i, 3, -1e15 if lo is None else lo, f)
    cel(ws, i, 4, 1e15 if hi is None else hi, f)
    cel(ws, i, 5, "Oui" if st else "Non")
    cel(ws, i, 6, u)
    cel(ws, i, 7, src, wrap=True)
com(ws.cell(NR["EQS"], 3), "Borne basse 0 : un ratio négatif traduit des ressources longues négatives (fonds propres négatifs en 2020) ; il est jugé non conforme, comme dans les versions 3 à 7 du rapport.")
com(ws.cell(NR["EXN"], 3), "Borne basse 0 : un excédent négatif (déficit de 2026) est jugé non conforme, comme dans les versions 3 à 7 du rapport.")


def LO(code):
    return f"'13b_parametres'!$C${NR[code]}"


def HI(code):
    return f"'13b_parametres'!$D${NR[code]}"


r = 5 + len(NORMES) + 1
entete(ws, r, ["Paramètre", "Libellé", "Valeur", "", "", "Unité", "Justification"])
PARAMS = [
    ("T0", "Taux de cotisation obligatoire actuel", 0.05, FMT_P, "%", "Statuts, art. 17 : 5 % de la solde de base (3,5 % + 1,5 %)."),
    ("CHOC", "Choc unitaire de comparaison des leviers", 100000000, FMT_M, "FCFA", "Convention de présentation (IAS 19 § 145 : une hypothèse à la fois) ; modifiable. Ce n'est pas une prévision."),
    ("D20", "Dotation minimale à la réserve légale", 0.20, FMT_P, "%", "Statuts, art. 90 : au moins 20 % à 50 % des excédents nets de l'exercice précédent (borne basse)."),
    ("D50", "Dotation maximale prévue par les statuts", 0.50, FMT_P, "%", "Statuts, art. 90 (borne haute)."),
    ("ANREF", "Exercice de référence des leviers", 2025, "0", "", "Dernier exercice complet (clos au 31/12/2025)."),
]
for k, (cle, lib, v, fmt, u, just) in enumerate(PARAMS, r + 1):
    cel(ws, k, 1, cle, gras=True)
    cel(ws, k, 2, lib)
    x = cel(ws, k, 3, v, fmt)
    x.fill = JAUNE
    cel(ws, k, 6, u)
    cel(ws, k, 7, just, wrap=True)
    P[cle] = f"'13b_parametres'!$C${k}"
largeurs(ws, [10, 44, 14, 14, 8, 8, 80])

# =============================================================================================== 13d écarts
FORM = {  # valeur du ratio, montant à combler, libellé de la variable d'ajustement, base pour le %
    "SIN": (lambda e: f"={V('PRE', e)}/{V('COT', e)}",
            lambda e: f"=IF({V('PRE', e)}<{LO('SIN')}*{V('COT', e)},{LO('SIN')}*{V('COT', e)}-{V('PRE', e)},IF({V('PRE', e)}>{HI('SIN')}*{V('COT', e)},{HI('SIN')}*{V('COT', e)}-{V('PRE', e)},0))",
            "Prestations à ajouter (+) ou à retirer (-)", "PRE"),
    "ROB": (lambda e: f"={V('RES', e)}*12/{V('PRE', e)}", lambda e: f"=MAX(0,{LO('ROB')}*{V('PRE', e)}/12-{V('RES', e)})", "Réserve obligatoire à constituer", "RES"),
    "FEP": (lambda e: f"={V('FP', e)}/{V('DET', e)}", lambda e: f"=MAX(0,{LO('FEP')}*{V('DET', e)}-{V('FP', e)})", "Fonds propres à constituer", "FP"),
    "LIQ": (lambda e: f"={V('TRE', e)}/{V('DCT', e)}", lambda e: f"=MAX(0,{LO('LIQ')}*{V('DCT', e)}-{V('TRE', e)})", "Trésorerie à constituer", "TRE"),
    "EQS": (lambda e: f"={V('EMP', e)}/{V('RESL', e)}", lambda e: f"=MAX(0,{V('EMP', e)}-{HI('EQS')}*{V('RESL', e)})", "Ressources longues à ajouter (ou emplois longs à céder)", "RESL"),
    "EXN": (lambda e: f"={V('RN', e)}/{V('COT', e)}", lambda e: f"=IF({V('RN', e)}<0,-{V('RN', e)},MAX(0,{V('RN', e)}-{HI('EXN')}*{V('COT', e)}))", "Excédent à réduire (déficit à combler si négatif)", "RN"),
    "RL1": (lambda e: f"={V('ACTEP', e)}/{V('EPA', e)}", lambda e: f"=MAX(0,{LO('RL1')}*{V('EPA', e)}-{V('ACTEP', e)})", "Actifs admis à affecter à l'épargne", "ACTEP"),
    "RL2": (lambda e: f"={V('RL2N', e)}/{V('DCT', e)}", lambda e: f"=MAX(0,{LO('RL2')}*{V('DCT', e)}-{V('RL2N', e)})", "Valeurs réalisables ou disponibles à constituer", "RL2N"),
    "RL3": (lambda e: f"={V('TRE', e)}/{V('DCT', e)}", lambda e: f"=MAX(0,{LO('RL3')}*{V('DCT', e)}-{V('TRE', e)})", "Disponibilités à constituer", "TRE"),
    "RL4": (lambda e: f"={V('RL4N', e)}/{V('RL4D', e)}", lambda e: f"=MAX(0,{LO('RL4')}*{V('RL4D', e)}-{V('RL4N', e)})", "Réserve de trésorerie à constituer", "RL4D"),
    "RL5": (lambda e: f"=IF({V('SORT2', e)}=\"\",\"\",{V('FP2', e)}/{V('SORT2', e)})", lambda e: f"=IF({V('SORT2', e)}=\"\",\"\",MAX(0,{LO('RL5')}*{V('SORT2', e)}-{V('FP2', e)}))", "Fonds propres à constituer", "SORT2"),
    "CHF": (lambda e: f"={V('CF', e)}/{V('COT', e)}", lambda e: f"=MAX(0,{V('CF', e)}-{HI('CHF')}*{V('COT', e)})", "Charges de fonctionnement à réduire", "CF"),
    "CHB": (lambda e: f"={V('CFB', e)}/{V('PROD', e)}", lambda e: f"=MAX(0,{V('CFB', e)}-{HI('CHB')}*{V('PROD', e)})", "Charges brutes à réduire", "CFB"),
    "CUP": (lambda e: f"={V('CF', e)}/{V('PRE', e)}", lambda e: f"=MAX(0,{V('CF', e)}-{HI('CUP')}*{V('PRE', e)})", "Charges de fonctionnement à réduire", "CF"),
}
EXPL_V = {
    "SIN": "Sinistralité = charges de prestations (60 + 6901 - 7901) / cotisations acquises (70 + 7906 - 6906). Norme : entre 70 % et 80 %.",
    "ROB": "Réserve obligatoire en mois = réserve (111) × 12 / charges de prestations. Norme : entre 6 et 9 mois.",
    "FEP": "Fonds propres et assimilés (10 à 15, 190) / dettes y compris provisions techniques. Norme : au moins 1.",
    "LIQ": "Liquidité = trésorerie actif (classe 5 débitrice) / (dettes à court terme + variation des provisions techniques). Norme : au moins 1.",
    "EQS": "Équilibre stable = emplois moyens et longs (classe 2 nette) / ressources moyennes et longues (10 à 19). Norme : au plus 1 ; négatif = non conforme.",
    "EXN": "Excédent net / cotisations acquises. Norme : au plus 5 % ; un déficit (valeur négative) est jugé non conforme.",
    "RL1": "Actifs admis dédiés à l'épargne / épargne des adhérents (41195100 + 41195200). Norme : au moins 1.",
    "RL2": "(Valeurs réalisables + disponibles) / dettes à court terme. Norme : strictement supérieur à 1.",
    "RL3": "Valeurs disponibles / dettes à court terme. Norme : au moins 1.",
    "RL4": "Réserve de trésorerie constituée / un trimestre de sorties d'épargne et de frais. Norme : au moins 1 trimestre (règlement d'exécution 004/2022, art. 19).",
    "RL5": "Fonds propres / sorties d'épargne des deux derniers exercices. Norme : au moins 1 (règlement d'exécution 004/2022, art. 21). 2020 : exercice 2019 non reçu.",
    "CHF": "Charges de fonctionnement (612, 613, 618, 62 à 68) / cotisations acquises. Norme : au plus 20 %.",
    "CHB": "Charges de fonctionnement brut (61 à 68, 691, 697) / total des produits. Norme : au plus 15 %.",
    "CUP": "Charges de fonctionnement / charges de prestations. Norme : au plus 15 %.",
}
EXPL_G = {
    "SIN": "Si la sinistralité est sous 70 % : 70 % × cotisations - prestations (prestations à ajouter). Si elle dépasse 80 % : 80 % × cotisations - prestations (montant négatif). Sinon 0.",
    "ROB": "Montant = max(0 ; 6 / 12 × prestations - réserve) : dotation qui porte la réserve à six mois de prestations.",
    "FEP": "Montant = max(0 ; 1 × dettes - fonds propres), à dettes inchangées.",
    "LIQ": "Montant = max(0 ; 1 × dettes à court terme - trésorerie).",
    "EQS": "Montant = max(0 ; emplois longs - 1 × ressources longues) : ressources longues à ajouter, ou emplois longs à céder.",
    "EXN": "Si l'excédent est négatif : déficit à combler (- excédent). Sinon max(0 ; excédent - 5 % × cotisations) : excédent à réduire.",
    "RL1": "Montant = max(0 ; épargne - actifs admis dédiés).",
    "RL2": "Montant = max(0 ; dettes à court terme - valeurs réalisables et disponibles) ; le texte exige un ratio strictement supérieur à 1.",
    "RL3": "Montant = max(0 ; dettes à court terme - disponibilités).",
    "RL4": "Montant = max(0 ; un trimestre de sorties et de frais - réserve de trésorerie constituée).",
    "RL5": "Montant = max(0 ; sorties d'épargne de deux exercices - fonds propres).",
    "CHF": "Montant = max(0 ; charges de fonctionnement - 20 % × cotisations).",
    "CHB": "Montant = max(0 ; charges brutes - 15 % × total des produits).",
    "CUP": "Montant = max(0 ; charges de fonctionnement - 15 % × prestations).",
}
ws = wb.create_sheet("13d_ecarts")
titre(ws, "Statut de chaque ratio et montant à combler, 2020 au 31/08/2026",
      "Montant à combler = montant qui ramène le ratio sur la borne la plus proche du texte, les autres termes restant inchangés (formule du texte inversée). "
      "Structurel : non conforme toutes les années mesurées. Récurrent : non conforme au moins deux exercices de jugement 2022 à 2025 (Règlement 07/2009, art. 84).")
entete(ws, 4, ANS + ["Années non conformes", "dont 2022 à 2025", "Années mesurées", "Minimum", "Maximum", "Moyenne", "Qualification"])
D = {}
r = 5
for code, nom, *_ in NORMES:
    fv, fg, lib, base = FORM[code]
    u = next(n[5] for n in NORMES if n[0] == code)
    fmt = FMT_P if u == "%" else FMT_R
    rv, rs, rg, rp = r, r + 1, r + 2, r + 3
    D[code] = {"val": rv, "st": rs, "gap": rg, "pct": rp}
    cel(ws, rv, 1, code, gras=True)
    cel(ws, rv, 2, f"{nom} : valeur", gras=True)
    cel(ws, rs, 2, "Statut")
    cel(ws, rg, 2, f"Montant à combler (FCFA) : {lib.lower()}", wrap=True)
    cel(ws, rp, 2, f"Montant à combler en % de {base}", wrap=True)
    com(ws.cell(rv, 2), EXPL_V[code] + " Source : " + next(n[6] for n in NORMES if n[0] == code) + ".")
    com(ws.cell(rs, 2), "Statut = « Conforme » si la valeur est comprise entre la borne basse et la borne haute de l'onglet 13b (borne basse stricte pour RL2), « Non conforme » sinon, « En attente » si la valeur manque.")
    com(ws.cell(rg, 2), EXPL_G[code] + " Les autres termes du ratio restent inchangés.")
    com(ws.cell(rp, 2), f"Montant à combler rapporté à la variable {base} de l'onglet 13c (valeur absolue), pour lire l'effort en pourcentage.")
    strict = next(n[4] for n in NORMES if n[0] == code)
    for ex in EX:
        c = COL[ex]
        cel(ws, rv, C0 + EX.index(ex), fv(ex), fmt)
        cmp_lo = ">" if strict else ">="
        cel(ws, rs, C0 + EX.index(ex),
            f'=IF({c}{rv}="","En attente",IF(AND({c}{rv}{cmp_lo}{LO(code)},{c}{rv}<={HI(code)}),"Conforme","Non conforme"))')
        cel(ws, rg, C0 + EX.index(ex), fg(ex), FMT_M)
        cel(ws, rp, C0 + EX.index(ex), f'=IF(OR({c}{rg}="",{V(base, ex)}=0),"",{c}{rg}/ABS({V(base, ex)}))', FMT_P)
    a, b = COL[2020], COL[2026]
    cel(ws, rv, 10, f'=COUNTIF({a}{rs}:{b}{rs},"Non*")', "0", gras=True)
    cel(ws, rv, 11, f'=COUNTIF({COL[2022]}{rs}:{COL[2025]}{rs},"Non*")', "0")
    cel(ws, rv, 12, f'=COUNTIF({a}{rs}:{b}{rs},"Conforme")+COUNTIF({a}{rs}:{b}{rs},"Non*")', "0")
    cel(ws, rv, 13, f"=MIN({a}{rg}:{b}{rg})", FMT_M)
    cel(ws, rv, 14, f"=MAX({a}{rg}:{b}{rg})", FMT_M)
    cel(ws, rv, 15, f"=AVERAGE({a}{rg}:{b}{rg})", FMT_M)
    cel(ws, rv, 16, f'=IF(AND(J{rv}>0,J{rv}=L{rv}),"Structurel",IF(K{rv}>=2,"Récurrent",IF(J{rv}>0,"Ponctuel","Conforme")))', gras=True)
    for rr in (rv, rs, rg, rp):
        ws.row_dimensions[rr].height = 28 if rr == rg else 16
    r += 5
com(ws["J4"], "Nombre d'années non conformes de 2020 au 31/08/2026 (formule NB.SI sur la ligne Statut).")
com(ws["P4"], "Structurel : non conforme toutes les années mesurées (colonnes J = L). Récurrent : au moins deux exercices non conformes de 2022 à 2025 (Règlement 07/2009, art. 84 : non-respect sur deux exercices). Sinon ponctuel.")
com(ws["M4"], "Minimum, maximum et moyenne du montant à combler sur 2020 à 2026 (2026 : huit mois, indicatif).")
largeurs(ws, [7, 46] + [16] * 7 + [11, 10, 10, 16, 16, 16, 13])
ws.freeze_panes = "C5"


def DD(code, cle, ex):
    return f"'13d_ecarts'!{COL[ex]}{D[code][cle]}"


def DS(code, col):
    return f"'13d_ecarts'!{col}{D[code]['val']}"


# =============================================================================================== 13e facteurs récurrents
FACT = [
    ("F1", "Épargne des adhérents non adossée à des actifs liquides dédiés",
     "41195100 et 41195200 (épargne) en dettes à court terme ; classe 5 et DAT dédiés", ["LIQ", "RL1", "RL2", "RL3", "RL4", "EQS"]),
    ("F2", "Pertes reportées non apurées", "129 (report à nouveau débiteur) dans les fonds propres", ["FEP", "RL5"]),
    ("F3", "Réserve obligatoire absente puis insuffisante", "111 (réserve légale) ; affectations des excédents", ["ROB"]),
    ("F4", "Charges de fonctionnement au-delà des plafonds de la Décision (cotisations et prestations)", "612, 613, 618, 62 à 68 (dont 6518, prime de la garantie décès, en 2020)", ["CHF", "CHB", "CUP"]),
    ("F5", "Prestations faibles par rapport aux cotisations (couverture limitée à l'hospitalisation, provisions figées)", "60 (prestations), 70 (cotisations), 6901 et 7901 absents", ["SIN", "EXN"]),
]
CODES = [n[0] for n in NORMES]
ws = wb.create_sheet("13e_facteurs")
titre(ws, "Facteurs récurrents : quels comptes expliquent les écarts chaque année ?",
      "Une croix indique que le facteur entre dans la formule du ratio (numérateur ou dénominateur). La dernière colonne compte les ratios du facteur qualifiés de structurels (onglet 13d).")
entete(ws, 4, ["Facteur", "Description", "Comptes en cause"] + CODES + ["Ratios structurels / ratios du facteur"])
FR = {}
for i, (f, desc, cpt, liste) in enumerate(FACT, 5):
    FR[f] = i
    cel(ws, i, 1, f, gras=True)
    cel(ws, i, 2, desc, wrap=True)
    cel(ws, i, 3, cpt, wrap=True)
    for j, code in enumerate(CODES, 4):
        x = cel(ws, i, j, "X" if code in liste else "")
        x.alignment = Alignment(horizontal="center", vertical="center")
    somme = "+".join(f'IF({DS(c, "P")}="Structurel",1,0)' for c in liste)
    cel(ws, i, 4 + len(CODES), f'=({somme})&" / {len(liste)}"', gras=True)
    ws.row_dimensions[i].height = 44
# preuves par année
r = 5 + len(FACT) + 2
ws.cell(r, 1, "Preuves par année (FCFA ou %)").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Facteur", "Indicateur"] + [str(e) for e in EX])
PREUVES = [
    ("F1", "Épargne des adhérents (PEP + EFC)", lambda e: f"={V('EPA', e)}", FMT_M),
    ("F1", "Actifs admis dédiés à l'épargne", lambda e: f"={V('ACTEP', e)}", FMT_M),
    ("F1", "Couverture de l'épargne par les actifs dédiés", lambda e: f"={V('ACTEP', e)}/{V('EPA', e)}", FMT_P),
    ("F1", "Part de l'épargne dans les dettes à court terme", lambda e: f"={V('EPA', e)}/{V('DCT', e)}", FMT_P),
    ("F2", "Pertes reportées (129)", lambda e: f"={V('P129', e)}", FMT_M),
    ("F2", "Pertes reportées / dettes", lambda e: f"={V('P129', e)}/{V('DET', e)}", FMT_P),
    ("F3", "Réserve obligatoire (111)", lambda e: f"={V('RES', e)}", FMT_M),
    ("F3", "Six mois de prestations (borne basse de ROB)", lambda e: f"={LO('ROB')}*{V('PRE', e)}/12", FMT_M),
    ("F4", "Charges de fonctionnement", lambda e: f"={V('CF', e)}", FMT_M),
    ("F4", "dont prime de la garantie décès (6518)", lambda e: f"={V('P6518', e)}", FMT_M),
    ("F4", "Part de la prime décès dans les charges de fonctionnement", lambda e: f"={V('P6518', e)}/{V('CF', e)}", FMT_P),
    ("F5", "Prestations / cotisations (toutes garanties)", lambda e: f"={V('PRE', e)}/{V('COT', e)}", FMT_P),
    ("F5", "Prestations maladie / cotisations maladie", lambda e: f"={V('PREM', e)}/{V('COTM', e)}", FMT_P),
]
for k, (f, lib, fx, fmt) in enumerate(PREUVES, r + 2):
    cel(ws, k, 1, f, gras=True)
    cel(ws, k, 2, lib)
    for ex in EX:
        cel(ws, k, C0 + EX.index(ex), fx(ex), fmt)
largeurs(ws, [8, 50, 40] + [7] * len(CODES) + [16])
for j in range(3, 10):
    ws.column_dimensions[L(j)].width = max(ws.column_dimensions[L(j)].width or 0, 16)

# =============================================================================================== 13f leviers : montants d'équilibre
LEV = [
    ("L1", "Taux de cotisation", "Statuts, art. 17 et 93 ; règlement intérieur, art. 86", "Variation des cotisations acquises (FCFA) qui ramène le ratio sur la borne ; signe négatif : baisse."),
    ("L2", "Régime des prestations", "Statuts, art. 93 ; règlement intérieur, art. 41, 45 et 86", "Variation des prestations (FCFA) qui ramène le ratio sur la borne ; signe négatif : baisse."),
    ("L3", "Charges de fonctionnement", "Décision 05/2022, tableau 1, lignes 2 et 3 ; tableau 2, ligne 5", "Réduction des charges (FCFA) qui ramène le ratio sur la borne."),
    ("L4", "Dotation à la réserve légale", "Statuts, art. 90 ; Décision 05/2022, tableau 1, ligne 10", "Réserve à doter (FCFA) ; comparaison avec les dotations de 20 % et 50 % des excédents passés."),
    ("L5", "Réallocation des placements", "Statuts, art. 85 ; règlement d'exécution 004/2022, art. 17 et 19", "Actifs longs à transformer en disponibilités dédiées (FCFA) ; comparaison avec toute la classe 2."),
    ("L6", "Classement de l'épargne en ressource longue", "Règlement d'exécution 004/2022, art. 3 (risques longs) ; PCMS", "Reclassement de présentation, sans effet économique (valeurs retraitées)."),
    ("L7", "Apport en fonds propres", "Statuts, art. 79 et 81 (dons et legs) ; règlement d'exécution 004/2022, art. 21", "Fonds propres à apporter (FCFA) ; nombre d'années d'excédent moyen nécessaires."),
]
EQ = [  # levier, ratio, formule de l'équilibre
    ("L1", "SIN", lambda e: f"=IF({V('PRE', e)}/{V('COT', e)}<{LO('SIN')},{V('PRE', e)}/{LO('SIN')}-{V('COT', e)},IF({V('PRE', e)}/{V('COT', e)}>{HI('SIN')},{V('PRE', e)}/{HI('SIN')}-{V('COT', e)},0))"),
    ("L1", "CHF", lambda e: f"=IF({V('CF', e)}/{V('COT', e)}>{HI('CHF')},{V('CF', e)}/{HI('CHF')}-{V('COT', e)},0)"),
    ("L1", "CHB", lambda e: f"=IF({V('CFB', e)}/{V('PROD', e)}>{HI('CHB')},{V('CFB', e)}/{HI('CHB')}-{V('PROD', e)},0)"),
    ("L1", "EXN", lambda e: f"=IF({V('RN', e)}<0,-{V('RN', e)},IF({V('RN', e)}>{HI('EXN')}*{V('COT', e)},({HI('EXN')}*{V('COT', e)}-{V('RN', e)})/(1-{HI('EXN')}),0))"),
    ("L2", "SIN", lambda e: f"={DD('SIN', 'gap', e)}"),
    ("L2", "CUP", lambda e: f"=IF({V('CF', e)}>{HI('CUP')}*{V('PRE', e)},{V('CF', e)}/{HI('CUP')}-{V('PRE', e)},0)"),
    ("L2", "EXN", lambda e: f"=IF({V('RN', e)}<0,{V('RN', e)},IF({V('RN', e)}>{HI('EXN')}*{V('COT', e)},{V('RN', e)}-{HI('EXN')}*{V('COT', e)},0))"),
    ("L2", "ROB", lambda e: f"=IF({V('RES', e)}*12/{V('PRE', e)}<{LO('ROB')},{V('RES', e)}*12/{LO('ROB')}-{V('PRE', e)},0)"),
    ("L3", "CHF", lambda e: f"={DD('CHF', 'gap', e)}"),
    ("L3", "CHB", lambda e: f"={DD('CHB', 'gap', e)}"),
    ("L3", "CUP", lambda e: f"={DD('CUP', 'gap', e)}"),
    ("L4", "ROB", lambda e: f"={DD('ROB', 'gap', e)}"),
    ("L5", "LIQ", lambda e: f"={DD('LIQ', 'gap', e)}"),
    ("L5", "RL1", lambda e: f"={DD('RL1', 'gap', e)}"),
    ("L5", "RL2", lambda e: f"={DD('RL2', 'gap', e)}"),
    ("L5", "RL3", lambda e: f"={DD('RL3', 'gap', e)}"),
    ("L5", "RL4", lambda e: f"={DD('RL4', 'gap', e)}"),
    ("L5", "EQS", lambda e: f"={DD('EQS', 'gap', e)}"),
    ("L7", "FEP", lambda e: f"={DD('FEP', 'gap', e)}"),
    ("L7", "RL5", lambda e: f"={DD('RL5', 'gap', e)}"),
]
EXPL_EQ = {
    ("L1", "SIN"): "Variation des cotisations qui amène prestations / cotisations sur la borne la plus proche : prestations / borne - cotisations. Négatif : baisse des cotisations.",
    ("L1", "CHF"): "Variation des cotisations qui amène charges / cotisations à 20 % : charges / 20 % - cotisations.",
    ("L1", "CHB"): "Variation des produits (cotisations) qui amène charges brutes / produits à 15 % : charges brutes / 15 % - produits.",
    ("L1", "EXN"): "Une cotisation de plus augmente aussi l'excédent : (5 % × cotisations - excédent) / (1 - 5 %). Déficit : variation égale au déficit.",
    ("L2", "SIN"): "Prestations à ajouter pour atteindre 70 % des cotisations (montant à combler de l'onglet 13d).",
    ("L2", "CUP"): "Prestations à ajouter pour que les charges ne dépassent pas 15 % des prestations : charges / 15 % - prestations.",
    ("L2", "EXN"): "Prestations à ajouter pour ramener l'excédent à 5 % des cotisations : excédent - 5 % × cotisations. Déficit : montant négatif (prestations à réduire).",
    ("L2", "ROB"): "Les prestations sont au dénominateur de ROB : variation qui donnerait 6 mois avec la réserve actuelle = réserve × 12 / 6 - prestations (négatif : levier inadapté).",
    ("L3", "CHF"): "Réduction des charges de fonctionnement pour atteindre 20 % des cotisations.",
    ("L3", "CHB"): "Réduction des charges brutes pour atteindre 15 % du total des produits.",
    ("L3", "CUP"): "Réduction des charges de fonctionnement pour atteindre 15 % des prestations.",
    ("L4", "ROB"): "Dotation à la réserve légale pour atteindre 6 mois de prestations (statuts, art. 90).",
    ("L5", "LIQ"): "Disponibilités à constituer par réallocation d'actifs longs pour une liquidité de 1.",
    ("L5", "RL1"): "Actifs à affecter à l'épargne pour une couverture de 1.",
    ("L5", "RL2"): "Valeurs réalisables ou disponibles à constituer pour un ratio de 1.",
    ("L5", "RL3"): "Disponibilités à constituer pour un ratio de 1.",
    ("L5", "RL4"): "Réserve de trésorerie d'un trimestre à constituer (règlement d'exécution 004/2022, art. 19).",
    ("L5", "EQS"): "Emplois longs à céder (ou ressources longues à ajouter) pour un ratio de 1.",
    ("L7", "FEP"): "Fonds propres à apporter pour un ratio de 1, à dettes inchangées.",
    ("L7", "RL5"): "Fonds propres à apporter pour couvrir les sorties d'épargne de deux exercices.",
}
ws = wb.create_sheet("13f_leviers")
titre(ws, "Leviers tirés des textes : montants d'équilibre par ratio et par année (FCFA)",
      "Chaque ligne donne le montant du levier qui ramène un ratio sur sa borne, toutes choses égales par ailleurs (IAS 19 § 145). Un même levier peut améliorer un ratio et en dégrader un autre : "
      "les signes opposés d'une même ligne de levier (L1, L2) révèlent ces conflits.")
entete(ws, 4, ["Levier", "Texte", "Ratio visé"] + [str(e) for e in EX] + ["Années où un effort est requis", "Moyenne"])
for j, w in enumerate([9, 44, 10] + [16] * 7 + [12, 16], 1):
    ws.column_dimensions[L(j)].width = w
EQR = {}
r = 5
for lv, nom, txt, expl in LEV:
    x = cel(ws, r, 1, lv, gras=True)
    x.fill = HS
    y = cel(ws, r, 2, f"{nom} : {expl}", gras=True, wrap=True)
    y.fill = HS
    cel(ws, r, 3, txt, wrap=True).fill = HS
    ws.merge_cells(start_row=r, end_row=r, start_column=3, end_column=12)
    ws.row_dimensions[r].height = 30
    r += 1
    for (l2, code, fx) in [q for q in EQ if q[0] == lv]:
        EQR[(lv, code)] = r
        cel(ws, r, 1, lv)
        x_ = cel(ws, r, 2, f"Montant d'équilibre pour {code}")
        com(x_, EXPL_EQ.get((lv, code), "Reprise du montant à combler de l'onglet 13d pour ce ratio."))
        cel(ws, r, 3, code, gras=True)
        for ex in EX:
            cel(ws, r, 4 + EX.index(ex), fx(ex), FMT_M)
        a, b = L(4), L(4 + len(EX) - 1)
        cel(ws, r, 11, f'=COUNTIF({a}{r}:{b}{r},"<>0")-COUNTBLANK({a}{r}:{b}{r})', "0")
        cel(ws, r, 12, f"=AVERAGE({a}{r}:{b}{r})", FMT_M)
        r += 1
    if lv == "L1":
        cel(ws, r, 2, "Taux de cotisation équivalent pour SIN (base 5 %)")
        for ex in EX:
            cel(ws, r, 4 + EX.index(ex), f"={P['T0']}*({V('COT', ex)}+{L(4 + EX.index(ex))}{EQR[('L1', 'SIN')]})/{V('COT', ex)}", FMT_P)
        EQR[("L1", "T_SIN")] = r
        r += 1
        cel(ws, r, 2, "Taux de cotisation équivalent pour CHF (base 5 %)")
        for ex in EX:
            cel(ws, r, 4 + EX.index(ex), f"={P['T0']}*({V('COT', ex)}+{L(4 + EX.index(ex))}{EQR[('L1', 'CHF')]})/{V('COT', ex)}", FMT_P)
        EQR[("L1", "T_CHF")] = r
        com(ws.cell(r, 2), "Taux = 5 % × (cotisations + variation) / cotisations : l'assiette (solde de base) et le recouvrement sont supposés inchangés.")
        r += 1
    if lv == "L3":
        cel(ws, r, 2, "Réduction requise pour CUP, en % des charges de fonctionnement")
        for ex in EX:
            c = L(4 + EX.index(ex))
            cel(ws, r, 4 + EX.index(ex), f"={c}{EQR[('L3', 'CUP')]}/{V('CF', ex)}", FMT_P)
        EQR[("L3", "PCT_CUP")] = r
        r += 1
    if lv == "L4":
        cel(ws, r, 2, "Excédents positifs cumulés depuis 2020 (exercices précédents)")
        for ex in EX:
            prec = [e for e in EX if e < ex]
            f = "=" + ("+".join(f"MAX(0,{V('RN', e)})" for e in prec) if prec else "0")
            cel(ws, r, 4 + EX.index(ex), f, FMT_M)
        cum = r
        com(ws.cell(r, 2), "Somme des excédents nets positifs des exercices antérieurs, à partir de 2020 (exercices antérieurs à 2020 non reçus). Les statuts (art. 90) assoient la dotation sur l'excédent net de l'exercice précédent.")
        r += 1
        for cle, txt in (("D20", "20 %"), ("D50", "50 %")):
            cel(ws, r, 2, f"ROB si la réserve avait reçu {txt} des excédents passés (mois)")
            for ex in EX:
                c = L(4 + EX.index(ex))
                cel(ws, r, 4 + EX.index(ex), f"=MAX({V('RES', ex)},{P[cle]}*{c}{cum})*12/{V('PRE', ex)}", FMT_R)
            EQR[("L4", cle)] = r
            r += 1
            cel(ws, r, 2, f"Statut de ROB avec {txt}")
            for ex in EX:
                c = L(4 + EX.index(ex))
                cel(ws, r, 4 + EX.index(ex), f'=IF(AND({c}{r - 1}>={LO("ROB")},{c}{r - 1}<={HI("ROB")}),"Conforme","Non conforme")')
            cel(ws, r, 11, f'=COUNTIF({L(4)}{r}:{L(10)}{r},"Conforme")', "0")
            EQR[("L4", cle + "_st")] = r
            r += 1
        EQR[("L4", "CUM")] = cum
        com(ws.cell(r - 2, 2), "Réserve hypothétique = maximum entre la réserve comptabilisée et la dotation statutaire cumulée ; ROB = réserve × 12 / prestations. Les pertes reportées (129) ont en réalité absorbé une partie des excédents (affectations 2023 et 2024).")
    if lv == "L5":
        cel(ws, r, 2, "Classe 2 nette disponible (plafond théorique de réallocation)")
        for ex in EX:
            cel(ws, r, 4 + EX.index(ex), f"={V('EMP', ex)}", FMT_M)
        r += 1
        cel(ws, r, 2, "LIQ si toute la classe 2 était transformée en disponibilités")
        for ex in EX:
            cel(ws, r, 4 + EX.index(ex), f"=({V('TRE', ex)}+{V('EMP', ex)})/{V('DCT', ex)}", FMT_R)
        EQR[("L5", "MAX")] = r
        com(ws.cell(r, 2), "Borne théorique : la classe 2 comprend le siège, les terrains et des créances longues difficilement mobilisables. Si même ce maximum laisse LIQ sous 1, la réallocation seule ne peut pas suffire.")
        r += 1
        cel(ws, r, 2, "Statut de LIQ au maximum théorique")
        for ex in EX:
            c = L(4 + EX.index(ex))
            cel(ws, r, 4 + EX.index(ex), f'=IF({c}{r - 1}>={LO("LIQ")},"Conforme","Non conforme")')
        EQR[("L5", "MAXST")] = r
        r += 1
    if lv == "L6":
        for code, k in (("LIQ", "LIQr"), ("RL2", "RL2r"), ("RL3", "RL3r"), ("EQS", "EQSr")):
            cel(ws, r, 2, f"{code} retraité")
            cel(ws, r, 3, code, gras=True)
            for ex in EX:
                cel(ws, r, 4 + EX.index(ex), f"={V(k, ex)}", FMT_R)
            EQR[("L6", code)] = r
            r += 1
            cel(ws, r, 2, f"Statut de {code} retraité")
            for ex in EX:
                c = L(4 + EX.index(ex))
                cmp_lo = ">" if code == "RL2" else ">="
                cel(ws, r, 4 + EX.index(ex), f'=IF(AND({c}{r - 1}{cmp_lo}{LO(code)},{c}{r - 1}<={HI(code)}),"Conforme","Non conforme")')
            cel(ws, r, 11, f'=COUNTIF({L(4)}{r}:{L(10)}{r},"Conforme")', "0")
            EQR[("L6", code + "_st")] = r
            r += 1
    if lv == "L7":
        cel(ws, r, 2, "Excédent net moyen 2023 à 2025")
        cel(ws, r, 4, f"=AVERAGE({V('RN', 2023)},{V('RN', 2024)},{V('RN', 2025)})", FMT_M)
        moy = r
        r += 1
        cel(ws, r, 2, "Années d'excédent moyen nécessaires pour FEP")
        for ex in EX:
            c = L(4 + EX.index(ex))
            cel(ws, r, 4 + EX.index(ex), f"={c}{EQR[('L7', 'FEP')]}/$D${moy}", FMT_R)
        EQR[("L7", "ANS")] = r
        EQR[("L7", "MOY")] = moy
        r += 1
    r += 1
ws.freeze_panes = "D5"

# =============================================================================================== 13h synthèse : choc et faisabilité
ws = wb.create_sheet("13h_synthese")
titre(ws, "Effet d'un choc de 100 000 000 FCFA sur chaque ratio (exercice de référence) et faisabilité conjointe des normes",
      "Écritures du choc en partie double : L1 cotisations encaissées (+ produits, + résultat, + fonds propres, + trésorerie) ; L2 prestations payées (inverse) ; "
      "L3 charges de fonctionnement évitées (+ résultat, + trésorerie) ; L4 dotation à la réserve légale (affectation interne aux fonds propres) ; "
      "L5 placement long transformé en disponibilités dédiées à l'épargne et désignées comme réserve de trésorerie ; L7 apport en fonds propres (dons et legs) encaissé.")
AN = P["ANREF"]


def X(k):  # valeur de la variable k pour l'exercice de référence
    return f"INDEX('13c_donnees'!$C${VR[k]}:$I${VR[k]},1,{AN}-2019)"


Q = P["CHOC"]
BASE = {
    "SIN": "{PRE}/{COT}", "ROB": "{RES}*12/{PRE}", "FEP": "{FP}/{DET}", "LIQ": "{TRE}/{DCT}", "EQS": "{EMP}/{RESL}",
    "EXN": "{RN}/{COT}", "RL1": "{ACTEP}/{EPA}", "RL2": "{RL2N}/{DCT}", "RL3": "{TRE}/{DCT}", "RL4": "{RL4N}/{RL4D}",
    "RL5": "{FP2}/{SORT2}", "CHF": "{CF}/{COT}", "CHB": "{CFB}/{PROD}", "CUP": "{CF}/{PRE}",
}
CHOCS = {  # variations des variables
    "Avant choc": {},
    "L1 cotisations +": {"COT": "+", "PROD": "+", "RN": "+", "FP": "+", "TRE": "+", "RL2N": "+", "RESL": "+", "FP2": "+"},
    "L2 prestations +": {"PRE": "+", "RN": "-", "FP": "-", "TRE": "-", "RL2N": "-", "RESL": "-", "FP2": "-"},
    "L3 charges -": {"CF": "-", "CFB": "-", "RN": "+", "FP": "+", "TRE": "+", "RL2N": "+", "RESL": "+", "FP2": "+"},
    "L4 réserve légale +": {"RES": "+"},
    "L5 réallocation": {"EMP": "-", "TRE": "+", "ACTEP": "+", "RL2N": "+", "RL4N": "+"},
    "L7 fonds propres +": {"FP": "+", "TRE": "+", "RL2N": "+", "RESL": "+", "FP2": "+"},
}


def expr(code, delta):
    s = BASE[code]
    for k in sorted({w.strip("{}") for w in s.replace("/", " ").replace("*", " ").split() if w.startswith("{")}, key=len, reverse=True):
        t = X(k)
        if k in delta:
            t = f"({t}{delta[k]}{Q})"
        s = s.replace("{" + k + "}", t)
    return "=" + s


entete(ws, 4, ["Choc"] + CODES)
CR = {}
EXPL_CH = {
    "L1 cotisations +": "Écritures : + cotisations, + produits, + excédent, + fonds propres, + trésorerie, + valeurs réalisables, + ressources longues (statuts, art. 17 et 93).",
    "L2 prestations +": "Écritures : + prestations, - excédent, - fonds propres, - trésorerie, - valeurs réalisables, - ressources longues (statuts, art. 93).",
    "L3 charges -": "Écritures : - charges de fonctionnement, - charges brutes, + excédent, + fonds propres, + trésorerie, + ressources longues.",
    "L4 réserve légale +": "Écriture : + réserve légale (111), par affectation interne aux fonds propres ; seul ROB change (statuts, art. 90).",
    "L5 réallocation": "Écritures : - classe 2, + trésorerie, + actifs dédiés à l'épargne, + valeurs réalisables, + réserve de trésorerie (hypothèse : fonds désignés comme réserve) ; statuts, art. 85.",
    "L7 fonds propres +": "Écritures : + fonds propres (dons et legs, statuts, art. 79 et 81), + trésorerie, + valeurs réalisables, + ressources longues.",
}
for i, (nom, delta) in enumerate(CHOCS.items(), 5):
    CR[nom] = i
    x_ = cel(ws, i, 1, nom, gras=True)
    if nom in EXPL_CH:
        com(x_, EXPL_CH[nom] + " Montant : paramètre CHOC de l'onglet 13b ; exercice : paramètre ANREF.")
    for j, code in enumerate(CODES, 2):
        u = next(n[5] for n in NORMES if n[0] == code)
        cel(ws, i, j, expr(code, delta), FMT_P if u == "%" else FMT_R)
com(ws["A5"], "Valeurs de l'exercice de référence (paramètre ANREF de l'onglet 13b) avant tout choc.")
r = 5 + len(CHOCS) + 1
ws.cell(r, 1, "Sens de l'effet (distance à la norme avant et après le choc)").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Levier"] + CODES)
SENS = {}
for k, nom in enumerate([n for n in CHOCS if n != "Avant choc"], r + 2):
    SENS[nom] = k
    cel(ws, k, 1, nom, gras=True)
    for j, code in enumerate(CODES, 2):
        c = L(j)
        av, ap = f"{c}{CR['Avant choc']}", f"{c}{CR[nom]}"
        dist = lambda x: f"(MAX({LO(code)}-{x},0)+MAX({x}-{HI(code)},0))"
        cel(ws, k, j, f'=IF(ABS({ap}-{av})<1E-12,"Sans effet",IF({dist(ap)}<{dist(av)},"Améliore",IF({dist(ap)}>{dist(av)},"Dégrade","Sans effet sur l\'écart")))')
com(ws.cell(r + 1, 1), "Distance à la norme = écart à la borne basse si la valeur est dessous, écart à la borne haute si elle est dessus, zéro si conforme. « Améliore » : la distance diminue.")
r = r + 2 + len(CHOCS)
# faisabilité conjointe
ws.cell(r, 1, "Faisabilité conjointe des normes de flux (démonstration)").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Indicateur"] + [str(e) for e in EX])
lignes = [
    ("Autres éléments du résultat A = excédent - cotisations + prestations + charges de fonctionnement",
     lambda e: f"={V('RN', e)}-{V('COT', e)}+{V('PRE', e)}+{V('CF', e)}", FMT_M),
    ("A en % des cotisations", lambda e: f"=({V('RN', e)}-{V('COT', e)}+{V('PRE', e)}+{V('CF', e)})/{V('COT', e)}", FMT_P),
    ("Seuil maximal de A pour respecter ensemble SIN, CUP et EXN (en % des cotisations)",
     lambda e: f"={HI('EXN')}-1+{HI('SIN')}*(1+{HI('CUP')})", FMT_P),
    ("Normes de flux conciliables ?", None, None),
    ("Exercices pour constituer 6 mois de réserve avec l'excédent maximal autorisé (5 %), à SIN constant",
     lambda e: f"={LO('ROB')}/12*{V('PRE', e)}/({HI('EXN')}*{V('COT', e)})", FMT_R),
]
FJ = {}
for k, (lib, fx, fmt) in enumerate(lignes, r + 2):
    FJ[lib[:20]] = k
    cel(ws, k, 1, lib, wrap=True, gras=fx is None)
    ws.row_dimensions[k].height = 30
    for ex in EX:
        c = L(2 + EX.index(ex))
        if fx is None:
            cel(ws, k, 2 + EX.index(ex), f'=IF({c}{k - 3}<={c}{k - 2},"Oui","Non")', gras=True)
        else:
            cel(ws, k, 2 + EX.index(ex), fx(ex), fmt)
FJ_A, FJ_AP, FJ_S, FJ_OK, FJ_ROB = range(r + 2, r + 7)
com(ws.cell(FJ_S, 1), "Démonstration : excédent RN = COT - PRE - CF + A. EXN ≤ 5 % impose A ≤ 5 % COT - COT + PRE + CF. Or PRE ≤ 80 % COT (SIN) et CF ≤ 15 % PRE ≤ 12 % COT (CUP). "
    "Donc A ≤ (5 % - 100 % + 80 % × 115 %) COT = -3,00 % des cotisations. Si A dépasse ce seuil, aucune combinaison des leviers L1 à L3 ne rend conformes ensemble SIN, CUP et EXN.", 520, 200)
com(ws.cell(FJ_ROB, 1), "Réserve visée = 6/12 × prestations ; excédent annuel maximal = 5 % × cotisations ; nombre d'exercices = rapport des deux (sans placement de la réserve).")
largeurs(ws, [58] + [13] * len(CODES))
for j in range(2, 2 + len(EX)):
    ws.column_dimensions[L(j)].width = 15

# =============================================================================================== 13g inflation médicale
ser = {}
for f in sorted(glob.glob(os.path.join(IHPC, "*", "*.csv")) + glob.glob(os.path.join(IHPC, "*.csv"))):
    nom = os.path.basename(f).split("_ ")[1].rsplit(" ", 1)[0].replace("#U00e9", "é")
    d = pd.read_csv(f, dtype=str)
    d["pos"] = d.groupby("annee_deb_couv").cumcount() + 1
    m = d["mois_deb_couv"].astype(int)
    assert ((m == 0) | (m == d["pos"])).all() and set(d.loc[m == 0, "pos"]) <= {2, 8, 12}, nom
    ser[nom] = {(int(a), int(p)): float(v) for a, p, v in zip(d.annee_deb_couv, d.pos, d.valeur)}
POSTES = ["Services hospitaliers", "Services de laboratoires et de radiologie", "Médicaments modernes", "Services médicaux et dentaires",
          "Services des auxiliaires médicaux", "Appareils et matériel thérapeutiques", "Produits medicaux divers", "Médicaments traditionnels"]
assert set(POSTES) == set(ser), sorted(ser)
mois = sorted({k for s in ser.values() for k in s})
ws = wb.create_sheet("13g_inflation")
titre(ws, "Inflation médicale : sous-indices IHPC santé de l'ANStat pondérés par les prestations du FPPN",
      "Source : ANStat, IHPC par poste de consommation (national), fichiers téléchargés le 19/10/2026, janvier 2017 à octobre 2024. Anomalie d'export : les mois 2, 8 et 12 sont codés 0 ; "
      "le mois est reconstitué par la position de la ligne dans l'année (12 lignes ordonnées par an). Année de base non indiquée dans les fichiers : seules les variations sont utilisées.")
# tableau mensuel en colonnes N et suivantes
c0 = 14
entete_m = ["Année", "Mois"] + POSTES
for j, v in enumerate(entete_m):
    x = ws.cell(4, c0 + j, v)
    x.font, x.fill = Font(bold=True, color="FFFFFF"), H
    x.alignment = Alignment(wrap_text=True, horizontal="center")
    ws.column_dimensions[L(c0 + j)].width = 12
for i, (a, p) in enumerate(mois, 5):
    ws.cell(i, c0, a)
    ws.cell(i, c0 + 1, p)
    for j, po in enumerate(POSTES, 2):
        ws.cell(i, c0 + j, ser[po].get((a, p))).number_format = FMT_R
dm = 4 + len(mois)
plage_a = f"${L(c0)}$5:${L(c0)}${dm}"
plage_m = f"${L(c0 + 1)}$5:${L(c0 + 1)}${dm}"
# moyennes annuelles et variations (colonnes A à L)
ANS_I = list(range(2017, 2025))
entete(ws, 4, ["Poste IHPC", "Retenu"] + [str(a) for a in ANS_I] + ["Oct. 2024 / oct. 2023"])
RET = {"Services hospitaliers": "Oui", "Services de laboratoires et de radiologie": "Oui", "Médicaments modernes": "Oui", "Services médicaux et dentaires": "Oui",
       "Services des auxiliaires médicaux": "Non (rééducation non isolée)", "Appareils et matériel thérapeutiques": "Non (optique exclue, RI art. 45)",
       "Produits medicaux divers": "Non (non isolé)", "Médicaments traditionnels": "Non (hors couverture)"}
MOY = {}
for i, po in enumerate(POSTES, 5):
    MOY[po] = i
    cel(ws, i, 1, po, wrap=True)
    cel(ws, i, 2, RET[po])
    cpo = f"${L(c0 + 2 + POSTES.index(po))}$5:${L(c0 + 2 + POSTES.index(po))}${dm}"
    for j, a in enumerate(ANS_I, 3):
        cel(ws, i, j, f"=AVERAGEIFS({cpo},{plage_a},{a})", FMT_R)
    cel(ws, i, 3 + len(ANS_I), f"=SUMIFS({cpo},{plage_a},2024,{plage_m},10)/SUMIFS({cpo},{plage_a},2023,{plage_m},10)-1", FMT_P)
com(ws.cell(4, 3 + len(ANS_I) - 1), "2024 : moyenne de janvier à octobre seulement (données publiées jusqu'en octobre 2024).")
r = 5 + len(POSTES) + 1
ws.cell(r, 1, "Variation des moyennes annuelles").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Poste IHPC", ""] + [str(a) for a in ANS_I[1:]])
VAR = {}
for k, po in enumerate(POSTES, r + 2):
    VAR[po] = k
    cel(ws, k, 1, po)
    for j, a in enumerate(ANS_I[1:], 3):
        cel(ws, k, j, f"={L(j + 1)}{MOY[po]}/{L(j)}{MOY[po]}-1", FMT_P)
r = r + 2 + len(POSTES) + 1
# pondération par les prestations du FPPN
ws.cell(r, 1, "Pondération par les prestations maladie du FPPN (exercice 2025)").font = Font(bold=True, color=BLEU)
NAT = {x["nature"]: x for x in RS["prestations_maladie_nature"]}
MAP = [("Hospitalisations", "Services hospitaliers"), ("Analyses médicales", "Services de laboratoires et de radiologie"),
       ("Radio et imagerie", "Services de laboratoires et de radiologie"), ("Pharmacie", "Médicaments modernes"),
       ("Consultations", "Services médicaux et dentaires"), ("Autres prestations maladie", None), ("Autres (transport, rééducation, bons)", None)]
entete(ws, r + 1, ["Nature des prestations (FPPN)", "Poste IHPC associé", "Montant 2025", "Poids retenu"])
r1 = r + 2
for k, (nat, po) in enumerate(MAP, r1):
    cel(ws, k, 1, nat)
    cel(ws, k, 2, po or "Non associé (nature non détaillée)")
    cel(ws, k, 3, round(NAT[nat]["2025"]), FMT_M)
rN = r1 + len(MAP) - 1
assoc = [k for k, (n, p) in enumerate(MAP, r1) if p]
den = "+".join(f"C{k}" for k in assoc)
for k, (nat, po) in enumerate(MAP, r1):
    cel(ws, k, 4, f"=C{k}/({den})" if po else 0, FMT_P)
com(ws.cell(r1 - 1, 4), "Poids = montant 2025 de la nature / total des natures associées à un poste IHPC. Les natures « Autres » ne sont pas détaillées dans les données : elles sont exclues de la pondération (hypothèse : même évolution de prix que la moyenne pondérée).")
r = rN + 2
ws.cell(r, 1, "Inflation médicale pondérée du FPPN").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Exercice", "Inflation pondérée", "Source"])
INF = {}
for k, a in enumerate(range(2018, 2026), r + 2):
    INF[a] = k
    cel(ws, k, 1, a)
    if a <= 2023:
        termes = "+".join(f"D{q}*{L(2 + ANS_I.index(a))}{VAR[MAP[q - r1][1]]}" for q in assoc)
        cel(ws, k, 2, f"={termes}", FMT_P)
        cel(ws, k, 3, "ANStat, variation des moyennes annuelles")
    elif a == 2024:
        termes = "+".join(f"D{q}*{L(3 + len(ANS_I))}{MOY[MAP[q - r1][1]]}" for q in assoc)
        cel(ws, k, 2, f"={termes}", FMT_P)
        cel(ws, k, 3, "ANStat, glissement octobre 2024 sur octobre 2023")
    else:
        cel(ws, k, 2, "=AVERAGE(-0.013,-0.024,-0.019)", FMT_P)
        cel(ws, k, 3, "BCEAO, RPM mars 2026, tableau 12 : santé UEMOA, moyenne T2 à T4 2025 (repli)")
r = r + 2 + 8 + 1
cel(ws, r, 1, "Moyenne 2021 à 2024 (ANStat)", gras=True)
cel(ws, r, 2, f"=AVERAGE(B{INF[2021]}:B{INF[2024]})", FMT_P, gras=True)
INF["MOY"] = r
cel(ws, r + 1, 1, "Maximum 2018 à 2024 (ANStat)", gras=True)
cel(ws, r + 1, 2, f"=MAX(B{INF[2018]}:B{INF[2024]})", FMT_P, gras=True)
INF["MAX"] = r + 1
r += 3
# projection de la sinistralité
ws.cell(r, 1, "Sinistralité si les prestations augmentaient du taux i, à cotisations inchangées").font = Font(bold=True, color=BLEU)
entete(ws, r + 1, ["Hausse des prestations i", "SIN 2025", "Statut 2025", "SIN 2026 (8 mois)", "Statut 2026"])
GRILLE = [-0.02, 0.0, 0.02, 0.04, 0.06, 0.08, 0.10, f"=B{INF['MOY']}", f"=B{INF['MAX']}"]
PRJ = {}
for k, g in enumerate(GRILLE, r + 2):
    PRJ[k] = g
    cel(ws, k, 1, g, FMT_P, gras=isinstance(g, str))
    for ex, cv, cs in ((2025, 2, 3), (2026, 4, 5)):
        cel(ws, k, cv, f"={V('PRE', ex)}*(1+A{k})/{V('COT', ex)}", FMT_P)
        cel(ws, k, cs, f'=IF(AND({L(cv)}{k}>={LO("SIN")},{L(cv)}{k}<={HI("SIN")}),"Conforme","Non conforme")')
kk = r + 2 + len(GRILLE)
cel(ws, kk, 1, "Hausse qui porterait SIN à 80 %", gras=True)
cel(ws, kk, 2, f"={HI('SIN')}*{V('COT', 2025)}/{V('PRE', 2025)}-1", FMT_P, gras=True)
cel(ws, kk, 4, f"={HI('SIN')}*{V('COT', 2026)}/{V('PRE', 2026)}-1", FMT_P, gras=True)
cel(ws, kk + 1, 1, "Hausse qui porterait SIN à 70 %", gras=True)
cel(ws, kk + 1, 2, f"={LO('SIN')}*{V('COT', 2025)}/{V('PRE', 2025)}-1", FMT_P, gras=True)
cel(ws, kk + 1, 4, f"={LO('SIN')}*{V('COT', 2026)}/{V('PRE', 2026)}-1", FMT_P, gras=True)
INF["S80"], INF["S70"], INF["GR0"] = kk, kk + 1, r + 2
com(ws.cell(kk, 1), "Seuil = borne × cotisations / prestations - 1 : hausse des prestations, à cotisations inchangées, qui amène la sinistralité sur la borne. Les cotisations sont assises sur la solde de base, non indexée sur les prix de la santé.")
largeurs(ws, [40, 22, 12, 12, 12, 12, 12, 12, 12, 12, 16])
ws.freeze_panes = "A5"

# =============================================================================================== 13i rendement de marché
ws = wb.create_sheet("13i_rendement")
titre(ws, "Rendement des placements : comparaison aux taux du marché (ratios RL6 et RL7)",
      "RL6 : le taux brut de rentabilité des placements doit atteindre au moins le taux moyen du marché (Décision 05/2022, section II). Taux de marché : BCEAO, Rapport sur la Politique Monétaire "
      "dans l'UMOA, mars 2026, tableau 22 (dépôts à terme, Côte d'Ivoire) et tableau 17 (titres publics de l'UMOA, sources UMOA-Titres et AMF-UMOA).")
entete(ws, 4, ["Série", "Unité", "T1 2024", "T2 2024", "T3 2024", "T4 2024", "T1 2025", "T2 2025", "T3 2025", "T4 2025", "Moyenne 2024", "Moyenne 2025", "Source"])
TAUX = [("Dépôts à terme, Côte d'Ivoire", [4.11, 4.24, 4.27, 4.36, 4.49, 4.61, 4.79, 5.01], "BCEAO, RPM mars 2026, tableau 22"),
        ("Obligations à 5 ans (UMOA)", [6.47, 6.99, 6.98, 7.32, 6.70, 7.13, 6.72, 6.78], "BCEAO, RPM mars 2026, tableau 17"),
        ("Obligations à 10 ans (UMOA)", [6.81, 6.47, 7.41, 6.99, 7.47, 6.95, 6.97, 6.87], "BCEAO, RPM mars 2026, tableau 17"),
        ("Rendement moyen pondéré des obligations (UMOA)", [7.31, 7.37, 7.02, 7.46, 7.07, 7.45, 7.11, 7.39], "BCEAO, RPM mars 2026, tableau 17")]
TX = {}
DAT = "Dépôts à terme, Côte d'Ivoire"
for i, (nom, v, src) in enumerate(TAUX, 5):
    TX[nom] = i
    cel(ws, i, 1, nom, wrap=True)
    cel(ws, i, 2, "%")
    for j, x in enumerate(v, 3):
        cel(ws, i, j, x / 100, FMT_P)
    cel(ws, i, 11, f"=AVERAGE(C{i}:F{i})", FMT_P, gras=True)
    cel(ws, i, 12, f"=AVERAGE(G{i}:J{i})", FMT_P, gras=True)
    cel(ws, i, 13, src, wrap=True)
com(ws.cell(5, 11), "Moyenne arithmétique des quatre trimestres. Le tableau 17 publie ses propres moyennes annuelles (5 ans : 6,99 % en 2024 et 6,81 % en 2025 ; 10 ans : 6,70 % et 6,93 %), calculées par la BCEAO sur les montants ; elles peuvent différer légèrement.")
r = 5 + len(TAUX) + 1
entete(ws, r, ["Indicateur du FPPN", "Unité", "2021", "2022", "2023", "2024", "2025", "", "", "", "", "", "Source"])
rl6 = R["RL6"]["valeurs"]
rl7 = R["RL7"]["valeurs"]
cel(ws, r + 1, 1, "RL6 : taux brut de rentabilité des placements")
cel(ws, r + 1, 2, "%")
for j, ex in enumerate(range(2021, 2026), 3):
    v = rl6.get(str(ex), rl6.get(ex))
    cel(ws, r + 1, j, v, FMT_P)
cel(ws, r + 1, 13, "Onglet 10_ratios (produits financiers / placements moyens)", wrap=True)
cel(ws, r + 2, 1, "Statut de RL6 face aux dépôts à terme de Côte d'Ivoire", gras=True)
for j, ex in ((6, 2024), (7, 2025)):
    cel(ws, r + 2, j, f'=IF({L(j)}{r + 1}>={"K" if ex == 2024 else "L"}{TX[DAT]},"Conforme","Non conforme")', gras=True)
cel(ws, r + 2, 13, "Taux de marché disponibles pour 2024 et 2025 seulement", wrap=True)
cel(ws, r + 3, 1, "RL7 : solde moyen des comptes courants non rémunérés")
cel(ws, r + 3, 2, "FCFA")
for j, ex in enumerate(range(2021, 2026), 3):
    v = rl7.get(str(ex), rl7.get(ex))
    cel(ws, r + 3, j, None if v is None else round(v), FMT_M)
cel(ws, r + 4, 1, "Produit annuel si ce solde était placé au taux moyen des dépôts à terme de Côte d'Ivoire", wrap=True)
cel(ws, r + 4, 2, "FCFA")
cel(ws, r + 4, 6, f"=F{r + 3}*K{TX[DAT]}", FMT_M, gras=True)
cel(ws, r + 4, 7, f"=G{r + 3}*L{TX[DAT]}", FMT_M, gras=True)
com(ws.cell(r + 4, 1), "Produit = solde moyen non rémunéré × taux moyen des dépôts à terme de l'année ; un dépôt à terme reste en classe 5 : la liquidité au sens de LIQ n'est pas réduite. Les échéances doivent rester compatibles avec les sorties d'épargne.")
cel(ws, r + 5, 1, "Taux servi sur l'épargne PEP et EFC (déclaration du chef comptable, 19/10/2026)")
cel(ws, r + 5, 2, "%")
cel(ws, r + 5, 3, 0.035, FMT_P)
cel(ws, r + 6, 1, "Écart entre les obligations à 5 ans (moyenne 2025) et le taux servi", wrap=True)
cel(ws, r + 6, 2, "%")
cel(ws, r + 6, 3, f"=L{TX['Obligations à 5 ans (UMOA)']}-C{r + 5}", FMT_P, gras=True)
RDT = {"TX": TX, "RL6": r + 1, "RL6ST": r + 2, "RL7": r + 3, "GAIN": r + 4, "SERVI": r + 5, "ECART": r + 6}
if INVEST:
    import datetime as _dt
    BLOCS = {}
    for duree, col0 in (("5", 16), ("10", 20)):
        f = next(x for x in glob.glob(os.path.join(INVEST, "*.csv")) if f" {duree}-Year" in x)
        d = pd.read_csv(f, dtype=str, encoding="utf-8-sig")
        d.columns = [c.strip() for c in d.columns]
        d["date"] = pd.to_datetime(d["Date"], format="%d/%m/%Y")
        d["v"] = d["Dernier"].str.replace(",", ".").astype(float) / 100
        d = d.sort_values("date")
        for j, t in enumerate(("Année", "Date", f"Rendement {duree} ans")):
            x = ws.cell(4, col0 + j, t)
            x.font, x.fill = Font(bold=True, color="FFFFFF"), H
            x.alignment = Alignment(horizontal="center", wrap_text=True)
            ws.column_dimensions[L(col0 + j)].width = 12
        for i, (dt_, v) in enumerate(zip(d["date"], d["v"]), 5):
            ws.cell(i, col0 + 1, dt_.to_pydatetime()).number_format = "DD/MM/YYYY"
            ws.cell(i, col0, f"=YEAR({L(col0 + 1)}{i})")
            ws.cell(i, col0 + 2, v).number_format = "0.000%"
        BLOCS[duree] = (col0, 4 + len(d), d["date"].max().strftime("%d/%m/%Y"), len(d))
    r2 = RDT["ECART"] + 2
    ws.cell(r2, 1, "Rendement des obligations de l'État de Côte d'Ivoire sur le marché secondaire (Investing.com)").font = Font(bold=True, color=BLEU)
    entete_ = ["Série", "Unité", "2021", "2022", "2023", "2024", "2025", "2026", "Dernière valeur", "Date", "Cotations", "", "Source"]
    for j, t in enumerate(entete_, 1):
        x = ws.cell(r2 + 1, j, t if t else None)
        if t:
            x.font, x.fill = Font(bold=True, color="FFFFFF"), H
            x.alignment = Alignment(horizontal="center", wrap_text=True)
    INV = {}
    for k, duree in enumerate(("5", "10")):
        rr = r2 + 2 + k
        col0, fin, der, n = BLOCS[duree]
        an_, va_ = f"${L(col0)}$5:${L(col0)}${fin}", f"${L(col0 + 2)}$5:${L(col0 + 2)}${fin}"
        cel(ws, rr, 1, f"Obligations de l'État de Côte d'Ivoire à {duree} ans (moyenne annuelle des cotations)", wrap=True)
        cel(ws, rr, 2, "%")
        for j, a in enumerate(range(2021, 2027), 3):
            cel(ws, rr, j, f'=IFERROR(AVERAGEIFS({va_},{an_},{a}),"")', FMT_P)
        cel(ws, rr, 9, f"={L(col0 + 2)}{fin}", FMT_P, gras=True)
        cel(ws, rr, 10, der)
        cel(ws, rr, 11, n, "0")
        cel(ws, rr, 13, f"Investing.com, « Rendement de l'Obligation Cote d'Ivoire {duree}-Year - Données Historiques », export du 19/10/2026 (colonnes {L(col0)} à {L(col0 + 2)})", wrap=True)
        INV[duree] = rr
    com(ws.cell(r2 + 2, 1), "Moyenne arithmétique des cotations de clôture de l'année (colonne « Dernier » du fichier). Investing.com est un agrégateur de cotations du marché secondaire : "
        "source secondaire, utilisée en complément des taux publiés par la BCEAO (source primaire). 2021 : à partir du 25/03/2021 ; 2026 : jusqu'à la dernière cotation. "
        "Plusieurs cotations identiques d'un jour à l'autre traduisent l'absence de transaction.", 520, 180)
    rr = r2 + 4
    cel(ws, rr, 1, "RL6 : taux brut de rentabilité des placements (rappel)")
    cel(ws, rr, 2, "%")
    for j, ex in enumerate(range(2021, 2026), 3):
        cel(ws, rr, j, f"={L(j)}{RDT['RL6']}", FMT_P)
    cel(ws, rr + 1, 1, "Statut de RL6 face aux obligations de l'État à 5 ans", gras=True)
    for j, ex in enumerate(range(2021, 2026), 3):
        cel(ws, rr + 1, j, f'=IF({L(j)}{INV["5"]}="","",IF({L(j)}{rr}>={L(j)}{INV["5"]},"Conforme","Non conforme"))', gras=True)
    cel(ws, rr + 2, 1, "Écart entre les obligations de l'État à 5 ans (dernière valeur) et le taux servi de 3,5 %", wrap=True)
    cel(ws, rr + 2, 2, "%")
    cel(ws, rr + 2, 3, f"=I{INV['5']}-C{RDT['SERVI']}", FMT_P, gras=True)
    com(ws.cell(rr + 1, 1), "Lecture complémentaire de RL6 : la Décision n° 05/2022 compare le rendement au « taux moyen du marché » sans préciser le marché ; les dépôts à terme (BCEAO) restent la lecture principale, les obligations de l'État une lecture exigeante.")
    RDT.update({"INV5": INV["5"], "INV10": INV["10"], "RL6B": rr, "RL6INV": rr + 1, "ECARTINV": rr + 2})
largeurs(ws, [52, 7] + [10] * 8 + [12, 12, 40])
for j_ in (16, 17, 18, 20, 21, 22):
    ws.column_dimensions[L(j_)].width = 12

wb.move_sheet("13h_synthese", offset=2)
wb.save(SORTIE)
CARTE = {"VR": VR, "NR": NR, "D": D, "EQR": {f"{a}|{b}": v for (a, b), v in EQR.items()}, "FR": FR, "CR": CR, "SENS": SENS,
         "FJ": [FJ_A, FJ_AP, FJ_S, FJ_OK, FJ_ROB], "INF": INF, "MOY": MOY, "VAR": VAR, "COL": COL, "CODES": CODES,
         "MAP_r1": r1, "MAP": MAP, "P": P, "RDT": RDT, "PREUVES": [p[:2] for p in PREUVES], "LEV": LEV, "FACT": FACT, "REFS": REFS, "NORMES": NORMES}
json.dump(CARTE, open(os.path.splitext(SORTIE)[0] + "_carte.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
print("onglets :", wb.sheetnames)
