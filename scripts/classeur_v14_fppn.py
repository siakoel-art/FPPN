# -*- coding: utf-8 -*-
"""
FPPN - Classeur des calculs, compléments de la version 14 du rapport (documents de la dataroom du 09/11/2026)
=============================================================================================================
Ajoute au classeur existant, avec des formules, cinq onglets écrits en bleu (ajouts du 10/11/2026) :
  30_pidr_fppn          calcul du FPPN pour la provision retraite (2020 à 2025), compte 196 et recalcul Adenka ;
  31_ratios_fppn        ratios prudentiels calculés par le FPPN comparés aux ratios recalculés ;
  32_actifs_contestes   actifs dont la valeur n'est pas établie et sensibilité des fonds propres ;
  33_grenier_ef         comptes propres du Grenier du Policier, 2024 et 2025 ;
  34_documents_0911     documents reçus le 09/11/2026 et usage dans le rapport ;
puis met à jour la traçabilité (numéros de diapositive du rapport livré avec ce classeur).
Utilisation : python classeur_v14_fppn.py <classeur.xlsx> <index_rapport.json> <v9.json> <ratios.json> <sortie.xlsx>
"""
import json
import os
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

SRC, IDXJ, V9J, RJ, SORTIE = sys.argv[1:6]
IDX = json.load(open(IDXJ, encoding="utf-8"))
V9 = json.load(open(V9J, encoding="utf-8"))
RT = {x["code"]: x for x in json.load(open(RJ, encoding="utf-8"))["ratios"]}
wb = openpyxl.load_workbook(SRC)
MODELE = wb["27_pidr_plafond"]
BLEU = "0070C0"
FB = Font(name="Helvetica", size=9, color=BLEU)
FBG = Font(name="Helvetica", size=9, color=BLEU, bold=True)
FMT, PCT = "# ##0", "0.00 %"


def barre(ws, sources):
    for j in range(1, 5):
        c, m = ws.cell(1, j), MODELE.cell(1, j)
        c.value, c._style = m.value, copy(m._style)
        if m.hyperlink is not None:
            c.hyperlink = copy(m.hyperlink)
    ws.cell(1, 5).value, ws.cell(1, 5)._style = "Chiffres tirés de :", copy(MODELE.cell(1, 5)._style)
    for k, src in enumerate(sources, 6):
        c = ws.cell(1, k, src)
        c._style = copy(MODELE.cell(1, 6)._style)
        if src in wb.sheetnames:
            c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{src}'!A1")
    ws.freeze_panes = "A3"


def explications(ws, titre, rubriques):
    c = ws.cell(2, 1, titre)
    c._style = copy(MODELE.cell(2, 1)._style)
    c.font = Font(name="Helvetica", size=11, bold=True, color=BLEU)
    r = 3
    for lab, txt in rubriques:
        a, b_ = ws.cell(r, 1, lab), ws.cell(r, 2, txt)
        a._style, b_._style = copy(MODELE.cell(3, 1)._style), copy(MODELE.cell(3, 2)._style)
        b_.font = FB
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        r += 1
    return r + 1


def entete(ws, r, cols, largeurs):
    for j, (h, w) in enumerate(zip(cols, largeurs), 1):
        c = ws.cell(r, j, h)
        c._style = copy(MODELE.cell(13, 1)._style)
        ws.column_dimensions[get_column_letter(j)].width = w


def val(ws, r, j, v, fmt=None, gras=False):
    c = ws.cell(r, j, v)
    c.font = FBG if gras else FB
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if fmt:
        c.number_format = fmt
    return c


# ------------------------------------------------------------------ 30 : provision retraite selon le FPPN
ws = wb.create_sheet("30_pidr_fppn")
barre(ws, ["14c_pidr_synthese", "12_pidr_196", "D1_balances"])
r = explications(ws, "Tableau : provision du personnel calculée par le FPPN, compte 196 et recalcul Adenka", [
    ("À quoi sert cet onglet", "Comparer le calcul du FPPN (fichiers reçus le 09/11/2026), le compte 196 et le recalcul agent par agent du rapport."),
    ("D'où viennent les chiffres", "Fichiers « FICHIER PERSONNEL CIVIL FPPN POUR PROVISION CONGES ET RETRAITE EN 2020 » à « … EN 2025 », onglet RETRAITE (total et nombre d'agents) ; compte 19600000 des balances ; onglet 14c (droits acquis et provision recalculée)."),
    ("Comment le calcul est fait", "Écart = compte 196 - calcul du FPPN. Écart relatif des droits acquis = (calcul du FPPN - droits acquis Adenka) / calcul du FPPN."),
    ("Exemple chiffré", "2025 : (43 666 678 - 43 195 084) / 43 666 678 = 1,08 %."),
    ("Valable pour", "Personnel civil du FPPN, exercices 2020 à 2025 ; recalcul Adenka sur les DISA 2021, 2022, 2023 et 2025."),
    ("Limites", "Le FPPN calcule en droits acquis, sans actualisation, ni mortalité, ni départ anticipé. Pas de recalcul Adenka en 2020 et 2024 (DISA non exploitables)."),
])
entete(ws, r, ["Exercice", "Agents (FPPN)", "Calcul du FPPN", "Compte 196", "Écart", "Agents (Adenka)", "Droits acquis (Adenka)", "Provision recalculée (Adenka)", "Écart relatif droits acquis"], [10, 12, 18, 18, 14, 14, 20, 22, 16])
PF = {2020: (45, 24155943, 24155943), 2021: (41, 22342542, 22542542), 2022: (64, 28325794, 28325794), 2023: (62, 35312179, 35312179),
      2024: (61, 34753547, 34753547), 2025: (78, 43666678, 43666678)}
PSY = {x["annee"]: x for x in V9["pidr_synthese"]}
for i, (a, (ag, tot, c196)) in enumerate(PF.items(), r + 1):
    val(ws, i, 1, a)
    val(ws, i, 2, ag, FMT)
    val(ws, i, 3, tot, FMT)
    val(ws, i, 4, c196, FMT)
    val(ws, i, 5, f"=D{i}-C{i}", FMT)
    if a in PSY:
        val(ws, i, 6, PSY[a]["agents"], FMT)
        val(ws, i, 7, round(PSY[a]["acquis"]), FMT)
        val(ws, i, 8, round(PSY[a]["dbo"]), FMT)
        val(ws, i, 9, f"=(C{i}-G{i})/C{i}", PCT)
    else:
        for j in (6, 7, 8, 9):
            val(ws, i, j, "n.d.")

# ------------------------------------------------------------------ 31 : ratios du FPPN et d'Adenka
ws = wb.create_sheet("31_ratios_fppn")
barre(ws, ["10_ratios", "10_ratios_detail"])
r = explications(ws, "Tableau : ratios prudentiels calculés par le FPPN et recalculés par Adenka", [
    ("À quoi sert cet onglet", "Réconcilier les ratios que le FPPN présente avec ceux du tableau de bord du rapport."),
    ("D'où viennent les chiffres", "Document du FPPN « INDICATEURS ET RATIOS PRUDENTIELS UEMOA. ACTUALISES 2 AA.docx » (2024, 2025 et 1er semestre 2026), reçu le 09/11/2026 ; onglet 10_ratios (valeurs Adenka, 2026 au 31/08/2026)."),
    ("Comment le calcul est fait", "Écart = valeur Adenka - valeur du FPPN, même unité. Liquidité et équilibre stable : lecture corrigée selon le PCMS, car le FPPN retraite aussi l'épargne."),
    ("Exemple chiffré", "Sinistralité 2025 : 57,48 % des deux côtés, écart nul. Réserve obligatoire 2025 : 5,01 mois (Adenka) contre 9 mois (FPPN)."),
    ("Valable pour", "Exercices 2024 et 2025 ; 2026 : 1er semestre pour le FPPN, 31/08/2026 pour Adenka (périodes différentes)."),
    ("Limites", "Le FPPN ne détaille ni sa base de calcul de la réserve obligatoire ni son retraitement des ressources stables."),
])
entete(ws, r, ["Ratio", "Norme", "FPPN 2024", "Adenka 2024", "Écart 2024", "FPPN 2025", "Adenka 2025", "Écart 2025", "FPPN 2026 (S1)", "Adenka 31/08/2026"], [34, 14, 12, 12, 12, 12, 12, 12, 14, 14])
RF = [("SIN", "Taux de sinistralité", "70 % à 80 %", (0.4034, 0.5748, 0.8287), True, False),
      ("CHF", "Charges de fonctionnement / cotisations", "Au plus 20 %", (0.3606, 0.3633, 0.4507), True, False),
      ("CHB", "Charges de fonctionnement brut / produits", "Au plus 15 %", (0.3220, 0.3486, 0.4487), True, False),
      ("CHP", "Charges de personnel / charges de fonctionnement", "Au plus 85 %", (0.248, 0.2657, 0.2391), True, False),
      ("AUT", "Autonomie financière", "Au moins 1", (1.29, 1.051, 0.78), False, False),
      ("ROB", "Réserve obligatoire (mois)", "6 à 9 mois", (5, 9, None), False, False),
      ("LIQ", "Liquidité (lecture corrigée)", "Au moins 1", (0.31, 1.006, 1.16), False, True),
      ("FEP", "Fonds propres et assimilés / dettes", "Au moins 1", (0.40, 0.31, 0.20), False, False),
      ("EQS", "Équilibre stable (lecture corrigée)", "Au plus 1", (0.48, 0.38, 0.36), False, True),
      ("EXN", "Taux d'excédent net", "Au plus 5 %", (0.2518, 0.1281, 0.2732), True, False),
      ("TFP", "Financement propre", "Au moins 100 %", (1.3248, 1.1347, 0.7865), True, False)]
for i, (c_, lib, nor, fp, pct, corr) in enumerate(RF, r + 1):
    vals = (RT[c_].get("valeurs_retraitees") or {}) if corr else RT[c_]["valeurs"]
    ad = [vals.get(str(a), vals.get(a)) for a in (2024, 2025, 2026)]
    f_ = PCT if pct else "0.00"
    val(ws, i, 1, lib)
    val(ws, i, 2, nor)
    val(ws, i, 3, fp[0], f_)
    val(ws, i, 4, ad[0], f_)
    val(ws, i, 5, f"=D{i}-C{i}", f_)
    val(ws, i, 6, fp[1], f_)
    val(ws, i, 7, ad[1], f_)
    val(ws, i, 8, f"=G{i}-F{i}", f_)
    val(ws, i, 9, fp[2] if fp[2] is not None else "n.d.", f_)
    val(ws, i, 10, ad[2], f_)

# ------------------------------------------------------------------ 32 : actifs dont la valeur n'est pas établie
ws = wb.create_sheet("32_actifs_contestes")
barre(ws, ["4_bilan", "10_actif_mobilisation", "D1_balances"])
r = explications(ws, "Tableau : actifs dont la valeur n'est pas établie et sensibilité des fonds propres", [
    ("À quoi sert cet onglet", "Mesurer l'effet sur les fonds propres d'une sortie de l'actif du terrain d'Attoban et de la créance sur l'Hôpital de la Police (sensibilité, décision du CA attendue)."),
    ("D'où viennent les chiffres", "Rapports du commissaire aux comptes 2022 à 2025 ; PV du CA du 06/05/2025 ; protocole SONEB du 16/02/2026 ; courrier de l'AIRMS du 05/01/2026 ; compte 47111100 des balances ; fonds propres de l'onglet 4_bilan."),
    ("Comment le calcul est fait", "Montant sorti = terrain d'Attoban + créance sur l'Hôpital de la Police. Fonds propres après sortie = fonds propres - montant sorti."),
    ("Exemple chiffré", "375 000 000 + 3 407 519 850 = 3 782 519 850 FCFA."),
    ("Valable pour", "Bilan au 31/12/2025 et au 31/08/2026 (balance provisoire)."),
    ("Limites", "Le terrain SONEB et la créance sur l'État ne sont pas sortis : leur traitement dépend de documents non reçus."),
])
entete(ws, r, ["Actif", "Montant (FCFA)", "Sorti dans la sensibilité ?", "Situation"], [46, 18, 16, 70])
LIG = [("Terrain « Terre rouge », Cocody Riviera Attoban (4 ha)", 375000000, "Oui", "Sans titre foncier ; vendeur en liquidation ; le commissaire aux comptes demande de le sortir du patrimoine"),
       ("Créance sur l'Hôpital de la Police nationale (47111100)", 3407519850, "Oui", "Inchangée depuis 2020 ; dépenses pouvant être des subventions selon la direction générale"),
       ("Terrain SONEB, Songon (10 ha)", 450000000, "Non", "Sans titre foncier ; remplacé en 2026 par une parcelle de 9 ha à Jacqueville"),
       ("Créance sur l'État (policiers blessés en service commandé)", 3932124730, "Non", "Recouvrement en cours avec l'appui de l'AIRMS")]
for i, (a, m, s_, t) in enumerate(LIG, r + 1):
    val(ws, i, 1, a)
    val(ws, i, 2, m, FMT)
    val(ws, i, 3, s_)
    val(ws, i, 4, t)
i0, i1 = r + 1, r + len(LIG)
k = i1 + 2
val(ws, k, 1, "Montant sorti dans la sensibilité", gras=True)
val(ws, k, 2, f'=SUMIF(C{i0}:C{i1},"Oui",B{i0}:B{i1})', FMT, True)
FP = {}
for x in V9.get("bilan_resume", []):
    FP[x.get("cloture")] = x.get("fonds_propres")
val(ws, k + 2, 1, "Fonds propres au 31/12/2025 (onglet 4_bilan)")
val(ws, k + 2, 2, -438242200, FMT)
val(ws, k + 3, 1, "Fonds propres au 31/12/2025 après sortie", gras=True)
val(ws, k + 3, 2, f"=B{k + 2}-B{k}", FMT, True)
val(ws, k + 4, 1, "Fonds propres au 31/08/2026 (onglet 4_bilan)")
val(ws, k + 4, 2, -879498700, FMT)
val(ws, k + 5, 1, "Fonds propres au 31/08/2026 après sortie", gras=True)
val(ws, k + 5, 2, f"=B{k + 4}-B{k}", FMT, True)

# ------------------------------------------------------------------ 33 : comptes du Grenier
ws = wb.create_sheet("33_grenier_ef")
barre(ws, ["5_grenier_encours", "23_grenier_depreciation"])
r = explications(ws, "Tableau : comptes propres du Grenier du Policier, 2024 et 2025", [
    ("À quoi sert cet onglet", "Lire les comptes du Grenier établis selon le référentiel des SFD et les rapprocher des données du rapport."),
    ("D'où viennent les chiffres", "« ETAT FINANCIER AU 31 DECEMBRE 2025 » du Grenier du Policier (rapport financier, juin 2026), annexes 1 et 2, reçu le 09/11/2026."),
    ("Comment le calcul est fait", "Résultat = produit financier net - charges d'exploitation - coût du risque + résultat exceptionnel. Variation = 2025 / 2024 - 1."),
    ("Exemple chiffré", "2025 : 163 822 737 - 133 288 861 - 0 + 9 500 = 30 543 376 FCFA."),
    ("Valable pour", "Exercices 2024 et 2025 du Grenier du Policier."),
    ("Limites", "Aucune provision pour créances douteuses : le résultat est avant dépréciation des prêts en retard (onglet 23_grenier_depreciation)."),
])
entete(ws, r, ["Poste", "2024", "2025", "Variation"], [46, 18, 18, 14])
G = [("Encours de crédit", 566643134, 829623319), ("Produits d'intérêts sur crédits", 100438256, 128134067), ("Commissions perçues", 25986419, 35688670),
     ("Produit financier net", None, None), ("Frais de personnel", 27780228, 30066030), ("Frais généraux", 88321232, 103222831),
     ("Charges d'exploitation", None, None), ("Coût net du risque (provisions)", 0, 0), ("Résultat exceptionnel", 0, 9500), ("Résultat net", None, None),
     ("Épargne ordinaire collectée", 101524680, 182245130), ("Épargne PEP et EFC inscrite au passif", 8493055501, 8493055501), ("Compte de liaison avec le FPPN", 204060848, 817121335)]
lig = {}
for i, (p, a, b_) in enumerate(G, r + 1):
    lig[p] = i
    val(ws, i, 1, p, gras=a is None)
    for j, v in ((2, a), (3, b_)):
        col = "B" if j == 2 else "C"
        if p == "Produit financier net":
            v = f"={col}{lig['Produits d’intérêts sur crédits'.replace('’', chr(39))]}+{col}{lig['Commissions perçues']}"
        elif p == "Charges d'exploitation":
            v = f"={col}{lig['Frais de personnel']}+{col}{lig['Frais généraux']}"
        elif p == "Résultat net":
            v = f"={col}{lig['Produit financier net']}-{col}{lig[chr(67) + 'harges d' + chr(39) + 'exploitation']}-{col}{lig['Coût net du risque (provisions)']}+{col}{lig['Résultat exceptionnel']}"
        val(ws, i, j, v, FMT, gras=a is None)
    val(ws, i, 4, f'=IF(B{i}<>0,C{i}/B{i}-1,"")', PCT)

# ------------------------------------------------------------------ 34 : documents reçus le 09/11/2026
ws = wb.create_sheet("34_documents_0911")
barre(ws, ["30_pidr_fppn", "31_ratios_fppn", "32_actifs_contestes", "33_grenier_ef"])
r = explications(ws, "Tableau : documents de la dataroom reçus le 09/11/2026 et usage dans le rapport", [
    ("À quoi sert cet onglet", "Relier chaque famille de documents reçus le 09/11/2026 à son usage dans le rapport et dans ce classeur."),
    ("D'où viennent les chiffres", "Inventaire « FPPN_Inventaire_dataroom_20261109.xlsx » : 5 366 fichiers comparés par empreinte SHA-256 ; 980 fichiers transmis et contrôlés."),
    ("Comment le calcul est fait", "Sans objet (inventaire)."),
    ("Exemple chiffré", "Balances 2020 à 2025 : 0 FCFA d'écart compte par compte avec les balances déjà utilisées."),
    ("Valable pour", "Dossiers « 03-Adenka-Audit » et « 05-Adenka Audit Juridique et Fiscal » téléchargés le 09/11/2026."),
    ("Limites", "Le lot C (relevés bancaires, budgets, CNPS, paie) n'a pas été transmis : il n'est pas nécessaire au volet actuariel."),
])
entete(ws, r, ["Famille", "Documents", "Usage", "Onglet"], [20, 60, 60, 20])
DOCS = [("Comptes", "Balances et grands livres 2020 à 2026 ; balance générale 2026 refaite", "Soldes identiques ; balance 2026 : quatre reclassements, dont 816 670 FCFA du 41191000 au 41591000", "D1_balances"),
        ("Provisions", "Provision pour congés et retraite 2020 à 2025", "Méthode du FPPN pour la provision du personnel", "30_pidr_fppn"),
        ("Prestations", "Offres et barèmes des allocations décès ; cartographie des risques", "Barème décès complet ; frais funéraires de 300 000 FCFA depuis 2024", "19a_deces_etats_fin_annee"),
        ("Ratios", "Indicateurs et ratios prudentiels du FPPN 2022 à 2026", "Comparaison des ratios", "31_ratios_fppn"),
        ("Gouvernance et contrôle", "PV des AG et des CA 2020 à 2026 ; rapports du commissaire aux comptes 2022 à 2025 ; courriers de l'AIRMS", "Résultats confirmés ; actifs dont la valeur n'est pas établie ; procès EFC ; projets", "32_actifs_contestes"),
        ("Grenier", "États financiers 2022 à 2025, balances, grands livres, production", "Comptes propres du Grenier", "33_grenier_ef"),
        ("Cotisations", "Fichiers aller-retour CGRAE et GNA modifiés ; tableaux de centralisation", "Lignes ajoutées dans douze fichiers ; écarts aller-retour à rapprocher", "1_aller_retour")]
for i, x in enumerate(DOCS, r + 1):
    for j, v in enumerate(x, 1):
        c = val(ws, i, j, v)
        if j == 4 and v in wb.sheetnames:
            c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{v}'!A1")


# ------------------------------------------------------------------ 35 : sinistralité lue dans les deux sens (v15, 11/11/2026)
import csv as _csv  # noqa: E402
_AS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
SQ = {int(k): v for k, v in json.load(open(os.path.join(_AS, "sinistralite_v15.json"), encoding="utf-8")).items()}
ws = wb.create_sheet("35_sinistralite")
barre(ws, ["D1_balances", "31_ratios_fppn"])
r = explications(ws, "Tableau : prestations maladie par nature réelle et sinistralité des soins seuls", [
    ("À quoi sert cet onglet", "Vérifier ce que contient le compte des prestations maladie et recalculer la sinistralité sans les dépenses étrangères aux soins."),
    ("D'où viennent les chiffres", "Grands livres des comptes 2021 à 2025 et du 01/01 au 31/08/2026 (comptes 601, 6018 et 60800007, hors à-nouveaux) ; cotisations, frais affectés et ratio réglementaire repris des balances (onglets D1 et 31)."),
    ("Comment le calcul est fait", "Chaque écriture est classée par son libellé (DON, K'DO, CADEAU, PANIER : dons ; GADGET, IMPRIMERIE, CARTE : gadgets et imprimés ; factures n° 00444… : honoraires d'un cabinet ; CMP : clinique. Les dons à caractère médical (dialyse, soins, médicaments, kit d'accouchement) restent dans les soins). Soins seuls = total - dépenses étrangères aux soins. Taux = montant / cotisations maladie. Ratio combiné = (prestations + frais affectés) / cotisations."),
    ("Exemple chiffré", "2025 : soins seuls = 1 771 635 361 - 487 492 003 = 1 284 143 358 FCFA ; 1 284 143 358 / 2 409 281 030 = 53,30 %."),
    ("Valable pour", "Exercices 2021 à 2025 et du 01/01 au 31/08/2026 (balance provisoire)."),
    ("Limites", "Classement par libellé, sans pièce justificative. La nature des honoraires du cabinet reste à confirmer par le FPPN."),
])
_A = list(range(2021, 2027))
entete(ws, r, ["Ligne"] + [str(a) for a in _A[:-1]] + ["31/08/2026"], [52, 18, 18, 18, 18, 18, 18])
_CATS = ["Clinique du FPPN (CMP)", "Autres prestataires de soins", "Dons (autorités, élus, délégués)", "Gadgets, imprimés et cartes", "Honoraires d'un cabinet (factures trimestrielles)", "Régularisations comptables"]
i = r + 1
for c_ in _CATS:
    val(ws, i, 1, c_)
    for j, a in enumerate(_A, 2):
        val(ws, i, j, round(SQ[a]["cats"].get(c_, 0.0)), FMT)
    i += 1
r1, r2 = r + 1, i - 1
L = {}
for lab, f in [("Total des prestations maladie", lambda c: f"=SUM({c}{r1}:{c}{r2})"),
               ("Dont dépenses étrangères aux soins", lambda c: f"=SUM({c}{r1 + 2}:{c}{r2})"),
               ("Soins seuls", lambda c: f"={c}{i - 2}-{c}{i - 1}")]:
    val(ws, i, 1, lab, gras=True)
    for j in range(2, 8):
        val(ws, i, j, f(get_column_letter(j)), FMT, gras=True)
    L[lab] = i
    i += 1
for lab, cle in [("Cotisations maladie", "cot"), ("Frais affectés à la maladie", "frais"), ("Numérateur du taux réglementaire (60 + 6901 - 7901)", None), ("Dénominateur du taux réglementaire (70 + 7906 - 6906)", None)]:
    val(ws, i, 1, lab)
    for j, a in enumerate(_A, 2):
        if cle:
            val(ws, i, j, round(SQ[a][cle]), FMT)
        else:
            val(ws, i, j, None, FMT)
    L[lab] = i
    i += 1
for j, a in enumerate(_A, 2):
    ws.cell(L["Numérateur du taux réglementaire (60 + 6901 - 7901)"], j).value = round(SQ[a]["sin_num"])
    ws.cell(L["Dénominateur du taux réglementaire (70 + 7906 - 6906)"], j).value = round(SQ[a]["sin_den"])
i += 1
for lab, f in [("Maladie : taux comptable", lambda c: f"={c}{L['Total des prestations maladie']}/{c}{L['Cotisations maladie']}"),
               ("Maladie : taux des soins seuls", lambda c: f"={c}{L['Soins seuls']}/{c}{L['Cotisations maladie']}"),
               ("Part de la clinique du FPPN dans les soins", lambda c: f"={c}{r1}/{c}{L['Soins seuls']}"),
               ("Maladie : ratio combiné comptable", lambda c: f"=({c}{L['Total des prestations maladie']}+{c}{L['Frais affectés à la maladie']})/{c}{L['Cotisations maladie']}"),
               ("Maladie : ratio combiné des soins seuls", lambda c: f"=({c}{L['Soins seuls']}+{c}{L['Frais affectés à la maladie']})/{c}{L['Cotisations maladie']}"),
               ("Toutes garanties : taux réglementaire comptable", lambda c: f"={c}{L['Numérateur du taux réglementaire (60 + 6901 - 7901)']}/{c}{L['Dénominateur du taux réglementaire (70 + 7906 - 6906)']}"),
               ("Toutes garanties : taux réglementaire des soins seuls", lambda c: f"=({c}{L['Numérateur du taux réglementaire (60 + 6901 - 7901)']}-{c}{L['Dont dépenses étrangères aux soins']})/{c}{L['Dénominateur du taux réglementaire (70 + 7906 - 6906)']}")]:
    val(ws, i, 1, lab, gras=True)
    for j in range(2, 8):
        val(ws, i, j, f(get_column_letter(j)), PCT, gras=True)
    i += 1
i += 2
val(ws, i, 1, "Détail des écritures étrangères aux soins (grands livres, une ligne par écriture)", gras=True)
i += 1
entete(ws, i, ["Libellé", "Exercice", "Date", "Journal", "Pièce", "Compte", "Montant net (FCFA)", "Nature"], [52, 18, 18, 18, 18, 18, 18, 30])
with open(os.path.join(_AS, "prestations_hors_soins_v15.csv"), encoding="utf-8-sig") as f_:
    for x in _csv.DictReader(f_, delimiter=";"):
        i += 1
        for j, v_ in enumerate([x["lib"], int(x["an"]), x["date"], x["journal"], x["piece"], x["compte"], round(float(x["m"])), x["cat"]], 1):
            val(ws, i, j, v_, FMT if j == 7 else None)


# ------------------------------------------------------------------ 36 : documents du 11/11/2026 (v16)
ws = wb.create_sheet("36_documents_1111")
barre(ws, ["D1_balances", "32_actifs_contestes", "33_grenier_ef", "30_pidr_fppn"])
r = explications(ws, "Tableau : faits tirés des documents reçus le 11/11/2026", [
    ("À quoi sert cet onglet", "Reprendre les montants des documents du 11/11/2026 et les rapprocher des comptes du FPPN."),
    ("D'où viennent les chiffres", "Mémo sur les comptes 47 et 27 ; rapport d'expertise du siège n° 005/AKKA/19 du 19/03/2019 ; balances et grands livres du Grenier ; déclarations mensuelles des salaires (EDI ITS) ; balances du FPPN."),
    ("Comment le calcul est fait", "Écart = montant FPPN - montant Grenier ou mémo. Bâtiment net = brut - amortissements. Salariés sans numéro CNPS : numéro 111111111111."),
    ("Exemple chiffré", "Compte de liaison au 31/08/2026 : 915 477 260 - 845 285 760 = 70 191 500 FCFA."),
    ("Valable pour", "Balances au 31/08/2026 (provisoire) ; déclarations de salaires de septembre 2024 à septembre 2026."),
    ("Limites", "Le mémo ne donne pas les pièces. L'écriture d'origine du compte 10620000 n'a pas été reçue."),
])
entete(ws, r, ["Bloc", "Élément", "Montant FPPN ou document", "Montant de rapprochement", "Écart", "Source"], [22, 48, 22, 22, 18, 44])
i = r + 1
LIG = [("Comptes 47 et 27", "47111104 Investissements au CMP", 5803025455, None, "Mémo et balance au 31/08/2026"),
       ("Comptes 47 et 27", "47111100 Hôpital de la Police", 3407519850, None, "Mémo et balance au 31/08/2026"),
       ("Comptes 47 et 27", "47111105 Charges de fonctionnement au CMP", 1280300860, None, "Mémo et balance au 31/08/2026"),
       ("Comptes 47 et 27", "47111102 Grenier (FPPN) / 37921400 (Grenier)", 915477260, 845285760, "Balance FPPN ; balance Grenier 2026"),
       ("Comptes 47 et 27", "27110000 + 27120000 Prêts PEP 2004-2015", 379204261, None, "Mémo et balance au 31/08/2026"),
       ("Comptes 47 et 27", "47111103 Ticket modérateur", 8357475, None, "Mémo et balance au 31/08/2026"),
       ("Comptes 47 et 27", "27200000 Prêts au personnel (balance / mémo)", 2500000, 2400000, "Balance ; mémo"),
       ("Siège", "Compte 10620000 / valeur vénale de l'expertise", 2899400000, 2899400000, "Balances 2020-2026 ; rapport du 19/03/2019"),
       ("Siège", "Bâtiments administratifs 23130000 (brut / amortissements)", 424341213, 273868454, "Balance au 31/08/2026 (écart = valeur nette)"),
       ("Grenier", "Épargne PEP + EFC : FPPN 31/12/2025 / Grenier", 13814066964, 8493055501, "Balances FPPN 2025 ; balance Grenier 2025"),
       ("Grenier", "Crédits aux membres au 31/08/2026 (encours / dépréciation)", 1044302107, 0, "Balance Grenier 2026"),
       ("Grenier", "Fonds de dotation reçu le 18/05/2022, reclassé en dette au 31/12/2025", 500000000, None, "Grand livre 56000000 et 37921400 du Grenier"),
       ("Grenier", "Résultat 2025 (produits - charges)", 163832237, 133288861, "Balance Grenier 2025"),
       ("Placements", "Dépôt à terme BOA PEP fermé le 30/10/2025 / FCP ouvert le 04/11/2025", 505383858, 500000000, "État des comptes bancaires ; balances"),
       ("Personnel", "Salariés août 2026 / sans numéro CNPS", 100, 33, "EDI ITS août 2026"),
       ("Personnel", "Brut d'août 2026 : total / sans numéro CNPS", 28171544, 6340264, "EDI ITS août 2026")]
for bloc_, el, a_, b_, src in LIG:
    val(ws, i, 1, bloc_)
    val(ws, i, 2, el)
    val(ws, i, 3, a_, FMT)
    val(ws, i, 4, b_, FMT)
    val(ws, i, 5, f"=C{i}-D{i}" if b_ is not None else None, FMT)
    val(ws, i, 6, src)
    i += 1
i += 2
val(ws, i, 1, "Déclarations mensuelles des salaires (EDI ITS)", gras=True)
i += 1
entete(ws, i, ["Mois", "Salariés", "Sans numéro CNPS", "Brut du mois (FCFA)", "Part sans numéro", ""], [22, 48, 22, 22, 18, 44])
for m_, n_, f_, b_ in [("Septembre 2024", 61, 4, 17326865), ("Octobre 2024", 61, 4, 19097548), ("Novembre 2024", 61, 4, 17171207), ("Janvier 2025", 58, 2, 16332109),
                       ("Février 2025", 60, 4, 16796156), ("Mars 2025", 61, 5, 17112262), ("Avril 2025", 62, 6, 16922659), ("Mai 2025", 64, 8, 17565658),
                       ("Juin 2025", 66, 10, 17601063), ("Juillet 2025", 78, 7, 21601248), ("Août 2025", 79, 8, 21866911), ("Septembre 2025", 79, 8, 20283633),
                       ("Octobre 2025", 79, 8, 19944900), ("Novembre 2025", 77, 7, 22734544), ("Décembre 2025", 79, 8, 45545641), ("Janvier 2026", 79, 10, 24649368),
                       ("Février 2026", 82, 12, 40287208), ("Mars 2026", 82, 14, 24528001), ("Avril 2026", 82, 14, 24723969), ("Mai 2026", 83, 15, 24989615),
                       ("Juin 2026", 83, 16, 25133146), ("Juillet 2026", 100, 33, 30999904), ("Août 2026", 100, 33, 28171544), ("Septembre 2026", 98, 30, 27940045)]:
    i += 1
    val(ws, i, 1, m_)
    val(ws, i, 2, n_, FMT)
    val(ws, i, 3, f_, FMT)
    val(ws, i, 4, b_, FMT)
    val(ws, i, 5, f"=C{i}/B{i}", PCT)

# ------------------------------------------------------------------ ordre, guide, traçabilité
for nom in ("30_pidr_fppn", "31_ratios_fppn", "32_actifs_contestes", "33_grenier_ef", "34_documents_0911", "35_sinistralite", "36_documents_1111"):
    wb.move_sheet(wb[nom], offset=wb.sheetnames.index("13a_methode") - wb.sheetnames.index(nom))
gd = wb["0_guide"]
rg = gd.max_row + 2
gd.cell(rg, 1, "Compléments du 10/11/2026 (documents reçus le 09/11/2026, en bleu)").font = Font(name="Helvetica", size=11, bold=True, color=BLEU)
for k, (nom, txt) in enumerate((("30_pidr_fppn", "Provision retraite calculée par le FPPN, compte 196 et recalcul Adenka."),
                                 ("31_ratios_fppn", "Ratios prudentiels du FPPN comparés aux ratios recalculés."),
                                 ("32_actifs_contestes", "Actifs dont la valeur n'est pas établie et sensibilité des fonds propres."),
                                 ("33_grenier_ef", "Comptes propres du Grenier du Policier, 2024 et 2025."),
                                 ("34_documents_0911", "Documents reçus le 09/11/2026 et usage dans le rapport."),
                                 ("35_sinistralite", "Prestations maladie par nature réelle et sinistralité des soins seuls (11/11/2026)."),
                                 ("36_documents_1111", "Documents reçus le 11/11/2026 : comptes 47 et 27, siège, Grenier, salariés sans numéro CNPS.")), rg + 1):
    c = gd.cell(k, 1, nom)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{nom}'!A1")
    c.font = Font(name="Helvetica", size=9, color=BLEU, underline="single")
    gd.cell(k, 2, txt).font = FB


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
_der = max(r for r in range(1, _lm.max_row + 1) if _lm.cell(r, 1).value)
if not str(_lm.cell(_der, 1).value).startswith("Onglets 30 à 34"):
    c = _lm.cell(_der + 1, 1, "Onglets 30 à 34 (en bleu) : documents reçus le 09/11/2026 : provision retraite du FPPN (30), ratios du FPPN (31), actifs dont la valeur n'est pas établie (32), comptes du Grenier (33), inventaire des documents (34).")
    c._style = copy(_lm.cell(_der, 1)._style)
    c.font = Font(name=c.font.name, size=c.font.size, color=BLEU)
_reecrire(wb["9_tracabilite"], [(t["diapositive"], t["indicateur"], t["valeur_exacte"], t["texte"], "Oui" if t["retrouve"] else "Non") for t in IDX["trace"]])
_reecrire(wb["9_figures_tableaux"], [("Figure", *x) for x in IDX["figures"]] + [("Tableau", *x) for x in IDX["tableaux"]])
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from depersonnaliser import dep as _dep  # noqa: E402
for _ws in wb.worksheets:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and not _c.value.startswith("="):
                _c.value = _dep(_c.value)
wb.save(SORTIE)
print("classeur écrit :", SORTIE)
