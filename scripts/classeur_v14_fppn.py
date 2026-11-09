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

# ------------------------------------------------------------------ ordre, guide, traçabilité
for nom in ("30_pidr_fppn", "31_ratios_fppn", "32_actifs_contestes", "33_grenier_ef", "34_documents_0911"):
    wb.move_sheet(wb[nom], offset=wb.sheetnames.index("13a_methode") - wb.sheetnames.index(nom))
gd = wb["0_guide"]
rg = gd.max_row + 2
gd.cell(rg, 1, "Compléments du 10/11/2026 (documents reçus le 09/11/2026, en bleu)").font = Font(name="Helvetica", size=11, bold=True, color=BLEU)
for k, (nom, txt) in enumerate((("30_pidr_fppn", "Provision retraite calculée par le FPPN, compte 196 et recalcul Adenka."),
                                 ("31_ratios_fppn", "Ratios prudentiels du FPPN comparés aux ratios recalculés."),
                                 ("32_actifs_contestes", "Actifs dont la valeur n'est pas établie et sensibilité des fonds propres."),
                                 ("33_grenier_ef", "Comptes propres du Grenier du Policier, 2024 et 2025."),
                                 ("34_documents_0911", "Documents reçus le 09/11/2026 et usage dans le rapport.")), rg + 1):
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
