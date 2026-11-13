# -*- coding: utf-8 -*-
"""
FPPN v13 (13/11/2026) : classeur de calcul et liste des demandes au FPPN datés du 13/11/2026 (nouvelles versions, sources non modifiées).
- Classeur : copie du classeur du 11/11/2026 (formules conservées) + onglet 37_extourne_v13 (chiffres v13 après extourne).
- Demandes : copie de la liste du 11/11/2026 + questions nouvelles issues de la vérification de la revue (lignes en bleu).
Utilisation : python livrables_v13_fppn.py <chiffres_v13.json>
"""
import json, shutil, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

K = json.load(open(sys.argv[1], encoding="utf-8"))
BLEU = Font(name="Helvetica", size=10, color="0070C0")
ENT = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
NOIR = PatternFill("solid", fgColor="060606")

shutil.copy2("rapport/FPPN_Calculs_rapport_20261111.xlsx", "rapport/FPPN_Calculs_rapport_20261113.xlsx")
wb = openpyxl.load_workbook("rapport/FPPN_Calculs_rapport_20261113.xlsx")
ws = wb.create_sheet("37_extourne_v13", 1)
ws.append(["Chiffres de la version du 13/11/2026 après extourne de la provision de la clinique (OD 805 à 807 du 31/12/2025, 255 000 000 FCFA)"])
ws.append(["Les autres onglets donnent les comptes tels quels (version du 11/11/2026). Calcul : scripts/v13/chiffres_v13_fppn.py sur la base retraitée (base_v13_fppn.py)."])
ws.append([])
ws.append(["Indicateur", "Valeur", "Unité"])
LIB = {"sc_mal_26": ("Prestations maladie pour 100 FCFA de cotisations maladie, 31/08/2026", "FCFA pour 100"),
       "regl_26": ("Prestations pour 100 FCFA de cotisations, formule du texte, 31/08/2026", "FCFA pour 100"),
       "res_tech_26": ("Résultat technique maladie et décès au 31/08/2026", "FCFA"),
       "res_tech_26_dec": ("Résultat technique prolongé à fin décembre au même rythme (× 12 / 8)", "FCFA"),
       "res_net_26": ("Résultat net au 31/08/2026", "FCFA"), "dette_cliniques_26": ("Factures des cliniques à payer, 40110000", "FCFA"),
       "dettes_hors_epargne": ("Dettes à moins d'un an hors épargne", "FCFA"), "couv_comptes": ("Argent disponible pour 100 FCFA à payer, lecture A (bilan tel quel)", "FCFA pour 100"),
       "couv_besoin": ("Argent disponible pour 100 FCFA à payer, lecture B (provisions au besoin)", "FCFA pour 100"),
       "couv_hors_epargne": ("Argent disponible pour 100 FCFA de dettes hors épargne", "FCFA pour 100"),
       "manque_5_ans_besoin": ("Écart cumulé à cinq ans, lecture B", "FCFA"), "interets_pep_dus": ("Intérêts du plan d'épargne dus, règle des entretiens (3,5 % une fois)", "FCFA"),
       "interets_payes_pep": ("Intérêts payés sur le compte 41195100, 09/2020 au 31/08/2026", "FCFA"),
       "fonds_propres_26": ("Fonds propres au 31/08/2026, extourne comprise", "FCFA"), "fp_corrige_avant_actifs": ("Fonds propres corrigés avant valeur des biens", "FCFA"),
       "fp_sans_contestes": ("Fonds propres sans les terrains sans titre et la créance sur l'hôpital", "FCFA"),
       "seuil_retrait": ("Retrait d'épargne qui vide la trésorerie disponible", "%")}
for k, (lib, u) in LIB.items():
    ws.append([lib, round(K[k], 2), u])
for c in ws[4]:
    c.font, c.fill = ENT, NOIR
for r in ws.iter_rows(min_row=5):
    for c in r:
        c.font = BLEU
ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 80, 20, 16
wb.save("rapport/FPPN_Calculs_rapport_20261113.xlsx")

shutil.copy2("FPPN_Demandes_au_FPPN_20261111.xlsx", "FPPN_Demandes_au_FPPN_20261113.xlsx")
wb = openpyxl.load_workbook("FPPN_Demandes_au_FPPN_20261113.xlsx")
ws = wb["1_Questions"]
n0 = max(int(r[0]) for r in ws.iter_rows(min_row=5, values_only=True) if r[0] and str(r[0]).isdigit())
NOUV = [("Épargne", "Quel texte fixe la règle des intérêts du plan d'épargne (3,5 % du montant versé, une seule fois) ? Le manuel écrit 3,5 % « tous les ans ».", "99", "Urgent"),
        ("Épargne", "Quel est le règlement de l'entraide de fin de carrière : montant versé à la sortie, intérêts, bonus ?", "99", "Urgent"),
        ("Provisions", "Pourquoi la provision de la clinique du 31/12/2025 (OD 805 à 807, 255 000 000 FCFA) n'a-t-elle pas été extournée en 2026 ?", "", "Urgent"),
        ("Provisions", "La facture générale de la clinique de février 2026 a-t-elle été enregistrée ? Merci de lister les factures de janvier à août 2026.", "", "Moyen"),
        ("Décès", "Comment rapprocher l'état des décès au 30/06/2026 (55 600 000 FCFA) des comptes 6021 et 6024 (152 816 666 FCFA au 31/08/2026) ?", "", "Moyen"),
        ("Décès", "Quelles sont les dates de dépôt des dossiers décès antérieurs à 2021 encore dus (257 683 332 FCFA) ?", "", "Moyen"),
        ("Placements", "Quelle décision affecte le fonds commun de placement à l'épargne ? Quelle est sa valeur liquidative au 31/12/2025 et au 31/08/2026 ?", "", "Moyen"),
        ("Actifs", "Quelle est l'écriture d'origine de la réévaluation du siège (10620000) et le détail du compte 22200000 par site ?", "59", "Urgent"),
        ("Résultat", "Quelle résolution de l'assemblée du 17/04/2026 affecte le résultat 2025 ?", "", "Urgent"),
        ("Prestations", "Quelle est la nature des « honoraires médicaux » mensuels (3 390 000 FCFA en 2026) et des honoraires trimestriels « Me … / RECO » ?", "", "Moyen"),
        ("Grenier", "Quels prêts du Grenier ont été rééchelonnés ou passés en perte en 2026 ?", "", "Moyen"),
        ("Épargne", "Quelles pièces justifient les dons, gadgets et honoraires payés sur les comptes de soins en 2025 (487 492 003 FCFA) ?", "", "Urgent")]
for k, (th, q, pbc, pr) in enumerate(NOUV, 1):
    ws.append([str(n0 + k), th, q, pbc, pr, "", "", ""])
    for c in ws[ws.max_row]:
        c.font = BLEU
        c.alignment = Alignment(wrap_text=True, vertical="top")
ws["A2"] = "Réponse écrite attendue, avec la pièce qui l'établit. Questions nouvelles du 13/11/2026 en bleu."
wb.save("FPPN_Demandes_au_FPPN_20261113.xlsx")
print("ok", n0 + len(NOUV), "questions")
