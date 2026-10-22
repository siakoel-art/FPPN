# -*- coding: utf-8 -*-
"""
FPPN - Pointage de la PBC actuarielle sur le dossier « 08-Nouveaux Dossiers Compta » reçu le 22/10/2026
======================================================================================================
Huit fichiers : états de fin d'année des allocations décès et frais funéraires 2021 à 2025, arrêté d'autorisation du CMP,
attestation de désignation du directeur général (mandature 2020-2023), dérogation d'urbanisme du projet de Cocody.
Reformule la demande de base des prestations (item 71 : dates des soins, du dépôt et du paiement de chaque facture),
ajoute les demandes issues de la revue actuarielle complète, des points à clarifier et l'onglet de pointage.
Un statut n'est jamais abaissé.
Utilisation : python maj_pbc_lot4_fppn.py <pbc.xlsx>   (modifié en place)
"""
import datetime as dt
import sys
from copy import copy

import openpyxl

PBC = sys.argv[1]
R = dt.datetime(2026, 10, 22)
DEC = "« PROV ALLOCATION DECES 2021.xlsx », « PROV ALLOCATION DECES 2022.xlsx », « ALLOC DECES 2023 PAYER ET A PROVISIONNER .xlsx », « PROV ALLOCATION DECES 2024.xlsx », « PROV ALLOCATIONS DECES ET FRAIS FUNERAIRES 2025.xlsx »"
CMP = "« ARRETE N° 0027-MSHP-DGS-DEPS DU 10-02-2020 PORTANT AUTORISATION D'OUVERTURE D'EXPLOITATION ET IMMATRICULATION DU CMP.pdf »"
COC = "« DEROGATION AUX REGLES D'URBANISME CONSTRUCTION 4 IMMEUBLES R+5 AVEC 01 CES DE 70% A USAGE D'HABITATION COCODY 2 PLTX.pdf »"
DG = "« ATTESTATION DE DESIGNATION DU DGFPPN AU TITRE DE LA MANDATURE 2020-2023.pdf »"
TEXTE_71 = ("Base des prestations maladie, une ligne par facture ou par demande de remboursement, du 01/01/2020 au 31/08/2026 : matricule du mutualiste, bénéficiaire, prestataire, "
            "numéro de facture, date des soins, date de dépôt de la facture au FPPN, date de paiement, montant facturé, montant payé, motif d'un rejet. Base des allocations décès : "
            "une ligne par dossier, date du décès, date de dépôt de la demande, date de paiement, montant. Ces trois dates permettent de calculer la provision maladie et décès avec la méthode courte retenue.")
MAJ = {
    92: ("Partiellement Reçu", "Reçus : " + DEC + ". Calcul de fin d'année au barème, comptabilisé au compte 41591000 (au franc près en 2024 et 2025). Fichiers des décès transmis par la direction du personnel non reçus ; écarts 2021 à 2023 à expliquer."),
    29: ("Partiellement Reçu", "Les états de fin d'année listent les dossiers non payés au 31/12 de 2021 à 2025 ; la liste au 31/08/2026 et les dossiers antérieurs à 2021 ne sont pas reçus."),
    28: ("Partiellement Reçu", "Les états 2021 et 2023 portent la date du décès et la date d'enregistrement ; les états 2022, 2024 et 2025 ne portent pas la date d'enregistrement."),
    72: ("Partiellement Reçu", "Les états de fin d'année expliquent la dette du 41591000 ; aucune note ne fonde la provision obsèques (15802000) ni la provision maladie (15801000)."),
    55: ("Reçu", "Lu : " + CMP + " : « Centre médical du FPPN », établissement sanitaire privé mutualiste, Yopougon Andokoi, immatriculé PV-ABIDJ-MEDIC/CMMU-02978-0022/2020, FPPN promoteur et propriétaire (art. 3). Comptabilité séparée non reçue (voir les nouveaux items)."),
    84: ("Partiellement Reçu", "Lu : " + COC + " : accord de janvier 2025 pour quatre immeubles R+5 sur les lots 3005 à 3007, îlot 83, TF n° 119 230, Cocody. Budget et financement non reçus."),
    53: ("Partiellement Reçu", "Lu : " + DG + " : directeur général désigné le 04/09/2020 pour la mandature 2020-2023 (AG du 27/08/2020) ; désignation après 2023 non reçue."),
}
MAJ = {k: v for k, v in MAJ.items() if v}
NOUVEAUX = [
    ("Liste des dossiers d'allocations décès composant le solde du compte 41591000 au 31/08/2026 (489 419 669 FCFA avec le compte 41592000) : nom, grade, date du décès, date de dépôt de la demande, montant, motif du non-paiement.",
     "Urgent", "Demande issue de la revue des provisions décès (rapport, partie 9)."),
    ("Texte qui fonde l'exclusion du suicide (quatre décès valorisés à 0 FCFA dans l'état 2025) et barème officiel 1 400 000, 1 600 000 et 2 200 000 FCFA appliqué dans les états.",
     "Moyen", "Demande issue de la lecture des états de fin d'année des allocations décès."),
    ("Décision de désignation ou de prolongation du directeur général après la mandature 2020-2023, et délégation de pouvoir pour les placements.",
     "Urgent", "Demande issue de la lecture de l'attestation de désignation du 04/09/2020."),
    ("Projet de Cocody : budget, plan de financement, calendrier, décision des organes, rendement attendu, part financée par l'épargne.",
     "Urgent", "Demande issue de la dérogation d'urbanisme de janvier 2025 (rapport, partie 9.3)."),
    ("Comptabilité séparée du Centre médical du FPPN (CMP) et résultats annuels de 2020 à 2025.",
     "Moyen", "Demande issue de l'arrêté d'autorisation du CMP et du règlement d'exécution n° 004/2022, art. 27."),
    ("Fichiers de la CGRAE avec la date de naissance de tous les retraités (seuls les fichiers retour de février à avril 2024 la portent, pour 105 retraités) et prestations payées aux retraités.",
     "Moyen", "Demande issue de la revue de la santé des retraités (rapport, partie 8)."),
]
POINTS = [
    ("Allocations décès 2021 à 2023", "Crédits du 41591000 différents des totaux des états : 2021 (370 503 332 contre 112 820 000 FCFA), 2022 (97 653 333 contre 88 000 000), 2023 (120 690 000 contre 124 690 000).", "Explication des écarts et pièces des écritures."),
    ("Provision obsèques", "La provision 15802000 (1 961 695 412 FCFA) double la dette du 41591000 déjà comptabilisée.", "Note de calcul ou reprise de la provision."),
    ("Paiements d'allocations par l'épargne", "En 2024, 93 allocations (49 000 000 FCFA) payées par le compte BOA PEP et 9 (4 288 000 FCFA) par le compte BOA EFC.", "Raison et régularisation entre comptes."),
    ("Barème décès", "Barème des états (1 400 000, 1 600 000, 2 200 000 FCFA) différent du barème déclaré en entretien (1 000 000 et 2 000 000 FCFA).", "Barème officiel signé."),
]


def main():
    wb = openpyxl.load_workbook(PBC)
    ws = wb["PBC Actuariat FPPN"]
    rows, n = {}, 0
    for r in range(11, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and b.startswith("="):
            n += 1
            rows[n] = r
    # nouveaux items 110 et suivants, sous une section dédiée
    sect = max(r for r in range(11, ws.max_row + 1) if ws.cell(r, 2).value is None and ws.cell(r, 3).value)
    dern, prec = rows[n], rows[n]
    cols = range(1, ws.max_column + 1)
    st_s = {c: copy(ws.cell(sect, c)._style) for c in cols}
    st_i = {c: copy(ws.cell(dern, c)._style) for c in cols}
    r = ws.max_row + 1
    for c in cols:
        ws.cell(r, c)._style = copy(st_s[c])
    ws.cell(r, 3, "Dossier « 08-Nouveaux Dossiers Compta » reçu le 22/10/2026 (demandes issues de la revue actuarielle)")
    ws.row_dimensions[r].height = ws.row_dimensions[sect].height
    for doc, prio, com in NOUVEAUX:
        r += 1
        for c in cols:
            ws.cell(r, c)._style = copy(st_i[c])
        ws.cell(r, 2, f"=+B{prec}+1")
        ws.cell(r, 3, doc)
        ws.cell(r, 5, prio)
        ws.cell(r, 7, R)
        ws.cell(r, 8, dt.datetime(2026, 10, 29))
        ws.cell(r, 11, com)
        ws.cell(r, 13, "Non Reçu")
        ws.row_dimensions[r].height = 75
        prec = r
        n += 1
        rows[n] = r
    ws.cell(rows[71], 3).value = TEXTE_71
    ws.cell(rows[71], 11).value = (ws.cell(rows[71], 11).value or "") + "\nReformulée le 22/10/2026 : la méthode courte de provision demande, pour chaque facture, la date des soins, du dépôt et du paiement."
    rang = {"Non Reçu": 0, "Partiellement Reçu": 1, "Reçu": 2}
    trace = []
    for item, (statut, texte) in sorted(MAJ.items()):
        r = rows[item]
        avant = ws.cell(r, 13).value
        if rang.get(statut, 0) > rang.get(avant, 0):
            ws.cell(r, 13).value = statut
            ws.cell(r, 9).value = R
        ancien = ws.cell(r, 11).value or ""
        ws.cell(r, 11).value = (ancien + "\n" if ancien else "") + "Dossier « 08-Nouveaux Dossiers Compta » lu le 22/10/2026 : " + texte
        trace.append((item, str(ws.cell(r, 3).value)[:160], avant, ws.cell(r, 13).value, texte))
    for k, (doc, _, com) in enumerate(NOUVEAUX):
        item = n - len(NOUVEAUX) + 1 + k
        trace.append((item, doc[:160], "Nouveau", "Non Reçu", com))
    sy = wb["Synthèse du pointage"]
    par = {}
    for item, r in rows.items():
        par.setdefault(ws.cell(r, 13).value, []).append(str(item))
    for k, st in ((6, "Reçu"), (7, "Partiellement Reçu"), (8, "Non Reçu")):
        sy.cell(k, 4).value = ", ".join(par.get(st, []))
    pc = wb["Points à clarifier"]
    d = pc.max_row
    for k, (sujet, constat, demande) in enumerate(POINTS, 1):
        rr = d + k
        for c in range(2, 6):
            pc.cell(rr, c)._style = copy(pc.cell(d, c)._style)
        pc.row_dimensions[rr].height = 60
        pc.cell(rr, 2, int(pc.cell(d, 2).value) + k)
        pc.cell(rr, 3, sujet)
        pc.cell(rr, 4, constat)
        pc.cell(rr, 5, demande)
    nom = "Pointage dossier compta 22-10"
    if nom in wb.sheetnames:
        del wb[nom]
    modele = wb["Pointage autres équipes"]
    po = wb.create_sheet(nom)
    po.sheet_view.showGridLines = False
    po["B2"] = "Pointage sur le contenu du dossier « 08-Nouveaux Dossiers Compta » (22/10/2026)"
    po["B2"].font = copy(modele["B2"].font)
    po["B3"] = ("Huit fichiers : " + DEC + " ; " + CMP + " ; " + DG + " ; " + COC + ". La demande de base des prestations (item 71) est reformulée par facture avec les dates des soins, "
                "du dépôt et du paiement.")
    po["B3"].font, po["B3"].alignment = copy(modele["B3"].font), copy(modele["B3"].alignment)
    po.merge_cells("B3:F3")
    po.row_dimensions[3].height = 48
    for j, t in enumerate(["Item", "Document demandé (PBC actuarielle)", "Statut avant", "Statut après", "Constat sur le contenu"], 2):
        po.cell(5, j, t)._style = copy(modele.cell(5, j)._style)
    for i, ligne in enumerate(trace, 6):
        for j, v in enumerate(ligne, 2):
            po.cell(i, j, v)._style = copy(modele.cell(6, j)._style)
        po.row_dimensions[i].height = 90
    for col, w in zip("BCDEF", (7, 48, 16, 16, 100)):
        po.column_dimensions[col].width = w
    po.freeze_panes = "C6"
    po.page_setup.orientation = "landscape"
    po.page_setup.fitToWidth, po.page_setup.fitToHeight = 1, 0
    po.sheet_properties.pageSetUpPr.fitToPage = True
    po.print_title_rows = "5:5"
    wb.save(PBC)
    print(f"{len(MAJ)} items pointés, {sum(1 for t in trace if t[2] != t[3] and t[2] != 'Nouveau')} statuts relevés, {len(NOUVEAUX)} items ajoutés (jusqu'à {n}), {len(POINTS)} points ajoutés")


if __name__ == "__main__":
    main()
