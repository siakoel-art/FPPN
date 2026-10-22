# -*- coding: utf-8 -*-
"""
FPPN - Mise à jour de la PBC actuarielle à partir du pointage des autres équipes Adenka (22/10/2026)
====================================================================================================
Sources : « FPPN - Suivi demande de documents.xlsx » (équipe d'audit comptable et financier, onglets « Suivi
documents » et « Relance du 211026 ») et « FPPN-CAC 2021-2026 -AUDIT IT PBC List.xlsx » (équipe d'audit IT).
Règle : un item de la PBC actuarielle n'est mis à jour que si une ligne d'une autre équipe porte sur le même
document. Le statut suit celui de l'autre équipe ; le commentaire précise que la copie doit être obtenue auprès
de l'équipe concernée quand le document n'a pas été transmis à l'équipe actuarielle. Onglet ajouté :
« Pointage autres équipes » (item, statut avant et après, source exacte, commentaire).
Utilisation : python maj_pbc_autres_equipes_fppn.py <pbc.xlsx>   (modifié en place)
"""
import datetime as dt
import sys
from copy import copy

import openpyxl

PBC = sys.argv[1]
SUIVI = "« FPPN - Suivi demande de documents.xlsx »"
IT = "« FPPN-CAC 2021-2026 -AUDIT IT PBC List.xlsx »"
SD, RL = SUIVI + ", onglet « Suivi documents »", SUIVI + ", onglet « Relance du 211026 »"
COPIE = " Copie à obtenir auprès de l'équipe d'audit comptable."
D = dt.datetime
# item : (nouveau statut ou None, date de réception ou None, source, commentaire)
MAJ = {
    1: ("Partiellement Reçu", D(2026, 10, 19), SD + ", lignes 1.1 et 1.2",
        "Statuts et règlement intérieur édités le 15/06/2023 (adoptés le 30/06/2021) reçus le 19/10/2026. Versions applicables en 2020 et au premier semestre 2021 non reçues."),
    45: ("Reçu", D(2026, 10, 19), SD + ", ligne 1.1",
         "Barème fixé par les statuts (art. 17) et le règlement intérieur (art. 33) : 3,5 % maladie-maternité et 1,5 % invalidité-décès."),
    3: (None, None, SD + ", ligne 5.2 ; " + RL + ", ligne « Taux et barèmes de cotisation »",
        "Barème de 5 % connu par les statuts. Document « cartographie des risques » contenant des barèmes reçu par l'équipe comptable le 21/10/2026 ; la relance du 21/10/2026 indique le barème encore en attente (M. Soumahoro). Barèmes des retraités et par année non reçus." + COPIE),
    2: ("Partiellement Reçu", D(2026, 10, 20), SD + ", ligne 6.1",
        "Document intitulé « cartographie des risques » contenant les barèmes des prestations, reçu par l'équipe comptable le 20/10/2026." + COPIE + " Cahier des charges (item 100) toujours attendu."),
    30: ("Partiellement Reçu", D(2026, 10, 20), SD + ", ligne 6.1",
         "Barèmes en vigueur reçus dans la « cartographie des risques » (20/10/2026). Barèmes successifs depuis 2001 et décisions non reçus." + COPIE),
    93: ("Partiellement Reçu", D(2026, 10, 20), SD + ", ligne 6.1",
         "Barèmes des prestations reçus dans la « cartographie des risques » (20/10/2026), à confronter au barème par grade demandé." + COPIE),
    101: ("Partiellement Reçu", D(2026, 10, 20), SD + ", ligne 6.1",
          "Barèmes des prestations reçus dans la « cartographie des risques » (20/10/2026) ; tableau des allocations officiel non identifié." + COPIE),
    13: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 6.2",
         "États des dossiers décès et invalidité de 2019 au 30/06/2026, tenus par le service social (M. Soumahoro), reçus par l'équipe comptable. Maladie et maternité : pas de suivi hors comptabilité." + COPIE),
    28: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 6.2",
         "États des dossiers décès et invalidité de 2019 au 30/06/2026 reçus par l'équipe comptable ; format à vérifier au regard du modèle demandé ; juillet et août 2026 manquants." + COPIE),
    71: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 6.2",
         "Décès et invalidité : états 2019 au 30/06/2026 reçus par l'équipe comptable (dates de survenance, déclaration et paiement à vérifier). Maladie : aucun suivi par dossier hors comptabilité, selon le FPPN." + COPIE),
    40: (None, None, SD + ", ligne 6.2",
         "Selon le FPPN, les prestations maladie et maternité ne font l'objet d'aucun suivi par dossier hors comptabilité : la base doit être extraite du logiciel ou reconstituée à partir des factures."),
    14: (None, None, SD + ", ligne 6.3 (Non applicable)",
         "Le FPPN déclare ne tenir aucun état des prestations dues et non réglées : la liste devra être reconstituée à partir des bases de prestations (item 71)."),
    29: (None, None, SD + ", ligne 6.3 (Non applicable)",
         "Le FPPN déclare ne tenir aucun état des prestations dues et non réglées : à reconstituer à partir des états du service social (item 28)."),
    9: (None, None, SD + ", ligne 5.1 ; " + RL + ", ligne 1.2",
        "Fichier des mutualistes partiel selon l'équipe comptable : années 2020 à 2023 et plusieurs mois manquants. Extraction Easy Money de la liste des mutualistes en attente (M. Kouamé, service informatique)."),
    11: (None, None, RL + ", ligne 1.2", "Extraction Easy Money de la liste des mutualistes et de l'état des cotisations en attente (M. Kouamé, service informatique)."),
    12: (None, None, SD + ", ligne 5.3 ; " + RL + ", lignes 1.1 et 2.1",
         "Précomptes 2020 à 2022 (Easy Money) et état des cotisations de la Solde 2020 à 2023 non reçus ; aucun rapprochement entre fichiers retour et comptabilité (relance 2.2 : non applicable)."),
    61: (None, None, RL + ", ligne 1.1", "Extraction Easy Money des fichiers de précomptes adressés à la Solde : exercices 2020 à 2022 non reçus (relance du 21/10/2026)."),
    63: (None, None, RL + ", ligne 1.1 ; " + SD + ", ligne 5.3", "Précomptes 2020 à 2022 et fichiers CGRAE 2020 à 2023 non reçus selon l'équipe comptable."),
    19: (None, None, RL + ", lignes 2.2 et 2.3 (Non applicable)",
         "Le FPPN ne rapproche ni les cotisations ni l'épargne prélevées par la Solde avec les encaissements comptabilisés : le rapprochement devra être fait par Adenka."),
    44: ("Partiellement Reçu", None, SD + ", ligne 1.3",
         "Actes d'agrément et correspondances avec l'AIRMS reçus en version papier par l'équipe comptable ; version numérique en attente." + COPIE),
    8: ("Partiellement Reçu", None, SD + ", ligne 1.3",
        "Correspondances avec la tutelle et l'AIRMS (rapports de contrôle, injonctions) reçues en version papier par l'équipe comptable ; contenu à vérifier." + COPIE),
    35: ("Partiellement Reçu", None, SD + ", ligne 1.3",
         "Rapports de contrôle de l'AIRMS reçus en version papier par l'équipe comptable ; plan de redressement à rechercher." + COPIE),
    95: (None, None, SD + ", ligne 1.3", "À rechercher dans les correspondances avec l'AIRMS reçues en version papier par l'équipe comptable."),
    53: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 1.7 ; " + RL + ", ligne 05.3",
         "Procès-verbaux des AG, CA et commissions reçus par l'équipe comptable, sauf 2020 et 2023." + COPIE),
    107: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 1.7",
          "PV d'AG reçus par l'équipe comptable, sauf 2020 et 2023 ; affectations des résultats à y vérifier." + COPIE),
    50: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 1.7",
         "PV d'AG reçus par l'équipe comptable sauf 2020 et 2023 : l'affectation de 2023 reste à obtenir." + COPIE),
    5: (None, None, SD + ", ligne 1.7", "À rechercher dans les PV reçus par l'équipe comptable (2020 et 2023 manquants)."),
    105: (None, None, SD + ", ligne 1.7", "À rechercher dans les PV du CA reçus par l'équipe comptable (2020 et 2023 manquants)."),
    106: (None, None, SD + ", ligne 1.7", "À rechercher dans les PV d'AG de 2024 et 2025 reçus par l'équipe comptable."),
    109: (None, None, SD + ", ligne 1.7", "À rechercher dans les PV du CA reçus par l'équipe comptable."),
    47: (None, None, SD + ", ligne 1.9 ; " + RL + ", ligne 05.4",
         "États financiers reçus par l'équipe comptable. Rapports du commissaire aux comptes 2022 à 2025 en attente, 2020 et 2021 à retrouver par le FPPN."),
    4: (None, None, SD + ", ligne 1.11", "Tableau récapitulatif des contrats et conventions (dont CGRAE) reçu partiellement par l'équipe comptable ; liste complète en attente."),
    41: (None, None, SD + ", ligne 1.11", "Tableau récapitulatif des conventions (prestataires médicaux) reçu partiellement par l'équipe comptable ; liste complète en attente."),
    89: (None, None, SD + ", ligne 1.11", "Tableau récapitulatif des conventions (dont GNA) reçu partiellement par l'équipe comptable ; liste complète en attente."),
    34: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 3.1 ; " + RL + ", ligne « Manuel de procédures »",
         "Manuels de procédures (adhésions, cotisations, prestations, crédits) reçus non signés par l'équipe comptable ; versions signées à transmettre." + COPIE),
    102: (None, None, SD + ", ligne 3.1", "À rechercher dans les manuels de procédures non signés reçus par l'équipe comptable le 21/10/2026."),
    70: ("Reçu", D(2026, 10, 15), SD + ", ligne 7.1",
         "Règlement des crédits (conditions d'octroi, taux, plafonds, garanties) reçu par l'équipe comptable le 15/10/2026." + COPIE),
    74: ("Partiellement Reçu", D(2026, 10, 15), SD + ", lignes 7.1 et 7.3",
         "Règlement des crédits reçu ; aucune politique écrite de dépréciation (le FPPN cite les garanties et un taux de 1 %)." + COPIE),
    73: ("Partiellement Reçu", D(2026, 10, 15), SD + ", lignes 7.2 et 7.3",
         "Portefeuille des crédits reçu ; politique de dépréciation non tenue (« garanties et 1 % » selon le FPPN)."),
    7: (None, None, SD + ", ligne 8.1", "Rapports actuariels et calcul des provisions : en attente pour l'équipe comptable également."),
    48: (None, None, SD + ", ligne 8.1", "Aucun rapport actuariel transmis à l'équipe comptable (en attente)."),
    90: ("Partiellement Reçu", D(2026, 10, 21), SD + ", ligne 9.1",
         "Liste des comptes bancaires reçue partiellement par l'équipe comptable ; liste par exercice en attente." + COPIE),
    20: (None, None, SD + ", ligne 9.5", "État des placements, dossiers des DAT et tableau de synthèse en attente pour l'équipe comptable également."),
    80: (None, None, SD + ", ligne 9.5", "Dossiers des DAT en attente pour l'équipe comptable également."),
    82: (None, None, SD + ", ligne 9.5", "Tableau de synthèse des placements en attente pour l'équipe comptable également."),
    23: (None, None, SD + ", ligne 13.3", "Dossiers des programmes immobiliers en attente (liste des immeubles : M. Koffi)."),
    84: (None, None, SD + ", ligne 13.4", "Titres de propriété et tableau de synthèse des biens en attente (M. Sangaré)."),
    76: (None, None, SD + ", lignes 15.1 et 15.2", "Mouvements du personnel et livres de paie en attente (M. Koffi)."),
    103: (None, None, SD + ", lignes 15.1 et 15.2", "Mouvements du personnel et livres de paie en attente (M. Koffi) : le statut civil ou policier devra y figurer."),
    78: (None, None, SD + ", ligne 15.4", "Rapport actuariel et extraction de la provision pour départ à la retraite en attente (M. Koffi)."),
    98: (None, None, SD + ", ligne 15.4", "Extraction de la provision pour départ à la retraite en attente (M. Koffi)."),
    97: ("Partiellement Reçu", D(2026, 10, 21), SD + ", lignes 15.5 et 16.2",
         "DISA et état 301 reçus sauf 2026 ; ITS et FDFP reçus (2021 manquant) ; déclarations CNPS en attente." + COPIE),
    33: (None, None, RL + ", lignes 4.1 et 4.2", "Liste des pièces du dossier d'allocation décès et conditions d'octroi en attente (M. Soumahoro)."),
    26: (None, None, SD + ", ligne 1.4", "Organigramme 2022 (identité et postes) reçu partiellement par l'équipe comptable ; version numérique en attente."),
}


def main():
    wb = openpyxl.load_workbook(PBC)
    ws = wb["PBC Actuariat FPPN"]
    rows, n = {}, 0
    for r in range(11, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and b.startswith("="):
            n += 1
            rows[n] = r
    trace = []
    for item, (statut, date, source, texte) in sorted(MAJ.items()):
        r = rows[item]
        avant = ws.cell(r, 13).value
        if statut:
            ws.cell(r, 13).value = statut
        if date:
            ws.cell(r, 9).value = date
        ancien = ws.cell(r, 11).value or ""
        ws.cell(r, 11).value = (ancien + "\n" if ancien else "") + "Pointage des autres équipes (22/10/2026) : " + texte
        trace.append((item, str(ws.cell(r, 3).value)[:160], avant, ws.cell(r, 13).value, source, texte))
    # synthèse : numéros d'items par statut
    sy = wb["Synthèse du pointage"]
    par = {}
    for item, r in rows.items():
        par.setdefault(ws.cell(r, 13).value, []).append(str(item))
    for k, st in ((6, "Reçu"), (7, "Partiellement Reçu"), (8, "Non Reçu")):
        sy.cell(k, 4).value = ", ".join(par.get(st, []))
    sy["B3"] = (str(sy["B3"].value or "") + " Mise à jour du 22/10/2026 à partir du pointage des équipes d'audit comptable et d'audit IT (onglet « Pointage autres équipes »).").strip()
    # onglet de traçabilité
    if "Pointage autres équipes" in wb.sheetnames:
        del wb["Pointage autres équipes"]
    pc = wb["Points à clarifier"]
    po = wb.create_sheet("Pointage autres équipes")
    po.sheet_view.showGridLines = False
    po["B2"] = "Pointage des autres équipes reporté dans la PBC actuarielle (22/10/2026)"
    po["B2"].font = copy(pc["B2"].font)
    po["B3"] = ("Sources : " + SUIVI + " (audit comptable et financier) et " + IT + " (audit IT). Un item n'est mis à jour que si une ligne d'une autre "
                "équipe porte sur le même document. La liste IT ne couvre aucun item actuariel ; le rapport d'audit du système d'information 2022 y figure comme reçu "
                "(déjà recensé dans l'onglet « Reçus hors liste »).")
    po["B3"].font = copy(pc["B3"].font)
    po["B3"].alignment = copy(pc["B3"].alignment)
    po.merge_cells("B3:G3")
    po.row_dimensions[3].height = 45
    tete = ["Item", "Document demandé (PBC actuarielle)", "Statut avant", "Statut après", "Source exacte", "Information reprise"]
    h = 5
    for j, t in enumerate(tete, 2):
        c = po.cell(h, j, t)
        c._style = copy(pc.cell(5, 2)._style)
    for i, ligne in enumerate(trace, h + 1):
        for j, v in enumerate(ligne, 2):
            c = po.cell(i, j, v)
            c._style = copy(pc.cell(6, min(j, 5))._style)
        po.row_dimensions[i].height = 60
    for col, w in zip("BCDEFG", (7, 50, 16, 16, 42, 70)):
        po.column_dimensions[col].width = w
    po.freeze_panes = "C6"
    po.page_setup.orientation = "landscape"
    po.page_setup.fitToWidth, po.page_setup.fitToHeight = 1, 0
    po.sheet_properties.pageSetUpPr.fitToPage = True
    po.print_title_rows = "5:5"
    wb.save(PBC)
    chg = sum(1 for t in trace if t[2] != t[3])
    print(f"{len(trace)} items annotés, {chg} statuts modifiés")


if __name__ == "__main__":
    main()
