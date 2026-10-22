# -*- coding: utf-8 -*-
"""
FPPN - Pointage de la PBC actuarielle sur l'inventaire des nouveaux fichiers reçus le 22/10/2026
==============================================================================================
Source : « FPPN_inventaire_nouveaux_fichiers_20261022.xlsx » (113 fichiers, inventaire produit par
inventaire_nouveaux_fppn.py sur les dossiers « 02 » à « 07 » de « 04-Documents reçus du client »).
Le pointage repose sur le nom du fichier, le nombre de pages ou de feuilles et la première page lisible :
le contenu reste à analyser. En cas de doute sur la couverture, le statut est « Partiellement Reçu ».
Onglet ajouté : « Pointage nouveaux fichiers ».
Utilisation : python maj_pbc_nouveaux_fichiers_fppn.py <pbc.xlsx>   (modifié en place)
"""
import datetime as dt
import sys
from copy import copy

import openpyxl

PBC = sys.argv[1]
R = dt.datetime(2026, 10, 22)
J, P, B = "03-Nouveaux Juridiques", "05-Nouveaux Prestations", "04-Nouveaux Personnel et paie"
A = " Contenu à analyser."
MAJ = {  # item : (statut, fichiers, commentaire)
    100: ("Reçu", J + " : « CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf » (17 p.) ; « CAHIER DE CHARGEEFC & PEP 2021.pdf » (10 p.)",
          "Cahiers des charges du régime des prestations sociales, version 2021. Versions antérieures à 2021 non reçues." + A),
    2: ("Reçu", J + " : cahiers des charges 2021 ; " + P + " : « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx »",
        "Règles et barèmes des prestations (maladie, invalidité-décès, EFC, PEP)." + A),
    101: ("Partiellement Reçu", J + " : « CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf »",
          "Le tableau des allocations doit figurer dans le cahier des charges 2021 ; versions 2020 et postérieures à vérifier." + A),
    30: ("Partiellement Reçu", J + " : cahiers des charges 2021 ; " + P + " : cartographie des risques", "Barèmes 2021 reçus ; barèmes successifs depuis 2001 non reçus." + A),
    93: ("Partiellement Reçu", J + " : « CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf »", "Barème des allocations décès par grade à lire dans le cahier des charges 2021." + A),
    16: ("Reçu", J + " : « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx » ; « CAHIER DE CHARGEEFC & PEP 2021.pdf »",
         "Description des produits d'épargne EFC et PEP." + A),
    46: ("Reçu", J + " : « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx » ; « CAHIER DE CHARGEEFC & PEP 2021.pdf »", "Règlement du PEP et de l'EFC." + A),
    57: ("Partiellement Reçu", J + " : manuel de produit d'épargne, cahiers des charges 2021, « POLITIK DE CREDIT DU GRENIER VFVF.doc »",
         "EFC, PEP et Grenier documentés ; MEDOC et flotte automobile non documentés." + A),
    102: ("Reçu", J + " : « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx »", "Manuel des produits d'épargne (PEP et EFC)." + A),
    18: ("Partiellement Reçu", J + " : « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx » ; cahier des charges EFC et PEP 2021",
         "Taux crédité à rechercher dans le manuel et le cahier des charges ; décisions annuelles non reçues." + A),
    99: ("Partiellement Reçu", J + " : manuel de produit d'épargne ; cahier des charges EFC et PEP 2021", "Taux servi de 3,5 % à vérifier dans ces documents ; décision non reçue." + A),
    74: ("Reçu", J + " : « POLITIK DE CREDIT DU GRENIER VFVF.doc »", "Politique de crédit du Grenier." + A),
    106: (None, "", ""),
    44: ("Partiellement Reçu", J + " : « NUMERO D'IMMATRICULATION AIRSM N° 1D-0052016-CI DU 05-04-2017.pdf » (scan)",
         "Attestation d'immatriculation n° 1D/0052016/CI du 05/04/2017 reçue ; décision d'agrément et liste des risques autorisés à vérifier." + A),
    47: ("Partiellement Reçu", J + " : « ETATS FINANCIERS 2020.pdf » à « ETATS FINANCIERS 2025.pdf », « ETATS FINANCIERS AU 31-08-2026.pdf » (scans) ; « BILAN_DE_L'EXERCICE_CLOS_31.12.2023_(corrigé_et_finalisé)[1].pdf »",
         "États financiers 2020 au 31/08/2026 reçus ; rapports du commissaire aux comptes et états par risque long non reçus." + A),
    107: ("Partiellement Reçu", J + " : « PV AG ORDINAIRE DU 15 FEVRIER 2024.pdf » (scan)", "PV de l'AG ordinaire du 15/02/2024 reçu ; autres AG d'affectation non reçues." + A),
    50: ("Partiellement Reçu", J + " : « PV AG ORDINAIRE DU 15 FEVRIER 2024.pdf » (scan)", "AG du 15/02/2024 : affectation à vérifier ; origine du compte 10620000 non documentée." + A),
    53: ("Partiellement Reçu", J + " : « PV AG ORDINAIRE DU 15 FEVRIER 2024.pdf » ; listes des administrateurs et membres de la commission de surveillance",
         "PV de 2024 et compositions des organes reçus ; calendrier statutaire et autres PV à compléter." + A),
    90: ("Reçu", J + " : « LISTE DES COMPTES BANCAIRES  DU FPPN.pdf » ; « Habilitation de signatures.pdf » (scans)", "Liste des comptes bancaires et habilitations de signature." + A),
    84: ("Partiellement Reçu", J + " : « Certificat de Propriété Terrain Marcori (Siège).pdf » ; « Certificat de Propriété Terrain Plateau.pdf » ; « Vente ferme par les consorts SIRANSY TOURE SOULEYMANE… »",
         "Deux certificats de propriété et un acte de vente reçus ; situation de chaque terrain (projet, attribution, litige) non reçue." + A),
    23: ("Partiellement Reçu", J + " : certificats de propriété, acte de vente, « CONTRAT DE BAIL SONAPIE N° 2124-03002 / 03037 FPPN.pdf »",
         "Titres et baux reçus ; valeurs et dernières évaluations non reçues." + A),
    4: ("Partiellement Reçu", J + " : « CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf » (scan, 9 p.)", "Convention avec la CGRAE reçue ; avenants à vérifier." + A),
    88: ("Partiellement Reçu", J + " : « CONVENTION TRANSVIE-CMP.pdf » (scan)", "Convention de la clinique (CMP) avec Transvie reçue ; bilan d'ouverture et détail des comptes 47111104 et 47111105 non reçus." + A),
    83: ("Partiellement Reçu", J + " : « CONVENTION TRANSVIE-CMP.pdf »", "La convention doit préciser la dénomination et le statut de la structure CMP." + A),
    97: ("Partiellement Reçu", B + " : « ETAT 301 & DISA 2020 à 2025 FPPN.xls », « DISA_256927_2021 à 2025.xls », DASC 2022, 2023, 2025",
         "DISA, état 301 et appels CNPS annuels reçus pour 2020 à 2025 ; déclarations mensuelles CNPS et 2026 non reçues." + A),
    76: ("Partiellement Reçu", B + " : « DISA_256927_2021.xls » à « DISA_256927_2025.xls » (colonnes ANNEE NAISS., DATE EMBAUCHE, DATE DEPART, SALAIRE BRUT)",
         "Les DISA donnent, pour chaque salarié déclaré à la CNPS, l'année de naissance, la date d'embauche et le salaire brut annuel : base possible de la provision pour départ à la retraite. Situation au 31/08/2026 non reçue." + A),
    103: ("Partiellement Reçu", B + " : DISA 2021 à 2025 ; « RECAPSALAIRES » des états 301",
          "Les DISA couvrent les salariés déclarés à la CNPS (personnel civil) ; les fonctionnaires de police en service n'y figurent a priori pas. Statut à confirmer." + A),
    96: (None, "", ""),
    33: ("Reçu", P + " : « Liste des pièces à fournir en cas de décès d'un membre participant.pdf », « … d'un ayant-droit.pdf », « … pour les frais funeraires.pdf »",
         "Listes des pièces exigées pour les prestations décès ; cahier des charges reçu (item 100)."),
    27: ("Partiellement Reçu", P + " : « ETUDE DES RATIO MALADIE DECES 2.xlsx » (onglets Tableau, Notes, Historique 2015-2021) ; « INTERPRETATION DES RATIOS… » (.docx, .pdf)",
         "Sinistralité observée par le FPPN (au 30 juin, historique 2015 à 2021) et interprétation : méthode du FPPN à confronter à la formule de la Décision n° 05/2022." + A),
    7: ("Partiellement Reçu", P + " : « ETUDE DES RATIO MALADIE DECES 2.xlsx » ; « RAPPORT GLOBAL SUR LE SERVICE ASSURANCE MALADIE » (.docx, .pdf)",
        "Études techniques internes reçues ; aucun rapport actuariel ni calcul des provisions." + A),
    26: ("Partiellement Reçu", "« ORGANIGRAMME FPPN 2021(1).pdf » ; " + J + " : « OrganigrammeFPPN_ORI (A) 2020.docx » ; « COURRIER DE NOMINATION DÉFINITIVE.pdf »",
         "Organigrammes 2020 et 2021 et un courrier de nomination reçus ; contacts des responsables non reçus." + A),
    34: ("Partiellement Reçu", P + " : « RAPPORT GLOBAL SUR LE SERVICE ASSURANCE MALADIE » ; listes des pièces décès", "Documents décrivant le traitement des prestations." + A),
    20: ("Partiellement Reçu", "07-Nouveaux Trésorerie et placements : « ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2020.xlsx » à « … 2026.xlsx » (onglets par banque : BOA PEP, BOA EFC, BOA BF, BOA FPPN, ORABANK, BANQUE ATLANTIQUE)",
         "États de rapprochement bancaire 2020 à 2026 reçus ; échéancier des placements (DAT, FCP) non reçu." + A),
    80: ("Partiellement Reçu", "07-Nouveaux Trésorerie et placements : états de rapprochement bancaire 2020 à 2026", "Mouvements des comptes BOA à lire ; conventions et relevés du DAT non reçus." + A),
}
MAJ = {k: v for k, v in MAJ.items() if v[0]}


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
    rang = {"Non Reçu": 0, "Partiellement Reçu": 1, "Reçu": 2}
    for item, (statut, fichiers, texte) in sorted(MAJ.items()):
        r = rows[item]
        avant = ws.cell(r, 13).value
        if rang.get(statut, 0) > rang.get(avant, 0):  # jamais de baisse de statut
            ws.cell(r, 13).value = statut
            ws.cell(r, 9).value = R
        ancien = ws.cell(r, 11).value or ""
        ws.cell(r, 11).value = (ancien + "\n" if ancien else "") + "Nouveaux fichiers du 22/10/2026 : " + texte + " Fichiers : " + fichiers + "."
        trace.append((item, str(ws.cell(r, 3).value)[:160], avant, ws.cell(r, 13).value, fichiers, texte))
    sy = wb["Synthèse du pointage"]
    par = {}
    for item, r in rows.items():
        par.setdefault(ws.cell(r, 13).value, []).append(str(item))
    for k, st in ((6, "Reçu"), (7, "Partiellement Reçu"), (8, "Non Reçu")):
        sy.cell(k, 4).value = ", ".join(par.get(st, []))
    nom = "Pointage nouveaux fichiers"
    if nom in wb.sheetnames:
        del wb[nom]
    modele = wb["Pointage autres équipes"]
    po = wb.create_sheet(nom)
    po.sheet_view.showGridLines = False
    po["B2"] = "Pointage sur les nouveaux fichiers reçus le 22/10/2026"
    po["B2"].font = copy(modele["B2"].font)
    po["B3"] = ("Source : « FPPN_inventaire_nouveaux_fichiers_20261022.xlsx » (113 fichiers). Pointage fondé sur le nom, la taille et la première page lisible de chaque "
                "fichier : le contenu reste à analyser. Un statut n'est jamais abaissé.")
    po["B3"].font, po["B3"].alignment = copy(modele["B3"].font), copy(modele["B3"].alignment)
    po.merge_cells("B3:G3")
    po.row_dimensions[3].height = 40
    for j, t in enumerate(["Item", "Document demandé (PBC actuarielle)", "Statut avant", "Statut après", "Fichiers reçus", "Commentaire"], 2):
        po.cell(5, j, t)._style = copy(modele.cell(5, j)._style)
    for i, ligne in enumerate(trace, 6):
        for j, v in enumerate(ligne, 2):
            po.cell(i, j, v)._style = copy(modele.cell(6, j)._style)
        po.row_dimensions[i].height = 75
    for col, w in zip("BCDEFG", (7, 48, 16, 16, 60, 60)):
        po.column_dimensions[col].width = w
    po.freeze_panes = "C6"
    po.page_setup.orientation = "landscape"
    po.page_setup.fitToWidth, po.page_setup.fitToHeight = 1, 0
    po.sheet_properties.pageSetUpPr.fitToPage = True
    po.print_title_rows = "5:5"
    wb.save(PBC)
    print(f"{len(trace)} items pointés, {sum(1 for t in trace if t[2] != t[3])} statuts relevés")


if __name__ == "__main__":
    main()
