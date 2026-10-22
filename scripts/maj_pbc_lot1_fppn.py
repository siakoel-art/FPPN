# -*- coding: utf-8 -*-
"""
FPPN - Pointage de la PBC actuarielle sur le contenu du lot 1 des nouveaux fichiers (22/10/2026)
===============================================================================================
Le lot 1 (61 fichiers, empreintes SHA-256 conformes au journal « journal_lot1.csv ») a été lu : le statut de chaque
item reflète désormais le contenu, et non plus le seul nom des fichiers. Ajoute des points à clarifier et l'onglet
« Pointage lot 1 (contenu) ». Un statut n'est jamais abaissé sans explication écrite dans le commentaire.
Utilisation : python maj_pbc_lot1_fppn.py <pbc.xlsx>   (modifié en place)
"""
import datetime as dt
import sys
from copy import copy

import openpyxl

PBC = sys.argv[1]
R = dt.datetime(2026, 10, 22)
CDC_M = "« CAHIER DE CHARGE  MALADIES ET INVAL-DECES 2021.pdf »"
CDC_E = "« CAHIER DE CHARGEEFC & PEP 2021.pdf »"
MAN = "« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx »"
CARTO = "« CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx »"
ETUDE = "« ETUDE DES RATIO MALADIE DECES 2.xlsx »"
ERB = "« ETAT DE RAPPROCHEMENT BANCAIRE FPPN 2020.xlsx » à « … 2026.xlsx »"
MAJ = {  # item : (statut, constat sur le contenu)
    100: ("Reçu", "Lu : cahiers des charges 2021 " + CDC_M + " (risques couverts, ticket modérateur 80 % / 20 %, exclusions, forclusions, procédures) et " + CDC_E
          + ". Le tableau des allocations n'y figure pas (renvoi) : il est dans " + CARTO + ". Versions antérieures à 2021 non reçues."),
    101: ("Reçu", "Lu : " + CARTO + " donne les allocations décès par grade (commissaire 2 200 000, officier 1 600 000, sous-officier 1 400 000 FCFA ; conjoint 1 000 000 / 700 000 / 600 000 ; "
          "enfant 230 000 / 150 000 / 120 000), pour 2021 à 2023 et 2023 à 2026, avec des frais funéraires de 300 000 FCFA pour le membre sur 2023 à 2026. Allocation d'invalidité non chiffrée."),
    93: ("Reçu", "Lu : barème par grade dans " + CARTO + " ; mêmes montants appliqués dans les notes de " + ETUDE + " (32 décès au 30/06/2026)."),
    30: ("Partiellement Reçu", "Barèmes 2021 à 2026 reçus (" + CARTO + ") ; barèmes successifs de 2001 à 2020 et décisions non reçus."),
    2: ("Reçu", "Lu : " + CDC_M + ", " + CDC_E + " et " + CARTO + " (maternité : 3 échographies, 3 consultations, allocation d'accouchement de 100 000 FCFA)."),
    16: ("Reçu", "Lu : " + CDC_E + " et " + MAN + ". Taux servis contradictoires : voir le point à clarifier sur l'EFC."),
    46: ("Reçu", "Lu : PEP 1 à 3 (10 000 à 50 000 FCFA et plus, durée de 3 à 11 ans, rachats partiels limités à deux, pénalité de 10 %, rupture anticipée pénalisée de 14 %) ; "
         "EFC 1 à 3 (6 000 ou 10 000 FCFA par mois jusqu'à la retraite, retrait seulement en fin de carrière)."),
    102: ("Reçu", "Lu : " + MAN + " (version 2021, certificat d'adoption non daté ni signé) : PEP rémunéré à 3,5 % par an, EFC à 5,5 % par an, intérêts non capitalisés et versés sur le compte courant."),
    18: ("Partiellement Reçu", "Trois taux écrits pour l'EFC : 5,5 % (" + MAN + " et " + CARTO + ", « à partir de 2024 »), intérêt simple de 14 % au-delà de 2024 et pour les souscriptions postérieures au cahier des charges (" + CDC_E
         + "), 3,5 % selon le chef comptable (19/10/2026). PEP : 3,5 % (" + MAN + "). Décision fixant le taux appliqué chaque année non reçue."),
    99: ("Partiellement Reçu", "Le taux de 3,5 % n'est écrit que pour le PEP (" + MAN + ") ; pour l'EFC, les textes donnent 5,5 % ou 14 % simple. Décision non reçue."),
    57: ("Partiellement Reçu", "EFC, PEP (" + CDC_E + ", " + MAN + "), Grenier (« POLITIK DE CREDIT DU GRENIER VFVF.doc ») et flotte automobile (" + MAN + ", chapitre II) documentés ; MEDOC non documenté."),
    74: ("Reçu", "Reçu : « POLITIK DE CREDIT DU GRENIER VFVF.doc » (environ 98 000 caractères) ; lecture détaillée à faire pour la dépréciation des prêts."),
    27: ("Partiellement Reçu", "Lu : " + ETUDE + " : cotisations et sinistres décès, nombre d'adhérents et de décès (adhérents, conjoints, enfants) de 2015 au 30/06/2026 ; notes de calcul de la sinistralité "
         "décès (13 %) et maladie (99,23 %) au 30/06/2026. Détail mensuel maladie non reçu."),
    7: ("Partiellement Reçu", "Études internes reçues (" + ETUDE + ", « RAPPORT GLOBAL SUR LE SERVICE ASSURANCE MALADIE », « INTERPRETATION DES RATIOS… ») ; aucun rapport actuariel ni calcul des provisions."),
    33: ("Reçu", "Lu : trois listes de pièces (décès d'un membre, d'un ayant droit, frais funéraires) et pièces exigées dans " + CDC_M + "."),
    76: ("Partiellement Reçu", "Lu : « DISA_256927_2021.xls » à « DISA_256927_2025.xls » (DISA 2024 dans « ETAT 301 & DISA 2024 FPPN.xls ») : nom, année de naissance, date d'embauche, salaire brut annuel, "
         "durée d'activité par salarié déclaré à la CNPS (80 lignes en 2025). Pas de sexe, de catégorie ni d'âge de départ ; une ligne agrégée « travailleurs occasionnels » ; situation au 31/08/2026 non reçue."),
    103: ("Partiellement Reçu", "Les DISA couvrent les salariés déclarés à la CNPS. Les états 301 distinguent des onglets « RECAPSALAIRES 2024 CMP » et « RECAPSALAIRES 2025 GRENIER » : Direction Générale, clinique et Grenier "
          "sont identifiables par nom. Le statut des policiers en service n'y figure pas."),
    97: ("Partiellement Reçu", "Lu : états 301 et DISA 2020 à 2025, appels CNPS annuels, DASC 2022, 2023, 2025, fichiers e-CNPS complémentaires 2022. Déclarations mensuelles CNPS et 2026 non reçues."),
    20: ("Partiellement Reçu", "Lu : " + ERB + " : solde bancaire et solde comptable mensuels par compte (BOA EFC, BOA PEP, BOA BF, BOA FPPN, Orabank, Banque Atlantique). Échéancier des placements (DAT, FCP) non reçu."),
    80: ("Partiellement Reçu", "Mouvements mensuels des comptes BOA lisibles dans " + ERB + " ; conventions, taux et relevés du DAT non reçus."),
    91: ("Partiellement Reçu", "Les rapprochements bancaires listent des chèques en suspens, dont certains émis depuis décembre 2021 sur le compte BOA EFC et encore non débités au 31/12/2025 (" + ERB + ")."),
}
POINTS = [
    ("Taux servi sur l'EFC", "Trois taux écrits : 5,5 % par an (manuel d'épargne 2021, cartographie), intérêt simple de 14 % (cahier des charges EFC et PEP 2021, au-delà de 2024 et pour les nouvelles souscriptions), 3,5 % (chef comptable). Bonus de 300 000 FCFA cité par le cahier des charges.",
     "Indiquer le taux appliqué chaque année de 2020 à 2026, la décision qui le fixe et le calcul du capital EFC servi à la retraite (items 18 et 99)."),
    ("Manuel d'épargne", "Le certificat d'adoption du manuel des produits d'épargne (2021) n'est ni daté ni signé.", "Transmettre la délibération d'adoption."),
    ("Ayants droit", "Règlement intérieur (art. 40) : enfants de 0 à 18 ans, cinq au plus (soins), quatre de 7 à 18 ans (décès). Cahier des charges 2021 : enfants mineurs de moins de 21 ans. Cartographie : six enfants de 5 à 21 ans.",
     "Préciser la règle en vigueur pour les soins et pour les allocations décès."),
    ("Frais funéraires du membre", "La cartographie prévoit 300 000 FCFA de frais funéraires pour le membre sur 2023 à 2026 ; l'étude des ratios les applique au 30/06/2026 ; le tableau 2021 à 2023 ne les mentionne pas.", "Confirmer la date d'entrée en vigueur."),
    ("Chèques en suspens", "Des chèques émis depuis décembre 2021 restent en suspens sur le compte BOA EFC au 31/12/2025.", "Expliquer ces chèques et leur traitement (item 91)."),
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
    rang = {"Non Reçu": 0, "Partiellement Reçu": 1, "Reçu": 2}
    trace = []
    for item, (statut, texte) in sorted(MAJ.items()):
        r = rows[item]
        avant = ws.cell(r, 13).value
        if rang.get(statut, 0) > rang.get(avant, 0):
            ws.cell(r, 13).value = statut
            ws.cell(r, 9).value = R
        ancien = ws.cell(r, 11).value or ""
        ws.cell(r, 11).value = (ancien + "\n" if ancien else "") + "Lot 1 lu le 22/10/2026 : " + texte
        trace.append((item, str(ws.cell(r, 3).value)[:160], avant, ws.cell(r, 13).value, texte))
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
    nom = "Pointage lot 1 (contenu)"
    if nom in wb.sheetnames:
        del wb[nom]
    modele = wb["Pointage autres équipes"]
    po = wb.create_sheet(nom)
    po.sheet_view.showGridLines = False
    po["B2"] = "Pointage sur le contenu du lot 1 des nouveaux fichiers (22/10/2026)"
    po["B2"].font = copy(modele["B2"].font)
    po["B3"] = ("Lot 1 : 61 fichiers reçus, empreintes SHA-256 conformes au journal « journal_lot1.csv ». Chaque statut repose sur la lecture du contenu. "
                "Contrôle croisé : les nouvelles balances 2021 et 2025 sont identiques, compte par compte, aux balances déjà exploitées (aucun écart de solde).")
    po["B3"].font, po["B3"].alignment = copy(modele["B3"].font), copy(modele["B3"].alignment)
    po.merge_cells("B3:F3")
    po.row_dimensions[3].height = 40
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
    print(f"{len(trace)} items, {sum(1 for t in trace if t[2] != t[3])} statuts relevés, {len(POINTS)} points ajoutés")


if __name__ == "__main__":
    main()
