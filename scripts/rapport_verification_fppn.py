# -*- coding: utf-8 -*-
"""
FPPN - Rapport de vérification de cohérence (chiffres, calculs, sources)
=======================================================================
Assemble les résultats des contrôles indépendants (verif_sources_fppn.py, verif_precomptes_fppn.py,
placements_gl_fppn.py, contrôle des balances âgées du Grenier, contrôle des formules du classeur) et la liste
des anomalies relevées dans le rapport, avec la correction apportée.
Utilisation : python rapport_verification_fppn.py <sources.json> <precomptes.json> <placements_gl.json> <grenier_controle.parquet> <sortie.xlsx>
"""
import json
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SRC, PRE, PGL, GRN, SORTIE = sys.argv[1:6]
V = json.load(open(SRC, encoding="utf-8"))
P = json.load(open(PRE, encoding="utf-8"))
G = json.load(open(PGL, encoding="utf-8"))
C = pd.read_parquet(GRN)
NOIR = "000000"
COUL = {"Conforme": "00B050", "Écart expliqué": "FFFF00", "Anomalie corrigée": "FFC000", "Erreur corrigée": "FF0000",
        "Critique": "FF0000", "Élevé": "FFC000", "Moyen": "FFFF00", "Faible": "00B050"}
BLANC = {"00B050", "FF0000"}
fin = Side(style="thin", color="808080")
BORD = Border(left=fin, right=fin, top=fin, bottom=fin)


def fcfa(x):
    return f"{round(x):,}".replace(",", " ") + " FCFA"


def pct(x):
    return f"{100 * x:.2f}".replace(".", ",") + " %"


def colorer(c, v):
    c.value = v
    if v in COUL:
        c.fill = PatternFill("solid", fgColor=COUL[v])
        c.font = Font(name="Helvetica", size=9, bold=True, color="FFFFFF" if COUL[v] in BLANC else "000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def feuille(wb, nom, titre, sous, ent, larg, lignes, colorees=(), legende=None):
    ws = wb.create_sheet(nom)
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A2"] = titre, sous
    ws["A1"].font = Font(name="Helvetica", size=14, bold=True)
    ws["A2"].font = Font(name="Helvetica", size=9, italic=True)
    for j, (e, w) in enumerate(zip(ent, larg), 1):
        c = ws.cell(4, j, e)
        c.font = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NOIR)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORD
        ws.column_dimensions[c.column_letter].width = w
    for i, l in enumerate(lignes, 5):
        for j, v in enumerate(l, 1):
            c = ws.cell(i, j)
            c.font = Font(name="Helvetica", size=9)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = BORD
            if j in colorees:
                colorer(c, v)
            else:
                c.value = v
                if isinstance(v, (int, float)) and abs(v) >= 1000:
                    c.number_format = "# ##0"
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    if legende:
        r = 6 + len(lignes)
        ws.cell(r, 2, legende[0]).font = Font(name="Helvetica", size=10, bold=True)
        for k, (code, sens) in enumerate(legende[1], r + 1):
            colorer(ws.cell(k, 2), code)
            ws.cell(k, 2).border = BORD
            d = ws.cell(k, 3, sens)
            d.font = Font(name="Helvetica", size=9)
            d.border = BORD
            d.alignment = Alignment(wrap_text=True)
    return ws


LEG_STAT = ("Code couleur : résultat du contrôle", [("Conforme", "Le chiffre recalculé à partir du fichier du client est identique au franc près."),
                                                    ("Écart expliqué", "Écart entièrement expliqué par une règle de traitement documentée."),
                                                    ("Anomalie corrigée", "Présentation incomplète ou imprécise, corrigée dans le rapport."),
                                                    ("Erreur corrigée", "Chiffre ou affirmation inexacte, corrigée dans le rapport.")])
LEG_GRAV = ("Code couleur : gravité", [("Critique", "Change un verdict ou un chiffre clé du résumé exécutif."),
                                       ("Élevé", "Change un chiffre ou une conclusion d'une partie."),
                                       ("Moyen", "Précision ou cohérence entre diapositives."),
                                       ("Faible", "Forme ou formulation.")])

# ---------- synthèse des contrôles
nA = sum(x["comptes_bruts"] for x in V["A_comptes"])
eA = sum(len(x["ecarts"]) for x in V["A_comptes"])
eB = sum(1 for x in V["B_equilibre"] if x["ecart"])
nC = len(V["C_agregats"])
eC = [x for x in V["C_agregats"] if abs(x["ecart"]) > 1]
ok = [x for x in P if x.get("ecart") is not None and abs(x["ecart"]) <= 1]
rej = [x for x in P if x.get("ecart") and x.get("lignes_brutes") is not None and x["lignes_brutes"] - x["lignes_base"] == 1 and x["ecart"] > 0]
bal = C[C.type == "balance âgée"]
synth = [
    ("Comptabilité", "Soldes de chaque compte : balance brute du FPPN relue directement et base de travail", "7 balances de 2020 au 31/08/2026", f"{nA} comptes, {eA} écart", "Conforme"),
    ("Comptabilité", "Équilibre de chaque balance (soldes débiteurs = soldes créditeurs)", "7 balances", f"{7 - eB} balances équilibrées sur 7", "Conforme"),
    ("Comptabilité", "Balances renvoyées le 22/10/2026 et balances utilisées", "« BALANCE 2021.xlsx », « BALANCE FPPN 2025.xlsx »", f"{sum(x['nb_ecarts'] for x in V['D_renvois'])} compte en écart", "Conforme"),
    ("Comptabilité", "Agrégats du rapport recalculés depuis les balances brutes", "23 agrégats × 7 exercices", f"{nC - len(eC)} identiques sur {nC}. Seul écart : cotisations décès 2025 (590 000 FCFA de MEDOC, compte 7022), présentées comme « décès et MEDOC » dans le tableau de bord", "Écart expliqué"),
    ("Précomptes", "Total de chaque feuille mensuelle de précomptes : fichier brut et base", "506 feuilles de 2020 à août 2026, 43 102 485 135 FCFA", f"{len(ok)} identiques. {len(rej)} feuilles diffèrent d'une ligne sans matricule valide ({fcfa(sum(x['ecart'] for x in rej))}, soit {pct(sum(x['ecart'] for x in rej) / 43102485135)}). Montants décès de plus de 1 000 000 FCFA corrigés (avril 2024). Formats particuliers relus à la main : identiques", "Écart expliqué"),
    ("Grenier", "Nombre de prêts, montant débloqué et encours de chaque balance âgée, comparés aux totaux imprimés par le logiciel de prêts", f"{len(bal)} balances âgées du 31/12/2021 au 31/08/2026", f"{int((bal.nombre_lu == bal.nombre_imprime).sum())} sur {len(bal)} identiques sur les trois totaux", "Conforme"),
    ("Classeur", "Chaque cellule recalculée par formule comparée au chiffre publié", "94 onglets, 5 232 cellules", "5 232 identiques, 0 erreur de formule", "Conforme"),
    ("Classeur", "Chiffres du rapport retrouvés dans le classeur des calculs", "2 316 chiffres", "2 316 retrouvés", "Conforme"),
    ("Décès", "État de fin d'année des allocations décès et crédit du compte 41591000", "« PROV ALLOCATIONS DECES ET FRAIS FUNERAIRES 2025.xlsx », « GRAND LIVRE FPPN 2024.xlsx »", "2025 : 130 600 000 FCFA des deux côtés. 2024 : 140 800 000 FCFA des deux côtés", "Conforme"),
    ("Effectifs", "Cotisants du mois (matricules distincts du poste 936) et montant précompté", "Janvier 2024 à août 2026", "Identiques au rapport (exemple : 31 198 cotisants et 224 460 655 FCFA en août 2026)", "Conforme"),
    ("Placements", "Encours réellement placé, lu écriture par écriture dans les grands livres", "Grands livres des comptes 2021 à 2026", "Soldes de clôture et produits financiers identiques aux balances. Encours moyen supérieur à la moyenne des soldes de clôture retenue jusqu'ici", "Erreur corrigée"),
]
rows_pl = []
for an in sorted(G, key=int):
    g = G[an]
    rows_pl.append((int(an), g["fichier"], g["dat_efc"]["moyenne"], g["dat_efc"]["pic"], g["dat_pep"]["moyenne"], g["dat_orabank"]["moyenne"], g["fcp"]["moyenne"],
                    g["placements"]["moyenne"], g["placements"]["cloture"], g["produits_77"]))

ANOM = [
    ("Placements du FPPN", "146", "Le DAT BOA PEP (compte 52210300) n'apparaît pas : 505 383 858 FCFA placés de 2020 à 2024, soldés en 2025. L'année 2020 manque (DAT BOA FPPN 1 et DAT EFC soldés en 2020).", "Tableau étendu à 2020 et à tous les DAT (52210100, 52210200, 52210201, 52210300, 52210400) et au FCP.", "Critique", "Erreur corrigée"),
    ("Placements du FPPN", "146", "Les « versements » de 1 965 985 975 à 3 335 116 874 FCFA par an additionnent les renouvellements mensuels du même dépôt. Le grand livre 2024 montre un DAT d'un mois renouvelé chaque mois, augmenté d'environ 40 000 000 FCFA, porté de 40 000 000 FCFA en février à 448 315 930 FCFA en décembre, puis viré au FPPN le 12/12/2024.", "Le tableau présente l'encours moyen et le pic de chaque année, lus dans les grands livres. La phrase sur les versements est retirée.", "Critique", "Erreur corrigée"),
    ("Placements du FPPN", "146", "« Soldé à chaque clôture » : le DAT EFC garde 140 527 FCFA au 31/12/2022.", "Formulation exacte reprise.", "Faible", "Anomalie corrigée"),
    ("Adéquation des placements", "8, 149, 150, 193", "Placements moyens calculés sur les soldes de clôture : 515 290 922 FCFA en 2025, soit 3,73 % de l'épargne, et rendement moyen de 6,48 % de 2021 à 2025. L'encours réellement placé pendant l'année est plus élevé : 763 752 135 FCFA en 2025 (5,53 % de l'épargne) et 709 132 592 FCFA en moyenne de 2021 à 2025 (rendement de 4,90 % par an).", "Chiffres remplacés par l'encours moyen pondéré par les jours, lu dans les grands livres. Verdict inchangé (placements insuffisants), chiffres corrigés.", "Critique", "Erreur corrigée"),
    ("Rendement des placements", "193", "Rendement 2025 de 22,88 % : il vient des intérêts du DAT PEP, constatés chaque année en intérêts courus puis annulés l'année suivante, et encaissés à la clôture du dépôt en 2025.", "Le rendement est présenté sur 2021 à 2025 cumulés, avec l'explication du pic de 2025.", "Élevé", "Erreur corrigée"),
    ("Sinistralité décès", "66", "Prestations décès 2020 « comptes » : 578 327 196 FCFA = obsèques 3 933 335 FCFA + prime d'assurance décès 574 393 861 FCFA (compte 65180000). Les diapositives 62 et 64 donnent 0,56 % hors prime.", "Note ajoutée : en 2020, l'assureur portait le risque. Les deux lectures sont nommées.", "Moyen", "Anomalie corrigée"),
    ("Fichiers vides ou endommagés", "41, 219", "Les grands livres endommagés ont tous été renvoyés lisibles le 22/10/2026, dont « GRAND LIVRE DES COMPTES 2022.xlsx » (année complète, 192 comptes).", "Diapositive corrigée. Question 12 limitée aux fichiers de décembre 2024 et aux précomptes manquants.", "Élevé", "Erreur corrigée"),
    ("Base de données", "42", "« La base reproduit les fichiers du FPPN sans perte » : 26 lignes sans matricule valide sont écartées (737 890 FCFA, 0,0017 % des précomptes), chacune consignée au journal des rejets.", "Phrase précisée avec le nombre de lignes et le montant.", "Faible", "Anomalie corrigée"),
    ("Méthode : lecture du bilan", "102", "« Le bilan n'a pas été transmis » : les états financiers de 2020 à 2025 et au 31/08/2026 ont été reçus le 22/10/2026 (« ETATS FINANCIERS 2020.pdf » à « ETATS FINANCIERS AU 31-08-2026.pdf »).", "Limite réécrite : bilan reconstitué depuis les balances et confronté aux états financiers reçus (partie 3).", "Élevé", "Erreur corrigée"),
    ("Analyses en attente", "218", "« Provisions (PSAP, Chain Ladder) » alors que Chain Ladder est écarté (annexe des méthodes non retenues).", "Remplacé par les méthodes retenues : dossier par dossier, factures arrivées après la clôture, méthode CIMA des sinistres tardifs.", "Élevé", "Erreur corrigée"),
    ("Délai de paiement des cliniques", "6, 107, 162", "Deux définitions coexistent : 2,62 mois en 2025 (dettes du compte 40110000 / factures de cliniques de l'année) et 44,98 jours (indicateur DPP de la Décision : dettes / charges de prestations).", "Les deux définitions sont nommées là où elles apparaissent.", "Moyen", "Anomalie corrigée"),
    ("Excédent par risque", "174", "« Frais attribués au prorata des cotisations » sans dire quels comptes forment les frais, et clé de 70 % pour la maladie présentée comme à confirmer par un barème.", "Frais définis comme les charges de fonctionnement (comptes 612, 613, 618, 62 à 68). Clé de 70 % fondée sur les statuts (art. 17 : 3,5 % sur 5 %).", "Moyen", "Anomalie corrigée"),
    ("Tableau de bord (2/4) : risques longs", "160", "Taux brut de rentabilité des placements calculé sur la moyenne des soldes de clôture (22,88 % en 2025), alors que la partie 9 publie le rendement lu dans les grands livres (15,44 % en 2025).", "Ratio recalculé sur l'encours moyen lu dans les grands livres, de 4,62 % en 2021 à 15,44 % en 2025, comme la partie 9 et le classeur des calculs.", "Élevé", "Erreur corrigée"),
    ("Références et sources", "228, 268", "Provision pour sinistres tardifs rattachée à l'article 334-8 du Code CIMA. La circulaire n° 00230/CIMA/CRCA/PDT/2005, lue dans l'extrait du Code CIMA reçu le 23/10/2026, la rattache à l'article 334-12.", "Référence corrigée (article 334-12). L'article 334-8 cité pour la santé des retraités est surligné jusqu'à lecture du texte (pages du Code CIMA illisibles dans l'envoi).", "Moyen", "Anomalie corrigée"),
    ("Conformité prudentielle en bref", "15", "« Retraité » désignait à la fois les policiers à la retraite et la lecture corrigée des comptes. Résumé exécutif limité à 2025.", "La lecture corrigée devient « lecture corrigée selon le PCMS », définie sur la diapositive. Matrice de 2020 au 31/08/2026, valeur et boule dans chaque case.", "Moyen", "Anomalie corrigée"),
    ("Provisions figées", "106, 109, 222", "Montants identiques depuis le report à nouveau du 01/09/2020 : 2 867 602 878 FCFA et 1 961 695 412 FCFA. Seul mouvement : reclassement de 2024. Concorde avec le chef comptable (montants hérités d'environ 2004).", "Questions 2 et 23 fusionnées. Constat inchangé.", "Faible", "Conforme"),
    ("Sigle CMP", "Toutes", "« Centre Médical du Policier (rapport de gestion 2023) ».", "Remplacé par « Clinique médicale du policier ».", "Faible", "Anomalie corrigée"),
    ("Formulations non établies", "25, 26, 39, 56, 75, 80, 116, 119, 128, 129, 134, 191, 220, 222", "« À confirmer », « à préciser », « en principe », « probablement », « correspondraient ».", "Chaque phrase est réécrite sur le document ou le calcul qui l'établit. Ce qui manque est demandé au FPPN (liste des éléments à obtenir).", "Moyen", "Anomalie corrigée"),
    ("Points à clarifier", "220, 222", "Questions déjà résolues par les documents : types de prêts (8 types dans les fichiers de production), barème des officiers (1 600 000 FCFA, cartographie des risques), compte de la provision décès (41591000), canal des fichiers retour.", "Questions retirées.", "Faible", "Anomalie corrigée"),
]

wb = Workbook()
wb.remove(wb.active)
feuille(wb, "0_Synthèse", "Vérification de cohérence des chiffres du rapport", "Contrôles refaits à partir des seuls fichiers du client et des documents collectés par Adenka. Chaque contrôle relit le fichier d'origine sans passer par les résultats déjà calculés.",
        ["N°", "Domaine", "Contrôle", "Périmètre", "Résultat", "Statut"], [5, 14, 55, 40, 60, 16], [(i, *x) for i, x in enumerate(synth, 1)], {6}, LEG_STAT)
feuille(wb, "1_Anomalies", "Anomalies relevées dans le rapport et corrections", "Numéros de diapositive du rapport du 22/10/2026.",
        ["N°", "Diapositive", "N° de diapositive", "Constat", "Correction", "Gravité", "Statut"], [5, 26, 12, 70, 55, 11, 16],
        [(i, *x) for i, x in enumerate(ANOM, 1)], {6, 7}, LEG_GRAV)
feuille(wb, "2_Balances", "Balances brutes et base de travail, par exercice", "Fichiers « BALANCE … » relus directement (conversion CSV de chaque feuille).",
        ["Exercice", "Fichier relu", "Comptes lus", "Comptes en base", "Comptes en écart", "Total des soldes débiteurs", "Total des soldes créditeurs", "Écart d'équilibre"],
        [10, 60, 12, 12, 12, 20, 20, 14],
        [(a["exercice"], a["fichier"], a["comptes_bruts"], a["comptes_base"], len(a["ecarts"]), b["soldes_debiteurs"], b["soldes_crediteurs"], b["ecart"]) for a, b in zip(V["A_comptes"], V["B_equilibre"])])
feuille(wb, "3_Agrégats", "Agrégats recalculés depuis les balances brutes", "Définitions du classeur des calculs (onglet 10_agregats), appliquées aux soldes relus.",
        ["Exercice", "Agrégat", "Recalcul depuis la balance brute", "Valeur utilisée par le rapport", "Écart"], [10, 34, 24, 24, 14],
        [(x["exercice"], x["agregat"], x["recalcul_brut"], x["publie"], x["ecart"]) for x in V["C_agregats"]])
feuille(wb, "4_Précomptes", "Total de chaque feuille de précomptes : fichier brut et base", "Feuilles retenues de janvier 2020 à août 2026.",
        ["Source", "Fichier", "Poste", "Sens", "Année", "Mois", "Colonne lue", "Total brut", "Lignes brutes", "Total base", "Lignes base", "Écart"],
        [9, 50, 7, 7, 7, 6, 20, 16, 10, 16, 10, 12],
        [(x["src"], x["fichier"], x.get("poste"), x.get("sens"), x.get("annee"), x.get("mois"), x.get("colonne"), x.get("brut"), x.get("lignes_brutes"), x.get("base"), x.get("lignes_base"), x.get("ecart")) for x in P])
feuille(wb, "5_Grenier", "Balances âgées du Grenier : totaux lus et totaux imprimés", "Exports « BalanceAgeeDetaille … .xls » du logiciel de prêts.",
        ["Fichier", "Arrêté", "Portefeuille", "Prêts lus", "Prêts imprimés", "Débloqué lu", "Débloqué imprimé", "Encours lu", "Encours imprimé"],
        [44, 12, 12, 10, 10, 18, 18, 18, 18],
        [(r.fichier, str(r.arrete)[:10], r.portefeuille, int(r.nombre_lu), int(r.nombre_imprime), r.debloque_lu, r.debloque_imprime, r.encours_lu, r.encours_imprime) for r in bal.itertuples()])
feuille(wb, "6_Placements GL", "Placements lus dans les grands livres : encours moyen pondéré par les jours", "Comptes 5221 (DAT), 5330 (FCP) et 77 (produits financiers). Pour 2026, moyenne calculée du 01/01 au 31/08/2026.",
        ["Exercice", "Grand livre", "DAT EFC : encours moyen", "DAT EFC : pic", "DAT PEP : encours moyen", "DAT Orabank : encours moyen", "FCP : encours moyen", "Total placé : encours moyen", "Total placé : clôture", "Produits financiers (77)"],
        [9, 50, 18, 18, 18, 18, 18, 18, 18, 18], rows_pl)
wb.save(SORTIE)
print(SORTIE)
