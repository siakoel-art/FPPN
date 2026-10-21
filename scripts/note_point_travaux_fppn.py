# -*- coding: utf-8 -*-
"""
FPPN - Note de point d'étape sur les travaux actuariels (réunion du 21/10/2026)
=============================================================================
Trois tableaux : (1) constats du résumé exécutif du rapport v8, avec les données, la méthode, les raisons et les
fichiers sources exacts ; (2) limites au regard des tâches de la mission ; (3) documents prioritaires.
Chiffres repris du rapport v8 et du classeur v8 (contrôle automatique : tous les chiffres retrouvés).
Utilisation : python note_point_travaux_fppn.py <dossier_skills> <sortie.docx>
"""
import os
import sys

SKILLS, SORTIE = sys.argv[1:3]
sys.path.insert(0, os.path.join(SKILLS, "adenka-rapport-word", "scripts"))
from adenka_docx import RapportAdenka  # noqa: E402

BAL20 = "« BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx » (AUDIT ADENKA\\DAFC\\2020)"
BAL21 = "« BALANCE FPPN FORMAT EXCEL DU 01-01-2021 AU 31-12-2026.xlsx » (AUDIT ADENKA\\DAFC\\2021)"
BAL25 = "« BALANCE FPPN 2025.xlsx » (AUDIT ADENKA\\DAFC\\2025)"
BAL26 = "« BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx » et « BALANCE DES TIERS FPPN DU 01-01-2026 AU 31-08-2026.xlsx » (AUDIT ADENKA\\DAFC\\2026)"
BALS = ("Les sept balances : « BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx », « BALANCE FPPN FORMAT EXCEL DU 01-01-2021 AU 31-12-2026.xlsx », "
        "« BALANCE FPPN 2022.xlsx », « BALANCE FPPN 2023.xlsx », « BALANCE FPPN 2024.xlsx », « BALANCE FPPN 2025.xlsx », « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx » (AUDIT ADENKA\\DAFC)")
TEXTES = ("Textes de l'AIRMS (dossiers AIRMS_01 et AIRMS_02 fournis le 18/10/2026) : Décision n° 05/2022/COM/UEMOA, règlement d'exécution n° 004/2022/COM/UEMOA, "
          "« Règlement n°07-2009-CM-UEMOA_Mutualité sociale (1) (3).pdf », « Reglement_03_2012_CM_UEMOA-adoption plan comptable (1).pdf »")

def sans_pv(lignes, col_sources=None):
    """Charte : pas de point-virgule dans les phrases. Sources : un fichier par ligne ; texte : phrases séparées."""
    out = []
    for ligne in lignes:
        nl = []
        for j, v in enumerate(ligne):
            if j == col_sources:
                v = v.replace(" ; ", "\n")
            else:
                morceaux = v.split(" ; ")
                v = ". ".join(m[0].upper() + m[1:] if k and m else m for k, m in enumerate(morceaux))
            nl.append(v)
        out.append(nl)
    return out


r = RapportAdenka(titre="Point d'étape des travaux actuariels", client="FPPN", type_livrable="Note de réunion", date="Octobre 2026")
r.sigles({"AIRMS": "Agence ivoirienne de régulation de la mutualité sociale", "BCEAO": "Banque centrale des États de l'Afrique de l'Ouest",
          "CGRAE": "Caisse générale de retraite des agents de l'État", "CIMA": "Conférence interafricaine des marchés d'assurances",
          "CMP": "clinique du FPPN (sigle des comptes, à confirmer)", "DAT": "dépôt à terme", "EFC": "Entraide de Fin de Carrière",
          "FCP": "fonds commun de placement", "FPPN": "Fonds de prévoyance de la Police nationale", "PBC": "liste des documents demandés au FPPN",
          "PCMS": "plan comptable des mutuelles sociales", "PEP": "Plan Épargne Plus", "PSAP": "provision pour sinistres à payer",
          "S/C": "prestations / cotisations", "UEMOA": "Union économique et monétaire ouest-africaine"})
r.abreviations({"FCFA": "franc de la Communauté financière africaine", "art.": "article"})

r.synthese("OBJET DE LA NOTE")
r.para("**Cette note accompagne le résumé exécutif du rapport actuariel et statistique (version 8) pour le point d'étape du 21/10/2026.** Elle répond, pour chaque constat, "
       "à quatre questions : à partir de quelles données, par quelle méthode, pour quelles raisons, et avec quels fichiers sources exacts.")
r.para("**Elle présente ensuite les limites des conclusions au regard des quatre tâches de la mission, puis les documents prioritaires.** Période étudiée : du 01/01/2020 au "
       "31/08/2026. Les tâches sont l'évaluation des engagements envers les adhérents et leurs ayants droit, la suffisance des provisions, l'adossement actif-passif et "
       "l'adéquation des placements aux engagements.")
r.para("**Tous les chiffres viennent du rapport v8 et du classeur des calculs v8, où chaque chiffre est retracé.** Les fichiers cités sont ceux de l'extraction "
       "« FPPN_extraction_20261018_1939.zip » (4 287 fichiers, journal « journal_fichiers.csv »), avec leur dossier d'origine.")

r.partie(1, "Constats, méthodes et sources", introduction="Un constat du résumé exécutif par ligne : la donnée de départ, le calcul, les raisons et les fichiers qui le prouvent.")
r.h2("1.1. POURQUOI CHAQUE CONSTAT")
LIGNES1 = [
    ["**Données incomplètes** : 4 454 416 lignes de précomptes contrôlées ; précomptes mensuels à partir de 2023 ou 2024 selon le poste ; aucune base de prestations",
     "Fichiers aller et retour de la Solde, de la CGRAE et de GNA ; inventaire des 4 287 fichiers reçus",
     "Lecture de chaque fichier et de chaque feuille ; contrôle des totaux, des doublons et des lignes écartées ; mois couverts par poste",
     "Précomptes de 2020 à 2022 non transmis (sauf PEP 2001 à 2022) ; prestations jamais extraites du logiciel",
     "« FPPN_extraction_20261018_1939.zip » ; « journal_fichiers.csv » ; dossiers FPPN-SOLDE, FPPN-CGRAE\\PRECOMTES et FPPN-GNA"],
    ["**Cotisations** : 5,00 % du salaire de base, répartis par la Solde en 3,77 % maladie et 1,23 % décès, au lieu de 3,5 % et 1,5 %",
     "Fichiers retour maladie (poste 936) et décès (poste 8Z6) avec le salaire de base de chaque agent : 209 670 lignes",
     "Pour chaque agent : précompte / salaire de base ; valeur médiane par poste ; comparaison au barème des statuts",
     "La Solde applique une répartition différente de l'article 17 des statuts ; la comptabilité suit une troisième clé",
     "FPPN-SOLDE\\EFFECTIF POLICE (MALADIE)\\2024 à 2026 (ex. « 012024_effectif_police_936_JANVIER 2024.xlsx ») ; FPPN-SOLDE\\ALLOCATION DECES\\2024 à 2026 (ex. « FPPN_ALLOC_DECES_8Z6_JANVIER_2025.xlsx ») ; « STATUTS FPPN 15 JUIN 2023.pdf », art. 17"],
    ["**Sinistralité maladie** : S/C de 6,00 % en 2021 et 73,53 % en 2025",
     "Prestations maladie (comptes 601, 6013, 6018, 60800007) et cotisations maladie (compte 7010)",
     "S/C = prestations / cotisations : 101 453 171 / 1 691 332 910 en 2021 ; 1 771 635 361 / 2 409 281 030 en 2025",
     "Prestations multipliées par 17,46 en quatre ans, cotisations en hausse de 42,45 % ; provisions inchangées, donc sans effet sur le ratio",
     BAL21 + " ; " + BAL25],
    ["**Provisions figées** : 4 829 298 290 FCFA de 2020 au 31/08/2026",
     "Soldes des provisions maladie (19011000 puis 15801000 : 2 867 602 878 FCFA) et obsèques (19012000 puis 15802000 : 1 961 695 412 FCFA) ; absence des comptes 6901 et 7901",
     "Comparaison des soldes d'une clôture à l'autre ; recherche des dotations et reprises de provisions",
     "Montants hérités d'environ 2004, jamais recalculés ; aucune note de calcul transmise",
     BALS + " ; « FPPN_CR_entretien_provisions_20261019.pptx »"],
    ["**Épargne** : 15 743 516 391 FCFA due au 31/08/2026, couverte à 16,28 % par des actifs dédiés",
     "Comptes 41195100 (PEP : 5 229 245 425 FCFA) et 41195200 (EFC : 10 514 270 966 FCFA) ; banques, DAT dédiés et FCP : 2 562 885 661 FCFA",
     "Couverture = actifs dédiés / épargne due = 2 562 885 661 / 15 743 516 391",
     "Épargne employée dans la créance sur l'hôpital, la clinique, les terrains et les prêts du Grenier ; DAT soldé à chaque clôture",
     BAL26],
    ["**Fonds propres** : -438 242 200 FCFA fin 2025",
     "Capitaux propres (classes 10 à 14) dont pertes reportées (compte 129) : 5 272 983 928 FCFA ; résultat 2025 : 437 899 911 FCFA",
     "Fonds propres = capitaux propres (-876 142 111 FCFA) + résultat de l'exercice (437 899 911 FCFA)",
     "Pertes passées supérieures aux réserves ; les excédents annuels réduisent l'écart sans le combler",
     BAL25],
    ["**Conformité** : sur 36 ratios, 14 non conformes, 7 conformes, 15 en attente ; 13 ratios sur 14 non conformes chaque année de 2020 à 2026",
     "Agrégats des balances 2020 au 31/08/2026 ; formules et normes des textes UEMOA",
     "Chaque ratio recalculé avec la formule et la norme de la Décision n° 05/2022 ; statut jugé sur 2022 à 2025 ; montant à combler par inversion de la formule",
     "Cinq causes récurrentes : épargne non adossée, pertes reportées, réserve obligatoire tardive, frais au-delà des plafonds, prestations faibles",
     BALS + " ; " + TEXTES + " ; « FPPN_Calculs_rapport_20261015_v8.xlsx » (onglets 10 et 13)"],
    ["**Grenier** : 1 044 302 107 FCFA d'encours au 31/08/2026, dont 915 477 260 FCFA financés par le FPPN",
     "Balances âgées détaillées du logiciel de prêts (2 959 prêts) ; compte 47111102",
     "Somme des encours prêt par prêt ; comparaison au financement comptabilisé par le FPPN",
     "Production soutenue de prêts, financée par l'épargne et les excédents du FPPN",
     "AUDIT ADENKA\\SERVICE GRENIER DU POLICIER\\FICHIER ALLER\\GRENIER AOÛT 2026 : « BalanceAgeeDetaille AOUT 2026.xls » et ses versions CRF, CRO, CRS ; « production_de_la_periode … .xls » ; " + BAL26],
    ["**Adossement** : les engagements connus (22 761 252 888 FCFA) sont couverts à 96,33 % au mieux, à 27,19 % sans les actifs de valeur non établie",
     "Actif par délai de mobilisation et passif au 31/08/2026 ; créance sur l'hôpital, clinique, terrains",
     "Quatre scénarios de valorisation tirés des textes ; couverture = actif retenu / engagements connus",
     "Actifs longs et non évalués ; toutes les corrections attendues aggravent l'écart : le constat est un minimum",
     BAL26 + " ; règlement d'exécution n° 004/2022, art. 17, 19, 20, 21 et 27"],
    ["**Placements** : liquidité et diversification insuffisantes ; rendement de 1,30 %, 1,28 % et 1,73 % de 2022 à 2024, sous l'État de Côte d'Ivoire à 5 ans (5,19 %, 6,86 %, 7,73 %)",
     "DAT BOA (compte 52210200), FCP (53300001), produits financiers (classe 77) ; taux du marché",
     "Rendement = produits financiers / placements moyens ; comparaison aux taux BCEAO et aux cotations des obligations de l'État",
     "Placements faibles en volume, débouclés chaque année ; pas de politique de placement écrite",
     BALS + " ; « FPPN_CR_entretien_placements_20261019.pptx » ; « RAPPORT-SUR-LA-POLITIQUE-MONETAIRE-DANS-L-UMOA-DE-MARS-2026.pdf » ; « Rendement de l'Obligation Cote d'Ivoire 5-Year - Données Historiques.csv » et « … 10-Year … »"],
    ["**Leviers de conformité** : les frais de fonctionnement devraient baisser de 558 364 763 FCFA en 2025 ; les normes de flux ne peuvent pas être respectées ensemble",
     "Numérateurs et dénominateurs des ratios ; statuts du FPPN (art. 17, 85, 90, 93)",
     "Montant d'équilibre de chaque levier, une hypothèse à la fois ; démonstration algébrique sur les formules de la Décision",
     "Le taux de cotisation améliore les frais mais éloigne la sinistralité de sa norme : les leviers se contredisent",
     "« FPPN_Calculs_rapport_20261015_v8.xlsx » (onglets 13a à 13i) ; « STATUTS FPPN 15 JUIN 2023.pdf » ; « REGLEMENT INTERIEUR FPPN 15 JUIN 2023.pdf »"],
]
r.tableau(["Constat et chiffre", "À partir de quoi", "Comment", "Principales raisons", "Sources exactes"], sans_pv(LIGNES1, 4),
          titre="Constats du résumé exécutif : données, méthode, raisons et fichiers sources",
          source="Rapport actuariel et statistique v8 (diapositives 4 à 14, parties 4 à 10) ; classeur des calculs v8 ; analyses Adenka.",
          largeurs_cm=[3.4, 3.5, 3.7, 3.5, 4.3], taille=7.5, premiere_col_grise=False)

r.partie(2, "Limites au regard des tâches de la mission", introduction="Ce qui est établi, ce qui ne l'est pas encore, et pourquoi.")
r.h2("2.1. LIMITES PAR TÂCHE")
LIGNES2 = [
    ["1. Engagements envers les adhérents et ayants droit", "Épargne due : 15 743 516 391 FCFA au 31/08/2026",
     "Provisions mathématiques (EFC, santé des retraités), sinistres tardifs, allocations décès restant à payer",
     "Pas de base de prestations, pas de dates de naissance, cahier des charges et règlements de l'EFC et du PEP non transmis", "Hausse des engagements"],
    ["2. Suffisance des provisions", "Les provisions sont des montants hérités, sans calcul",
     "Montant juste des provisions, donc leur suffisance face aux prestations payées",
     "Pas de base de prestations avec dates de survenance, de déclaration et de paiement ; pas de note de calcul", "Inconnu"],
    ["3. Adossement actif-passif", "Non-adossement démontré, même dans le cas le plus favorable",
     "Ampleur réelle de l'écart ; adossement en flux de trésorerie futurs",
     "Échéancier de l'épargne absent ; actifs non évalués ; provisions à recalculer", "Aggravation"],
    ["4. Adéquation des placements", "Liquidité et diversification insuffisantes ; rendement sous le marché de 2022 à 2024",
     "Sécurité des placements et conformité aux conditions fixées en Côte d'Ivoire",
     "Décision ministérielle, relevés du DAT et du FCP, délibérations du Conseil d'administration non transmis", "Stable"],
    ["Analyses statistiques complémentaires", "Effectifs, cotisations, sinistralité par exercice comptable, Grenier",
     "Mortalité observée, sinistralité par année de survenance, délais de paiement",
     "Dates de naissance et bases de prestations non transmises", "À établir"],
    ["Provision du personnel du FPPN", "Méthode et barème connus (décret n° 2017-210)",
     "Montant recalculé de la provision", "Fichier du personnel avec le statut civil ou policier non transmis", "Faible montant"],
    ["Toutes les tâches", "Résultats sur les balances 2020 à 2025",
     "Valeurs définitives au 31/08/2026", "Balance 2026 provisoire (huit mois) ; précomptes 2020 à 2022 absents", "À confirmer"],
]
r.tableau(["Tâche", "Ce qui est établi", "Ce qui ne l'est pas", "Pourquoi", "Sens attendu"], sans_pv(LIGNES2),
          titre="Limites des conclusions au regard des tâches de la mission",
          source="Rapport v8, réponse aux quatre questions de la mission (diapositive 5) et partie 14 ; PBC pointée.",
          largeurs_cm=[3.6, 3.8, 3.9, 4.6, 2.5], taille=8)

r.partie(3, "Documents prioritaires", introduction="Les documents à obtenir du FPPN, par ordre de priorité, avec la tâche qu'ils débloquent.")
r.h2("3.1. DOCUMENTS NÉCESSAIRES")
LIGNES3 = [
    ["71, 13, 28, 40", "Bases des prestations maladie et décès depuis 2019, une ligne par dossier, avec dates de survenance, de déclaration et de paiement", "1, 2, statistiques", "1",
     "Calcul de la PSAP dossier par dossier et des sinistres tardifs (circulaire CIMA n° 00230/2005)"],
    ["72", "Note de calcul des provisions maladie et obsèques, et liste des dossiers", "2", "1", "Justifier ou reprendre les montants hérités"],
    ["92, 93", "Fichiers des décès, calcul de fin d'année des allocations restant à payer, barème officiel des allocations décès", "1, 2", "1", "Allocations décès restant dues"],
    ["100, 101", "Cahier des charges du régime des prestations sociales et tableau des allocations", "1, 2", "1", "Montants des prestations auxquels renvoient les statuts"],
    ["57, 102, 16, 17, 99", "Règlements et manuels de l'EFC et du PEP, comptes individuels d'épargne, décision fixant le taux servi de 3,5 %", "1, 3", "1", "Valeur et échéancier de l'épargne"],
    ["9", "Fichier des mutualistes avec date de naissance et sexe", "1, 3, statistiques", "1", "Mortalité, départs à la retraite, projection des engagements"],
    ["79, 108", "Règles de la garantie santé des retraités et situation des policiers retraités", "1", "2", "Provision mathématique éventuelle"],
    ["84, 86, 87, 88", "Titres et évaluation des terrains, réévaluation du siège, dossier de la créance sur l'hôpital, bilan de la clinique", "3", "2", "Valeur de marché des actifs (règlement d'exécution n° 004/2022, art. 20)"],
    ["54, 80, 81, 82, 105", "Décision ministérielle sur les placements, relevés du DAT et du FCP, inventaire des placements, délibérations du Conseil d'administration", "4", "2", "Sécurité, rendement et autorisation des placements"],
    ["76, 78, 103, 104", "Fichier du personnel avec statut, note de calcul et convention appliquée", "Provision du personnel", "3", "Recalcul de la provision pour départ à la retraite"],
]
r.tableau(["N° PBC", "Document", "Tâches débloquées", "Priorité", "Pourquoi"], sans_pv(LIGNES3),
          titre="Documents prioritaires et nécessaires à la réalisation des tâches",
          source="« FPPN_PBC_complete_pointee_20261019.xlsx » (items 1 à 109). Priorité 1 : bloque une conclusion. Priorité 2 : précise une conclusion. Priorité 3 : complément.",
          largeurs_cm=[2.4, 6.6, 2.6, 1.6, 5.2], taille=8)
r.para("**Les items 71, 72, 92 et 57 débloquent à la fois les engagements, les provisions et l'adossement.** Leur transmission conditionne la version définitive du rapport.")
r.enregistrer(SORTIE)
print("note écrite :", SORTIE)
