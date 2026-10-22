# -*- coding: utf-8 -*-
"""Domaine de validité exact de chaque diapositive du rapport (titre -> texte de la ligne « Valable pour »).

Chaque texte précise le fichier (liste complète en annexe « Fichiers sources »), le compte, la feuille ou la colonne
utilisés, la période, la population couverte et ce que le résultat ne couvre pas.
"""
BAL = "balances « BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx » à « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx »"
P936 = "fichiers retour du poste 936 (dossier « FPPN-SOLDE\\EFFECTIF POLICE (MALADIE) »)"
P8Z6 = "fichiers retour du poste 8Z6 (dossier « FPPN-SOLDE\\ALLOCATION DECES »)"
PCG = "fichiers retour de la CGRAE (dossier « FPPN-CGRAE\\PRECOMTES\\FICHIER RETOUR (CGRAE) »)"
GRN = "exports « BalanceAgeeDetaille … .xls » du logiciel de prêts du Grenier"
PROD = "exports « production_de_la_periode … .xls » du logiciel de prêts du Grenier"
SENS = "classeur des calculs, onglets 13a à 13i"
RAT = "ratios de la Décision n° 05/2022 recalculés sur les " + BAL

VAL = {
    # ---------------------------------------------------------------- résumé exécutif
    "Résumé exécutif": "chiffres repris des parties 4 à 10, chacun dans le domaine indiqué sur sa diapositive de détail ; situation au 31/08/2026 (provisoire) sauf mention d'une autre date",
    "Graphiques clés (1/2)": "S/C maladie : comptes 601, 6013, 6018 et 60800007 (prestations) et 7010 (cotisations) des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026 ; épargne : comptes 41195100 et 41195200 ; actifs réservés : banques et DAT de l'épargne et FCP",
    "Graphiques clés (2/2)": "fonds propres : comptes 10, 11, 12 et résultat des " + BAL + " ; encours du Grenier : " + GRN + ", arrêtés annuels du 31/12/2021 au 31/12/2025 et au 31/08/2026",
    "Conformité prudentielle en bref": "exercice 2025 seulement, " + RAT + " ; ratios non calculés faute de données signalés comme tels",
    "Recommandations prioritaires": "constats au 22/10/2026 ; numéros de la liste des documents demandés (PBC) pointée au 22/10/2026",
    # ---------------------------------------------------------------- partie 1
    "Cadre juridique": "textes cités dans leur version fournie le 18/10/2026 (documentation publique de l'AIRMS) ; répertoire AIRMS du 15/10/2026 ; ne vaut pas avis juridique",
    "Statuts et règlement intérieur : cotisations, provisions, prestations": "statuts et règlement intérieur du FPPN édités le 15/06/2023 (adoptés le 30/06/2021), articles cités ; leur application n'est testée que là où le rapport le dit",
    "Statuts et règlement intérieur : épargne, personnel, organes": "statuts et règlement intérieur du FPPN édités le 15/06/2023, articles cités ; décrets n° 2017-210 et n° 2012-652 ; application non testée sauf mention",
    "Garanties et produits": "montants précomptés par poste dans les fichiers retour de la Solde et de la CGRAE, mois reçus entre 2023 et septembre 2026 selon le poste (diapositive « Couverture de la période »)",
    "Circuit des fichiers par produit": "circuit déclaré par le responsable des cotisations le 19/10/2026 ; non confirmé par une procédure écrite",
    "Flux financiers 2025": "exercice 2025, classes 6 et 7 de « BALANCE FPPN 2025.xlsx » ; flux d'épargne et du Grenier exclus du résultat",
    "Notions de base": "définitions de la Décision n° 05/2022 (annexe) et du règlement d'exécution n° 004/2022 ; sans chiffre propre au FPPN",
    "Mission et périmètre": "périmètre fixé par le dossier de préparation du 13/10/2026 et la demande de M. Wabi ; période du 01/01/2020 au 31/08/2026",
    # ---------------------------------------------------------------- partie 3
    "Couverture de la période": "fichiers de précomptes reçus le 15/10/2026, par poste et par mois, de janvier 2020 à août 2026 (80 mois) ; un mois est « reçu » si au moins un fichier lisible le couvre",
    "Rapprochement aller et retour": "mois pour lesquels le fichier aller et le fichier retour d'un même poste sont tous deux reçus ; montants totaux du mois, sans rapprochement agent par agent",
    "Fichiers vides ou endommagés": "fichiers reçus le 15/10/2026 et le 22/10/2026 dont la lecture a échoué ou qui ne contiennent aucune ligne de données ; liste arrêtée au 22/10/2026",
    # ---------------------------------------------------------------- partie 4
    "Effectifs cotisants": "agents précomptés chaque mois dans les " + P936 + ", janvier 2024 à septembre 2026 ; décembre 2024 vide ; hors retraités et civils de la Direction Générale",
    "Entrées et sorties": "matricules présents un mois et absents le suivant (ou l'inverse) dans les " + P936 + ", janvier 2024 à septembre 2026 ; ne distingue pas retraite, décès et radiation",
    "Structure des effectifs": "ancienneté de service calculée sur la date d'entrée des " + P936 + " de janvier 2025 ; âge non disponible",
    "Taux de cotisation": "taux maladie : colonne du salaire de base des " + P936 + " (huit mois détaillés) ; taux décès : " + P8Z6 + " rapportés au 936 de mars 2024 à septembre 2026",
    "Sensibilité : clé de répartition": "exercice 2025 ; cotisations des comptes 7010 et 7021 de « BALANCE FPPN 2025.xlsx » réparties entre maladie et décès selon la clé testée",
    "Cotisations en termes réels": "cotisations des comptes 7010 et 7021 des " + BAL + ", 2020 à 2025 ; inflation de la Banque mondiale (FP.CPI.TOTL.ZG, Côte d'Ivoire), francs de 2020",
    "Cotisations comptabilisées": "comptes 7010 (maladie) et 7021 (décès) des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026 ; 2020 = à-nouveaux au 01/09/2020 plus mouvements du 01/09 au 31/12/2020",
    "Règles des produits": "déclarations du 19/10/2026 confrontées aux montants des fichiers de précomptes reçus le 15/10/2026, postes 870, 998, 9X4, 9W3 et 8Z7, mois reçus",
    # ---------------------------------------------------------------- partie 5
    "Sinistralité comptable": "comptes 601, 6013, 6018, 60800007, 6021 et 6024 (prestations) et 7010, 7021 (cotisations) des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026 ; sans provisions ni sinistres tardifs",
    "Choix de la formule : sinistralité": "formules de la Décision n° 05/2022 (annexe, § 1.1.1 et § 1.2.1.2.5) appliquées aux comptes de prestations (601 à 6024), de cotisations (7010, 7021) et de provisions (158) des " + BAL,
    "Sinistralité selon la formule": "comptes de prestations, de cotisations (7010, 7021) et de provisions (158) des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026 ; variation de provision nulle car les provisions ne bougent pas",
    "Nature des risques": "garanties décrites par les statuts du 15/06/2023 et le règlement d'exécution n° 004/2022, art. 8 à 15 ; qualification sans chiffre",
    "Croissance des prestations": "taux annuels moyens des comptes de prestations et de cotisations (7010, 7021) des " + BAL + " et des fichiers de précomptes ; périodes indiquées ligne par ligne",
    "Prestations maladie par nature": "comptes 601, 6013, 6018 et 60800007 des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026 ; nature telle que libellée par le FPPN",
    "Frais de gestion et résultat": "comptes 61 à 66 hors 6518 (frais) et résultat net des " + BAL + ", 2020 à 2025 et du 01/01 au 31/08/2026",
    "Choix de la formule : frais de gestion": "formules de la Décision n° 05/2022 (annexe, tableau 1) appliquées aux comptes 61 à 66 hors 6518 et 7010, 7021 des " + BAL,
    # ---------------------------------------------------------------- partie 6
    "Adossement de l'épargne": "comptes 41195100 et 41195200 (épargne) et comptes bancaires, DAT et FCP réservés à l'épargne des " + BAL + ", au 31/12 de 2020 à 2025 et au 31/08/2026",
    "Versements d'épargne": "versements des postes 998 (PEP) et 870 (EFC) dans les fichiers retour de la Solde, mois reçus de 2023 à 2026, comparés aux crédits des comptes 41195100 et 41195200",
    "Sensibilité : retraits d'épargne": "épargne due et actifs à moins d'un an au 31/08/2026 (« BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx ») ; taux de retrait hypothétiques",
    "Poste 998 : autres références": "références du poste 998 dans les fichiers retour « FPPN-SOLDE\\PEP\\FICHIER RETOUR », de 2020 à août 2026 ; 2001 à 2019 non retenus",
    # ---------------------------------------------------------------- partie 7
    "Encours du Grenier": "encours des " + GRN + " aux arrêtés du 31/12/2021 au 31/12/2025 et mensuels de 2026 ; financement : compte 47111102 des " + BAL,
    "Production de prêts": "montants débloqués des " + PROD + ", 2021 à septembre 2026, par type de crédit ; prêts non rapprochés des contrats",
    "Retards de remboursement": "colonne « retard » des " + GRN + ", arrêtés du 31/12/2021 au 31/08/2026 ; retard compté sur l'échéance la plus ancienne impayée",
    "Échéances et précomptes": "échéances des prêts en cours des " + GRN + " et précomptes du poste 9X4 (« FPPN-SOLDE\\GRENIER\\FICHIER ALLER RETOUR »), mois de 2026 reçus",
    "Sensibilité : pertes sur prêts": "encours au 31/08/2026 des " + GRN + " ; taux de perte hypothétiques (la politique de crédit est appliquée en partie 9)",
    # ---------------------------------------------------------------- partie 8
    "Retraités CGRAE": "retraités précomptés dans les " + PCG + ", février 2024 à septembre 2026 ; ne couvre pas les retraités sans quotité cessible",
    "Produits annexes": "postes 9W3 (GNA) et 8Z7 (MEDOC) des fichiers « FPPN-GNA » et « FPPN-SOLDE\\MEDOC », mois reçus de 2023 à 2026 ; compte 71800000 des " + BAL,
    # ---------------------------------------------------------------- partie 9
    "Provisions techniques": "comptes 15801 et 15802 (19011 et 19012 avant leur reclassement) des " + BAL + ", rapportés aux prestations de l'année ; 2026 sur huit mois",
    "Provisions attendues par produit": "provisions au 31/12/2025 des comptes 158 et 196 de « BALANCE FPPN 2025.xlsx » ; besoins mesurés dans les diapositives suivantes de la partie 9",
    "PSAP et dettes envers les cliniques": "compte 15801000 et compte 40110000 (dettes envers les cliniques) des " + BAL + ", au 31/12 de 2020 à 2025 et au 31/08/2026",
    "Flotte automobile : risque porté par GNA": "compte 47111106 (GNA) et compte 71800000 des balances 2023 à 2026 ; précomptes 9W3 de 2024 ; convention avec GNA non reçue",
    "Départ à la retraite du personnel": "comptes 196, 6911 et 7911 des " + BAL + ", exercices 2020 à 2025",
    "Provision du personnel par entité de 2021 à 2025": "agents des DISA 2021, 2022, 2023 et 2025 (« DISA_256927_<année>.xls »), au 31/12 de chaque année ; 2024 non calculé (DISA scannée) ; hors policiers détachés",
    # ---------------------------------------------------------------- partie 9.2
    "Structure de l'actif": "comptes d'actif (classes 2 à 5, 41 et 47) des " + BAL + ", au 31/12 de 2020 à 2025 et au 31/08/2026 ; valeurs comptables non réévaluées",
    "Adossement de liquidité provisoire": "actifs classés par délai de mobilisation et passif par échéance d'après la nature des comptes des " + BAL + ", au 31/12/2025 et au 31/08/2026",
    "Pourquoi l'adossement est une borne favorable": "corrections attendues tirées des textes cités et des constats du rapport ; sens de l'effet seulement, sans montant",
    "Règles de valorisation des actifs": "règles du règlement d'exécution n° 004/2022 (art. 17, 19 à 21, 27), de la Décision n° 05/2022 et du PCMS ; liste ivoirienne des actifs admis non reçue",
    "Scénarios de couverture des engagements": "actifs et engagements au 31/12/2025 et au 31/08/2026 des " + BAL + " ; valorisation selon les règles de la diapositive précédente",
    # ---------------------------------------------------------------- partie 9.3
    "Fonds propres": "comptes 10, 11, 12 et résultat des " + BAL + ", au 31/12 de 2020 à 2025 et au 31/08/2026 ; sans retraitement des provisions",
    "Placements du FPPN": "comptes 52210200 et 52210201 (DAT BOA) et 53300001 (FCP) des balances 2021 à 2026 ; décision de placement selon l'entretien du 19/10/2026",
    "Rendement des placements": "produits financiers (classe 77) et placements moyens (comptes 52 et 53) des balances 2021 à 2025 ; taux directeurs de la BCEAO de 2020 à 2026",
    # ---------------------------------------------------------------- partie 10
    "Textes applicables": "textes de la mutualité sociale UEMOA et ivoirienne fournis le 18/10/2026 ; versions en vigueur au 22/10/2026",
    "Risques courts et risques longs": "garanties du FPPN qualifiées selon le règlement d'exécution n° 004/2022 (art. 1, 3, 12 à 15), sur les fichiers de précomptes et les " + BAL,
    "Deux lectures : comme le FPPN, puis retraitée": "reclassements de comptes selon le PCMS (fiches 1 et 25) appliqués aux " + BAL,
    "Tableau de bord (1/4) : solvabilité et réserves": RAT + ", 2020 au 31/08/2026 ; statut opposable de 2022 à 2025, indicatif en 2020, 2021 et 2026",
    "Tableau de bord (2/4) : risques longs": RAT + ", 2020 au 31/08/2026 ; épargne traitée comme risque long ; statut opposable de 2022 à 2025",
    "Tableau de bord (3/4) : gestion": RAT + " et des fichiers de précomptes (mois reçus) ; statut opposable de 2022 à 2025",
    "Tableau de bord (4/4) : données en attente": "indicateurs de la Décision n° 05/2022 non calculables avec les documents reçus au 22/10/2026 ; numéros de la PBC",
    "Liquidité et équilibre : deux lectures": RAT + ", lecture comme le FPPN et lecture retraitée, 2020 au 31/08/2026",
    "Excédent par risque": "comptes de prestations, de frais (61 à 66 hors 6518) et de cotisations (7010, 7021) des " + BAL + ", 2020 à 2025 ; frais répartis au prorata des cotisations (choix Adenka)",
    "Affectation des excédents": "résultats et réserves (comptes 10 à 13) des " + BAL + ", affectations de 2021 à 2025",
    "Marge de solvabilité": "marge de la Décision n° 05/2022 (annexe, § 1.1.1) sur les " + BAL + ", 2020 au 31/08/2026 ; pertes reportées déduites ou non selon la lecture",
    "Contrôle du plan comptable": "comptes des " + BAL + " comparés à la liste des comptes du PCMS (mai 2012, fiches 1, 2, 6 et 25)",
    "Exigences non chiffrées": "exigences des textes cités confrontées aux documents reçus jusqu'au 22/10/2026",
    "Exposition aux sanctions": "sanctions de l'arrêté n° 00115/MEPS/CAB du 12/12/2023 et du règlement n° 07/2009 rapprochées des constats du rapport ; ne vaut pas avis juridique",
    "Analyses de sensibilité : méthode": "méthode appliquée dans les " + SENS + " ; références ISAP 1, IAS 19, IFRS 17 et IFRS 7",
    "Ratios non conformes de 2020 au 31/08/2026": RAT + ", 2020 au 31/08/2026, recalculés par formules dans les " + SENS,
    "Cinq facteurs récurrents expliquent les écarts": "comptes en cause des " + BAL + ", 2020 au 31/08/2026 ; " + SENS,
    "Montant à combler par ratio": "écarts aux normes de la Décision n° 05/2022 calculés dans les " + SENS + ", 2020 au 31/08/2026",
    "Effet de chaque levier sur chaque ratio": "choc de 100 000 000 FCFA appliqué à l'exercice 2025 dans les " + SENS + " ; effets de premier ordre seulement",
    "Montants d'équilibre de chaque levier en 2025": "exercice 2025, " + SENS + " ; un levier à la fois, sans interaction",
    "Cotisations : deux ratios demandent des mouvements opposés": "exercice 2025, cotisations des comptes 7010 et 7021 de « BALANCE FPPN 2025.xlsx » ; " + SENS,
    "Liquidité : la vente des placements longs ne suffirait pas": "bilan au 31/12/2025 de « BALANCE FPPN 2025.xlsx » ; plafonds de réallocation des " + SENS,
    "Réserve et fonds propres : la règle des statuts aurait suffi pour la réserve": "dotation statutaire (statuts, art. 81 et 90) appliquée aux excédents de 2020 à 2025 des " + BAL,
    "Les normes de flux ne peuvent pas être respectées ensemble": "normes du tableau 1 (lignes 2 et 5) et du tableau 2 de la Décision n° 05/2022, exercice 2025",
    "Inflation médicale : sous-indices de prix de la santé": "IHPC santé de l'ANStat (national, extraction du 19/10/2026), 2020 à 2025, pondéré par les prestations maladie 2025 du FPPN",
    "Sinistralité face à une hausse des prestations": "exercice 2025 ; hausses de prestations hypothétiques appliquées aux prestations de « BALANCE FPPN 2025.xlsx »",
    "Rendement des placements et taux du marché": "rendement du FPPN de 2021 à 2025 (classe 77 et comptes 52 et 53) ; taux de la BCEAO et obligations ivoiriennes à 5 et 10 ans",
    # ---------------------------------------------------------------- partie 11
    "Choix des tables": "tables « CIMA_H.xlsx », « CIMA_F.xlsx », « TD_CIMA.xlsx », « TV_CIMA.xlsx » fournies le 18/10/2026 ; non comparées à la mortalité des policiers",
    "Courbes de survie": "colonne Lx des tables CIMA H, CIMA F, TD et TV, âges 0 à 110 ans ; population générale de la zone CIMA, non le FPPN",
    "Probabilités de décès": "colonne qx des tables CIMA H, CIMA F, TD et TV, âges 20 à 80 ans ; population générale de la zone CIMA, non le FPPN",
    "Indicateurs de référence": "tables CIMA H, CIMA F, TD et TV, taux technique de 3,5 % ; valeurs de référence, non engagements du FPPN",
    # ---------------------------------------------------------------- partie 13 et 14
    "Environnement de travail": "scripts du dossier de la mission tels qu'utilisés pour ce rapport",
    "Étapes du traitement des données": "traitements appliqués aux fichiers reçus jusqu'au 22/10/2026",
    "Analyses en attente": "documents demandés non reçus au 22/10/2026 (PBC pointée)",
    "Points à clarifier": "questions ouvertes au 22/10/2026",
    "Points issus de l'entretien cotisations": "déclarations du responsable des cotisations le 19/10/2026, non confirmées par document sauf mention",
    "Provisions : documents attendus": "documents de provisions demandés et non reçus au 22/10/2026 (PBC pointée)",
    "Points issus des entretiens du chef comptable": "déclarations du chef comptable le 19/10/2026 (14h et 16h), non confirmées par document sauf mention",
}
