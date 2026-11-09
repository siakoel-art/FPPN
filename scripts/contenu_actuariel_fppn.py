# -*- coding: utf-8 -*-
"""Contenu du volet actuariel et statistique, partagé par le résumé exécutif général (parties actuarielles) et le
résumé exécutif Word du seul volet actuariel. Tous les chiffres viennent du rapport actuariel et statistique du
23/10/2026, mis à jour (départ à la retraite du personnel à 60 ans), et du classeur des calculs.

Syntaxe : **gras**, [[surligné jaune = à confirmer]].
"""
RAP = "Rapport actuariel et statistique du 10/11/2026"

# ---------------------------------------------------------------- objectifs
QUESTIONS = [
    ("Engagements", "Que doit le FPPN à ses adhérents et à leurs ayants droit : épargne et intérêts, allocations décès, soins, retraite du personnel ?"),
    ("Provisions", "L'argent mis de côté au bilan suffit-il pour payer ce qui reste dû sur les prestations déjà nées ?"),
    ("Adossement actif-passif", "Le FPPN a-t-il l'argent disponible au bon moment pour payer ce qu'il doit, en particulier l'épargne ?"),
    ("Placements", "L'épargne des policiers est-elle placée de façon sûre, disponible et assez rentable pour payer les intérêts promis ?"),
]

# ---------------------------------------------------------------- approche
APPROCHE_INTRO = ("Le volet actuariel et statistique a répondu à quatre questions posées par la mission : les engagements envers les "
                  "adhérents, la suffisance des provisions, l'adossement actif-passif et l'adéquation des placements. Pour y répondre, il a "
                  "d'abord contrôlé les données, puis mesuré les effectifs, les cotisations et la sinistralité, et recalculé les 36 indicateurs "
                  "de solidité de la Décision n° 05/2022 de l'UEMOA.")
APPROCHE_ETAPES = [
    "**Collecte et contrôle des données :** lecture automatisée de 4 287 fichiers transmis par le FPPN, contrôle de 4 454 416 lignes de cotisations sur 80 mois (janvier 2020 à août 2026), rapprochement avec les balances comptables et les états financiers ;",
    "**Mesure de l'activité :** effectifs cotisants, taux réellement prélevés, sinistralité par garantie (maladie, décès), frais de gestion et résultat ;",
    "**Évaluation des engagements et des provisions :** recalcul de chaque provision et comparaison avec le montant inscrit au bilan ;",
    "**Adossement et placements :** classement des actifs selon leur disponibilité, rythme de sortie de l'épargne, tests de résistance, rendement des placements ;",
    "**Indicateurs prudentiels :** calcul des 36 indicateurs de la Décision n° 05/2022 de 2020 au 31/08/2026, en deux lectures (comptes tels que tenus par le FPPN, puis classés selon le plan comptable des mutuelles sociales), et chiffrage des leviers de conformité ;",
    "**Traçabilité :** chaque chiffre du rapport (2 885 au total) est relié à une formule du classeur des calculs et à sa source.",
]
METHODES = [  # question, méthode retenue, pourquoi ce choix, données utilisées
    ("Engagements : épargne (EFC, PEP)", "Solde comptable de l'épargne, plus les intérêts promis calculés sur l'encours moyen de chaque année, aux taux du manuel d'épargne (3,50 % pour le PEP, 5,50 % pour l'EFC), avec une variante à 3,50 % pour l'EFC",
     "Aucun relevé individuel ni taux arrêté n'existe : seul le calcul global sur les soldes est possible et vérifiable", "Balances de 2020 au 31/08/2026, manuel des produits d'épargne"),
    ("Engagements : décès", "Rapprochement des états de fin d'année des décès non payés avec la comptabilité, et mesure des décès déclarés après la clôture",
     "Le FPPN calcule déjà sa dette décès dossier par dossier : la vérifier est plus fiable que la refaire", "États de fin d'année 2021 à 2025, compte 41591000"),
    ("Provisions maladie", "Mesure des soins d'une année facturés l'année suivante (méthode des déclarations tardives, Code CIMA, art. 334-12)",
     "Les cliniques sont payées en quelques mois : les méthodes sur plusieurs années (Chain Ladder) ne s'appliquent pas", "Journal des factures de 2023 à 2025, délais de paiement de 2020 à 2026"),
    ("Provision du personnel", "Méthode actuarielle des unités de crédit projetées, agent par agent, avec un départ à la retraite à 60 ans",
     "Elle tient compte du départ possible avant la retraite, des hausses de salaire et du temps : c'est la méthode de référence", "Déclarations de salaires (DISA) 2021 à 2025, 79 agents en 2025"),
    ("Prêts du Grenier", "Dépréciation des prêts en retard selon les taux de la politique de crédit du Grenier (40 %, 80 %, 100 %)",
     "C'est la règle que le Grenier s'est lui-même fixée", "Balances âgées de 2 959 prêts"),
    ("Adossement actif-passif", "Classement des actifs selon leur délai de mobilisation, rythme de sortie de l'épargne observé de 2021 à 2025, écarts de liquidité année par année et trois scénarios (central, choc, choc sévère)",
     "Sans échéancier de l'épargne ni données individuelles, les méthodes plus fines (duration, modèle stochastique) ne sont pas applicables", "Bilans au 31/12/2025 et au 31/08/2026, sorties d'épargne de 2021 à 2025"),
    ("Placements", "Encours de chaque placement lu jour par jour dans les grands livres, rendement comparé aux taux de la BCEAO et des titres de l'État",
     "L'inventaire signé des placements n'a pas été transmis : les grands livres sont la seule source complète", "Grands livres de 2021 au 31/08/2026"),
]

# ---------------------------------------------------------------- limites
LIMITES_INTRO = ("Les conclusions reposent sur les documents reçus jusqu'au 09/11/2026. Plusieurs données n'ont pas été transmises. "
                 "Elles ne changent pas le sens des conclusions, mais elles en limitent la précision.")
LIMITES = [
    "**Base des prestations datée :** sans la date des soins, du dépôt et du paiement de chaque facture, les provisions ne peuvent pas être calculées dossier par dossier ;",
    "**Dates de naissance :** absentes des fichiers de cotisations, et connues pour 105 retraités seulement. La mortalité par âge, l'engagement santé des retraités et la projection par âge ne peuvent pas être calculés ;",
    "**Relevés individuels d'épargne et taux servi :** aucun compte par épargnant ni taux arrêté par une décision signée. L'épargne due est mesurée globalement, sur les soldes comptables ;",
    "**Fichiers de décembre 2024 :** les fichiers maladie et décès de ce mois sont vides ;",
    "**Placements et actifs :** ni inventaire ni échéancier signés des placements, ni évaluation récente des terrains et de la clinique. Deux terrains n'ont pas de titre foncier et deux créances n'ont pas d'accord de remboursement ;",
    "**Situation au 31/08/2026 :** établie sur une balance provisoire.",
]
LIMITES_EFFET = ("En conséquence, deux réponses sur quatre (engagements et provisions) sont seulement « partiellement démontrées ». "
                 "Elles deviendront définitives à réception de ces données (liste des documents demandés au FPPN).")

# ---------------------------------------------------------------- résultats par question
VERDICTS = [  # question, boule, verdict, en une phrase, détail
    ("1. Engagements envers les adhérents", "🟠", "Partiellement démontré",
     "L'épargne, les allocations décès, la maladie et la provision du personnel sont mesurés. Les intérêts promis sur l'épargne ne sont pas comptabilisés et la santé des retraités n'est pas mesurable.", "Diapositive 5"),
    ("2. Suffisance des provisions", "🟠", "Partiellement démontré",
     "Les provisions maladie et décès dépassent largement les besoins mesurés, mais ne résultent d'aucun calcul. Les intérêts de l'épargne et les pertes probables sur prêts ne sont pas provisionnés.", "Diapositive 6"),
    ("3. Adossement actif-passif", "🔴", "Insuffisant",
     "L'argent disponible dans l'année couvre environ un tiers de ce que le FPPN devra payer dans l'année, épargne comprise.", "Diapositive 7"),
    ("4. Adéquation des placements", "🔴", "Insuffisant",
     "Les placements financiers sont trop petits, peu rémunérateurs et décidés sans politique votée. Ils ne couvrent ni les intérêts promis ni les sorties d'épargne.", "Diapositive 8"),
]
RESULTATS = [  # titre, message en gras, paragraphes, renvoi
    ("Les engagements envers les adhérents",
     "La première dette du FPPN est l'épargne des policiers, et elle est sous-évaluée d'environ 3,3 milliards de FCFA.",
     ["L'épargne inscrite au bilan vaut 15 743 516 391 FCFA au 31/08/2026. Les intérêts promis par le manuel d'épargne n'ont jamais été comptabilisés depuis le 01/01/2020 : ils représentent 3 278 913 238 FCFA, ou 2 541 799 542 FCFA si l'EFC est servie à 3,50 %. L'engagement réel d'épargne atteint donc 19 022 429 629 FCFA.",
      "Les autres engagements mesurés sont des dettes de court terme : 489 419 669 FCFA d'allocations décès dues aux familles, en hausse chaque année, et environ 7 100 000 FCFA de décès déclarés après la clôture. La santé des 610 policiers retraités cotisants n'est pas mesurable, faute de suivi de leurs soins et de leurs dates de naissance."],
     "Rapport, diapositives 5, 101 à 104, 117 et 130"),
    ("La suffisance des provisions",
     "Les provisions ne manquent pas en montant global, mais elles sont figées depuis 2020 et mal affectées.",
     ["Une provision est de l'argent mis de côté pour une dépense déjà due mais pas encore payée. Le bilan porte 4 829 298 290 FCFA de provisions maladie et décès, identiques au franc près de 2020 au 31/08/2026. [[Ce montant aurait été passé vers 2004 selon la direction comptable.]] Le besoin réellement mesuré n'est que de 19 581 392 FCFA : l'excédent atteint 4 809 716 898 FCFA.",
      "À l'inverse, rien n'est mis de côté pour les intérêts promis sur l'épargne (3 278 913 238 FCFA) ni pour les pertes probables sur les prêts du Grenier (14 016 392 FCFA). La provision pour la retraite du personnel est suffisante : 43 666 678 FCFA au bilan pour 17 184 450 FCFA recalculés avec un départ à 60 ans."],
     "Rapport, diapositives 6, 121 à 140 et 152"),
    ("L'adossement actif-passif",
     "Le FPPN n'a pas assez d'argent disponible pour payer ce qu'il devra payer dans l'année.",
     ["Au 31/08/2026, l'argent mobilisable en moins d'un an vaut 3 457 866 951 FCFA. Les paiements attendus dans l'année, sorties d'épargne comprises, valent 10 046 113 184 FCFA. La couverture n'est que de 34,42 %, et de 26,22 % dans le scénario de choc sévère. Le manque cumulé atteint 12 596 052 822 FCFA à cinq ans.",
      "La cause est simple : l'épargne, que les policiers peuvent retirer à tout moment, finance des biens qui ne se vendent pas vite (terrains, clinique, créance sur l'Hôpital de la Police, prêts du Grenier). Seuls 16,28 % de l'épargne sont couverts par des actifs qui lui sont réservés."],
     "Rapport, diapositives 7, 96 et 155 à 170"),
    ("L'adéquation des placements",
     "Les placements sont trop faibles pour payer les intérêts promis aux épargnants.",
     ["En 2025, l'argent réellement placé (dépôts à terme et fonds communs de placement) vaut en moyenne 763 752 135 FCFA, soit 5,53 % de l'épargne due. Son rendement moyen de 2021 à 2025 est de 4,90 % par an. Les produits financiers de 2025 (117 912 488 FCFA) sont inférieurs aux intérêts d'une seule année sur l'EFC au taux le plus bas (309 709 577 FCFA).",
      "Le problème n'est pas d'abord le rendement, mais le volume et la disponibilité : 2 356 399 123 FCFA dorment sur des comptes courants sans rapporter, et aucune politique de placement n'a été votée par le conseil d'administration."],
     "Rapport, diapositives 8 et 171 à 178"),
]
AUTRES_CONSTATS = [
    "**La garantie maladie perd de l'argent depuis 2025 :** les soins consomment 73,53 % des cotisations maladie en 2025 et 97,40 % sur les huit premiers mois de 2026 (109,87 % frais compris en 2025). Le décès, excédentaire, finance la maladie ;",
    "**La répartition des cotisations diffère des statuts :** la Solde prélève 3,77 % pour la maladie et 1,23 % pour le décès, au lieu de 3,50 % et 1,50 % ;",
    "**Des données essentielles manquent :** ni base des prestations datée, ni dates de naissance, ni fichier des ayants droit ;",
    "**Le Grenier ne déprécie aucun prêt et inscrit aussi l'épargne PEP et EFC dans ses comptes :** 8 493 055 501 FCFA d'épargne figurent dans les deux comptabilités, et son résultat 2025 (30 543 376 FCFA) est calculé avant toute perte sur prêts ;",
    "**Gouvernance :** l'AIRMS a jugé nulle la prorogation de deux ans des mandats votée le 17/04/2026 et a mis le FPPN en demeure de transmettre ses comptes 2025. Ces points relèvent du volet juridique.",
]

# ---------------------------------------------------------------- ratios prudentiels (travaux du volet actuariel)
RATIOS_INTRO = ("Le volet actuariel a recalculé les 36 indicateurs de solidité financière de la Décision n° 05/2022 de l'UEMOA, "
                "de 2020 au 31/08/2026, à partir des balances du FPPN.")
RATIOS_PARAS = [
    "**Le FPPN ne respecte pas 14 règles de solidité sur 36 de 2022 à 2025, ce qui expose son agrément.** 7 indicateurs sont conformes et 15 restent en attente de pièces. Deux exercices de suite non conformes permettent un retrait d'agrément (règlement n° 07/2009/CM/UEMOA, art. 84) et les sanctions de l'arrêté n° 00115/MEPS/CAB du 12/12/2023.",
    "**Trois écarts dominent en 2025.** La réserve obligatoire couvre 5,01 mois de prestations, pour une norme de 6 à 9 mois. Le ratio de liquidité vaut 0,13 pour une norme d'au moins 1 (1,15 si l'épargne est classée en ressource longue). Le ratio de fonds propres et assimilés vaut 0,22 pour une norme d'au moins 1.",
    "**Les fonds propres sont négatifs à chaque clôture depuis 2020 :** -438 242 200 FCFA au 31/12/2025 et -879 498 700 FCFA au 31/08/2026. Les frais de fonctionnement atteignent 36,34 % des cotisations en 2025, pour un maximum de 20 %.",
    "**Le FPPN calcule lui-même ses ratios, et ses valeurs confirment le diagnostic.** Elles concordent pour la sinistralité (57,48 % en 2025) et les frais de fonctionnement (36,33 %). Elles divergent pour la réserve obligatoire (9 mois contre 5,01 mois recalculés), la liquidité et les fonds propres (0,31 contre 0,22), car le FPPN retraite l'épargne. Même avec ses chiffres, les fonds propres restent hors norme.",
]
RATIOS_TABLE = [  # indicateur, norme, valeur 2025, statut
    ("Réserve obligatoire", "6 à 9 mois de prestations", "5,01 mois", "🔴"),
    ("Liquidité", "Au moins 1", "0,13 (1,15 en lecture corrigée)", "🔴"),
    ("Fonds propres et assimilés", "Au moins 1", "0,22", "🔴"),
    ("Frais de fonctionnement", "Au plus 20 % des cotisations", "36,34 %", "🔴"),
    ("Sinistralité réglementaire, toutes garanties", "70 % à 80 %", "57,48 % (75,20 % au 31/08/2026)", "🔴"),
]
RATIOS_RENVOI = "Rapport, diapositives 19 et 180 à 231"

# ---------------------------------------------------------------- constats classés par criticité (13 constats codés)
CRIT_LIB = {"🔴": "Critique", "🟠": "Élevé", "🟡": "Moyen", "🟢": "Faible"}
CONSTATS = [  # boule, code, constat en une phrase simple, ce que cela veut dire concrètement, chiffre clé, détail
    ("🔴", "ENG-1", "Les intérêts promis aux épargnants ne sont pas comptabilisés.",
     "Le FPPN doit plus à ses épargnants que ce que montrent ses comptes, et il ne peut pas dire à chacun ce qui lui revient.",
     "3 278 913 238 FCFA d'intérêts non inscrits, sur 15 743 516 391 FCFA d'épargne", "5, 100 à 103"),
    ("🔴", "ALM-1", "L'argent disponible ne couvre qu'un tiers des paiements de l'année.",
     "Si les retraits d'épargne suivent leur rythme habituel, le FPPN devra vendre ou mobiliser des biens qui ne se vendent pas vite.",
     "Couverture de 34,42 % (26,22 % en choc sévère), manque de 12 596 052 822 FCFA à cinq ans", "7, 152 à 165"),
    ("🔴", "PLA-1", "Les placements sont trop faibles pour payer les intérêts promis.",
     "L'épargne rapporte moins que ce qui est promis aux épargnants, et une grande partie dort sans rien rapporter.",
     "763 752 135 FCFA placés en 2025, soit 5,53 % de l'épargne, rendement de 4,90 %", "8, 166 à 173"),
    ("🔴", "REG-1", "Le FPPN ne respecte pas 14 règles de solidité sur 36.",
     "Deux exercices de suite non conformes permettent à l'autorité de contrôle de retirer l'agrément du FPPN.",
     "Fonds propres de -879 498 700 FCFA, liquidité de 0,13 en 2025", "19, 175 à 225"),
    ("🔴", "ACT-1", "Certains actifs ne sont pas correctement évalués.",
     "Deux terrains n'ont pas de titre foncier et deux créances n'ont pas d'accord de remboursement : leur valeur au bilan surestime ce que le FPPN peut mobiliser pour payer l'épargne.",
     "3 782 519 850 FCFA à sortir de l'actif selon le commissaire aux comptes (terrain d'Attoban et créance sur l'Hôpital de la Police)", "7, 158"),
    ("🟠", "ALM-2", "Des engagements hors bilan pèsent sur la trésorerie future.",
     "Un procès perdu bloque une partie de l'argent, et un projet immobilier plus grand que l'épargne due est approuvé sans plan de financement.",
     "263 000 000 FCFA bloqués à la BOA, projet d'Azaguié de 19 539 377 111 FCFA", "159"),
    ("🟠", "ENG-2", "La dette envers les familles des policiers décédés grossit chaque année.",
     "Des familles attendent leur allocation, et la date de leur demande n'est pas enregistrée.",
     "489 419 669 FCFA dus au 31/08/2026, contre 370 503 332 FCFA fin 2021", "114 à 116, 127 à 130"),
    ("🟠", "PROV-1", "Les provisions maladie et décès sont figées depuis 2020.",
     "L'argent mis de côté ne correspond à aucun calcul : il gonfle les dettes du bilan, alors que d'autres dettes réelles ne sont pas couvertes.",
     "4 829 298 290 FCFA au bilan pour un besoin mesuré de 19 581 392 FCFA", "6, 120 à 130"),
    ("🟠", "STAT-2", "La garantie maladie perd de l'argent, et le décès la finance.",
     "La cotisation maladie ne suffit plus à payer les soins et les frais : sans correction, le déficit va augmenter.",
     "73,53 % des cotisations maladie consommées en 2025, 97,40 % en 2026", "79 à 92"),
    ("🟠", "REG-2", "Les décisions qui engagent l'épargne ne sont pas encadrées.",
     "L'épargne des policiers est employée sans décision du conseil d'administration, alors que les statuts l'exigent.",
     "[[Placements décidés par la direction générale]], projet de Cocody sans budget", "164, 170, 171"),
    ("🟠", "DON-1", "Des données essentielles manquent pour finir les calculs.",
     "Plusieurs calculs restent impossibles tant que ces fichiers ne sont pas transmis : provisions dossier par dossier, mortalité, santé des retraités.",
     "Ni base des prestations datée, ni dates de naissance, fichiers de décembre 2024 vides", "55 à 66"),
    ("🟡", "STAT-1", "La répartition des cotisations diffère des statuts.",
     "La maladie reçoit plus et le décès moins que ce que prévoient les statuts, ce qui fausse le résultat de chaque garantie.",
     "3,77 % et 1,23 % prélevés au lieu de 3,50 % et 1,50 %", "67 à 78"),
    ("🟡", "STAT-3", "Les calculs par âge restent à faire.",
     "La projection de l'équilibre du FPPN n'est que globale tant que les âges des policiers ne sont pas connus.",
     "Projection de 2026 à 2030 en deux scénarios", "226 à 236"),
    ("🟡", "GRE-1", "Le Grenier ne déprécie aucun prêt et inscrit l'épargne PEP et EFC dans ses propres comptes.",
     "La même épargne figure dans deux comptabilités, et le résultat du Grenier est calculé avant toute perte sur prêts.",
     "8 493 055 501 FCFA d'épargne au passif du Grenier, résultat 2025 de 30 543 376 FCFA", "133"),
    ("🟢", "PROV-3", "Les pertes probables sur les prêts du Grenier ne sont pas comptées.",
     "Le montant est faible, mais le vrai risque est l'épargne prêtée : le FPPN finance presque tout le crédit du Grenier.",
     "14 016 392 FCFA à déprécier, aucune dépréciation depuis 2022", "132"),
    ("🟢", "PROV-2", "La provision pour la retraite du personnel est suffisante.",
     "C'est le seul poste où le montant du FPPN couvre largement le besoin. Ses fichiers de calcul confirment sa méthode : les droits acquis, sans actualisation.",
     "43 666 678 FCFA au bilan pour 17 184 450 FCFA recalculés (départ à 60 ans)", "135 à 147"),
]

# ---------------------------------------------------------------- plan d'actions priorisé (modèle validé)
ECHEANCE = {"P1": "Dans les 3 mois", "P2": "Dans les 6 mois", "P3": "De 6 à 18 mois", "P4": "Au-delà de 18 mois"}
BOULE_P = {"P1": "🔴", "P2": "🟠", "P3": "🟡", "P4": "🟢"}
PLAN = [  # axe, action, constats, priorité, responsable
    ("Engagements et épargne", "Fixer par délibération le taux servi sur l'épargne", "ENG-1", "P1", "Conseil d'administration"),
    ("Engagements et épargne", "Comptabiliser les intérêts dus et tenir un compte par épargnant", "ENG-1", "P1", "Service comptable, direction générale"),
    ("Engagements et épargne", "Faire réaliser l'étude actuarielle des risques longs tous les cinq ans", "ENG-1, STAT-3", "P4", "Direction générale"),
    ("Provisions et comptes", "Établir une note de calcul annuelle des provisions, signée et présentée à l'Assemblée générale", "PROV-1, PROV-2", "P1", "Service comptable, direction générale"),
    ("Provisions et comptes", "Reclasser l'excédent de provisions après décision des organes et avis de l'AIRMS", "PROV-1", "P1", "Conseil d'administration, Assemblée générale"),
    ("Provisions et comptes", "Déprécier les prêts du Grenier et tenir une comptabilité séparée de la clinique", "PROV-3", "P2", "Service comptable, Grenier"),
    ("Provisions et comptes", "Fixer le propriétaire de l'épargne PEP et EFC et supprimer sa double inscription dans les comptes du Grenier", "GRE-1", "P2", "Direction générale, service comptable"),
    ("Liquidité et placements", "Constituer une réserve de liquidité égale aux sorties d'épargne d'une année", "ALM-1", "P1", "Conseil d'administration, direction générale"),
    ("Liquidité et placements", "Établir l'échéancier des sorties d'épargne (départs à la retraite)", "ALM-1", "P1", "Direction générale"),
    ("Liquidité et placements", "Faire voter une politique de placement et évaluer la clinique et les terrains", "PLA-1", "P2", "Conseil d'administration"),
    ("Liquidité et placements", "Décider du sort du terrain d'Attoban et de la créance sur l'Hôpital de la Police, et déprécier ce qui n'est pas recouvrable", "ACT-1", "P1", "Conseil d'administration"),
    ("Liquidité et placements", "Soumettre le projet d'Azaguié à une étude de faisabilité et à un plan de financement avant tout engagement", "ALM-2", "P1", "Conseil d'administration"),
    ("Conformité et gouvernance", "Présenter à l'AIRMS un plan de mise en conformité chiffré", "REG-1", "P1", "Conseil d'administration, direction générale"),
    ("Conformité et gouvernance", "Réconcilier les ratios calculés par le FPPN avec le recalcul du rapport avant leur transmission à l'AIRMS", "REG-1", "P1", "Direction générale"),
    ("Conformité et gouvernance", "Faire délibérer le conseil sur chaque placement et créer un comité actif-passif", "REG-2", "P2", "Conseil d'administration"),
    ("Conformité et gouvernance", "Recalculer la cotisation maladie et contrôler après paiement au moins 5 % des factures", "STAT-2", "P2", "Conseil d'administration, service des prestations"),
    ("Données et pilotage", "Transmettre les données manquantes (dates de naissance, prestations datées)", "DON-1, ENG-2", "P1", "Direction générale, service informatique"),
    ("Données et pilotage", "Rapprocher chaque mois les fichiers de cotisations envoyés et reçus", "STAT-1, DON-1", "P2", "Service des cotisations"),
    ("Données et pilotage", "Tenir un tableau de bord semestriel des indicateurs de solidité", "REG-1", "P3", "Direction générale"),
]
PLAN_RENVOI = "Rapport, diapositives 25 à 27"

# ---------------------------------------------------------------- v14 : ajouts et corrections du 10/11/2026 (écrits en bleu dans les livrables)
BLEU_CA = [
    "jusqu'au 09/11/2026",
    "Deux terrains n'ont pas de titre foncier et deux créances n'ont pas d'accord de remboursement",
    "Le Grenier ne déprécie aucun prêt et inscrit aussi l'épargne PEP et EFC dans ses comptes :",
    "8 493 055 501 FCFA d'épargne figurent dans les deux comptabilités, et son résultat 2025 (30 543 376 FCFA) est calculé avant toute perte sur prêts",
    "Gouvernance :",
    "l'AIRMS a jugé nulle la prorogation de deux ans des mandats votée le 17/04/2026 et a mis le FPPN en demeure de transmettre ses comptes 2025. Ces points relèvent du volet juridique.",
    "Le FPPN calcule lui-même ses ratios, et ses valeurs confirment le diagnostic.",
    "Elles concordent pour la sinistralité (57,48 % en 2025) et les frais de fonctionnement (36,33 %). Elles divergent pour la réserve obligatoire (9 mois contre 5,01 mois recalculés), la liquidité et les fonds propres (0,31 contre 0,22), car le FPPN retraite l'épargne. Même avec ses chiffres, les fonds propres restent hors norme.",
    "Ses fichiers de calcul confirment sa méthode : les droits acquis, sans actualisation.",
    "Certains actifs ne sont pas correctement évalués.",
    "Deux terrains n'ont pas de titre foncier et deux créances n'ont pas d'accord de remboursement : leur valeur au bilan surestime ce que le FPPN peut mobiliser pour payer l'épargne.",
    "3 782 519 850 FCFA à sortir de l'actif selon le commissaire aux comptes (terrain d'Attoban et créance sur l'Hôpital de la Police)",
    "Des engagements hors bilan pèsent sur la trésorerie future.",
    "Un procès perdu bloque une partie de l'argent, et un projet immobilier plus grand que l'épargne due est approuvé sans plan de financement.",
    "263 000 000 FCFA bloqués à la BOA, projet d'Azaguié de 19 539 377 111 FCFA",
    "Le Grenier ne déprécie aucun prêt et inscrit l'épargne PEP et EFC dans ses propres comptes.",
    "La même épargne figure dans deux comptabilités, et le résultat du Grenier est calculé avant toute perte sur prêts.",
    "8 493 055 501 FCFA d'épargne au passif du Grenier, résultat 2025 de 30 543 376 FCFA",
    "Décider du sort du terrain d'Attoban et de la créance sur l'Hôpital de la Police, et déprécier ce qui n'est pas recouvrable",
    "Soumettre le projet d'Azaguié à une étude de faisabilité et à un plan de financement avant tout engagement",
    "Fixer le propriétaire de l'épargne PEP et EFC et supprimer sa double inscription dans les comptes du Grenier",
    "Réconcilier les ratios calculés par le FPPN avec le recalcul du rapport avant leur transmission à l'AIRMS",
    "ACT-1", "ALM-2", "GRE-1",
]

