"""Version PowerPoint de la note interne de point d'avancement (réunion de 30 minutes avec le Chief).

Même fil conducteur que la note Word : contexte, périmètre, tâches principales, tâches ajoutées, travaux réalisés,
conclusions, limites, reste à faire, recommandations, taux d'avancement, prochaines étapes.
Le moteur de mise en page (charte Adenka, Timesaver, tableaux, sigles en bas de diapositive) est repris du script du
rapport : rapport_v12_fppn.py (en-tête, aides et construction), exécutés avec le contenu ci-dessous.
Le texte à dire figure dans les notes du présentateur de chaque diapositive.

Usage : python3 deck_point_chief_fppn.py <skill adenka-powerpoint> <sortie.pptx>
"""
import os
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
SKILL, SORTIE = sys.argv[1:3]
SRC = open(os.path.join(ICI, "rapport_v12_fppn.py"), encoding="utf-8").read().split("\n")


def lignes(a, b):
    """Lignes a à b (numérotation de l'éditeur, bornes comprises) du script du rapport."""
    return "\n".join(SRC[a - 1:b])


ENTETE = "\n".join([lignes(82, 87), lignes(90, 91), lignes(94, 94), lignes(96, 104), lignes(136, 187), lignes(238, 375), lignes(396, 411), lignes(602, 635)])
debut = next(i for i, l in enumerate(SRC, 1) if l.startswith("tmp = SORTIE.replace"))
fin = next(i for i, l in enumerate(SRC, 1) if l.startswith("finaliser(SORTIE)"))
MOTEUR = lignes(debut, fin + 1)
for avant, apres in [('"Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"', '"Audit Comptable et financier FPPN", "Text Placeholder 4": "26 octobre 2026"'),
                     ('r.text = "Rapport actuariel et statistique"', 'r.text = "Point d\'avancement de la mission"'),
                     ('entrees = [("Résumé exécutif", 3)]', 'entrees = [("L\'essentiel", 3)]')]:
    assert avant in MOTEUR, avant
    MOTEUR = MOTEUR.replace(avant, apres)

CONTENU = r'''
VALIDITE = {}
PBCM = {}
GAB = {"couverture": 1, "sommaire": 5, "separateur": 2, "contenu": 5, "annexe": 5, "merci": 110}
RAPPORT = "Source : rapport actuariel et statistique du 23/10/2026"
SL.append({"type": "couverture", "notes": "Point d'avancement interne du volet actuariel et statistique de l'audit du FPPN. Durée : 30 minutes, dont 7 minutes de questions."})
SL.append({"type": "sommaire", "notes": "Annoncer les quatre temps : la mission et mon périmètre, ce qui a été fait, ce qu'on en conclut, puis les recommandations, l'avancement et la suite."})

gab(13, {"titre": "L'essentiel", "sous_titre": "Le FPPN dégage des excédents, mais ne peut pas démontrer qu'il remboursera l'épargne de ses adhérents",
         "synthese": "Deux réponses sur quatre sont partiellement démontrées, deux sont insuffisantes, et les travaux sont réalisés à 63,85 %"},
    {"cartes": [("La mission", "Répondre à quatre questions sur la solidité du FPPN : engagements, provisions, adossement, placements, du 01/01/2020 au 31/08/2026."),
                ("Les conclusions", "Épargne due de 15 743 516 391 FCFA sans les intérêts promis. Provisions figées depuis 2020. Actifs à moins d'un an : 34,42 % des paiements de l'année."),
                ("L'avancement", "63,85 % des travaux réalisés après 25,71 % du calendrier. Le reste attend surtout les données du FPPN."),
                ("La suite", "Relance du FPPN, calculs dossier par dossier, restitution provisoire les 23 et 24/11/2026, rapport définitif le 30/11/2026.")]},
    "Phrase à garder en tête : le FPPN dégage des excédents chaque année, mais il n'est pas en mesure de démontrer qu'il pourra rembourser l'épargne de ses adhérents. Tout le reste du rapport la démontre ou la nuance.",
    source=RAPPORT + ", diapositives 3 à 8 ; plan de travail du 23/10/2026")

# ---------------------------------------------------------------- 1. mission et périmètre
separateur(1, "Mission et périmètre", "Le contexte, mon périmètre, les tâches principales et les tâches ajoutées")
gab(13, {"titre": "Contexte de la mission", "sous_titre": "Le FPPN gère la prévoyance et l'épargne des policiers, et l'audit doit dire s'il peut tenir ses promesses",
         "synthese": "Adenka réalise l'audit comptable et financier du FPPN sur la période du 01/01/2020 au 31/08/2026"},
    {"cartes": [("Le FPPN", "Mutuelle des policiers de Côte d'Ivoire. Les cotisations sont prélevées sur la solde, ou sur la pension pour les retraités."),
                ("Ses garanties", "Remboursement des soins, allocation au décès, épargne de fin de carrière (EFC), épargne volontaire (PEP), prêts du Grenier du Policier."),
                ("Son cadre", "Mutualité sociale de l'UEMOA, sous le contrôle de l'AIRMS. Le Code CIMA ne sert que de référence technique."),
                ("L'audit Adenka", "Audit comptable et financier du 01/01/2020 au 31/08/2026, sur la base d'une offre technique validée par le FPPN.")]},
    "« Le FPPN est la mutuelle des policiers. En échange des cotisations prélevées sur la solde, il rembourse les soins, verse une allocation au décès, gère deux épargnes et prête aux policiers. Il est suivi par l'AIRMS. » Transition : « Dans cet audit, mon intervention porte sur un volet précis. »",
    source="Source : offre technique Adenka validée par le FPPN ; statuts du FPPN édités le 15/06/2023")
gab(31, {"titre": "Mon périmètre", "sous_titre": "J'interviens sur le volet actuariel et statistique : ce que le FPPN devra payer et sa capacité à le payer",
         "synthese": "Le périmètre est fixé par le profil de l'actuaire (p. 74) et le pôle 4 (p. 70) de l'offre technique"},
    {"cartes": [("Ce que vérifie l'audit comptable", "Que les comptes sont justes, à la date de clôture."),
                ("Ce que vérifie mon volet", "Ce que le FPPN doit, ce qu'il a mis de côté, et si son argent sera disponible à temps."),
                ("La période", "Du 01/01/2020 au 31/08/2026, à chaque clôture annuelle et au 31/08/2026."),
                ("Ce qui relève de l'expert financier", "La valorisation et le rendement des placements, et les ratios des TDR (offre technique, p. 70 et 74).")]},
    "« L'audit comptable regarde le passé, mon volet regarde l'avenir. » Transition : « Concrètement, ce périmètre se traduit par quatre questions. »",
    source="Source : offre technique Adenka validée par le FPPN, pôle 4 (p. 70) et profil de l'actuaire (p. 74)")
contenu({"titre": "Tâches principales", "sous_titre": "Quatre questions structurent tout le rapport, et chacune reçoit une réponse dès la diapositive 4",
         "synthese": "Les quatre tâches de l'offre technique, traduites en langage simple"},
        "Lire la colonne de droite seulement. Transition : « Pour y répondre correctement, l'offre technique et la nature du FPPN imposent d'autres travaux. »",
        tableau={"titre": None, "entetes": ["Tâche de l'offre technique", "Ce que cela veut dire en clair"], "largeurs": [13.6, 17.6], "taille": 12, "sans_code": True,
                 "lignes": [["1. Évaluer les engagements du FPPN envers ses adhérents et leurs ayants droit", "Combien le FPPN doit-il aujourd'hui et devra-t-il plus tard : épargne, intérêts, prestations maladie et décès ?"],
                            ["2. Apprécier la suffisance des provisions au regard des prestations servies", "Les sommes mises de côté au bilan suffisent-elles à payer les prestations déjà nées ?"],
                            ["3. Apprécier l'adossement actif-passif", "L'argent du FPPN sera-t-il disponible au moment où il devra payer ?"],
                            ["4. Analyser l'adéquation des placements aux engagements", "Les placements sont-ils sûrs, disponibles et assez rentables pour servir l'épargne promise ?"]]},
        source="Source : offre technique Adenka validée par le FPPN, profil de l'actuaire (p. 74) ; rapport du 23/10/2026, diapositive 4")
contenu({"titre": "Tâches ajoutées", "sous_titre": "Ces travaux sont prévus par l'offre ou indispensables pour répondre aux quatre questions",
         "synthese": "Sept tâches complètent les quatre questions, chacune avec son fondement"},
        "Insister : ce n'est pas un dépassement du périmètre. Transition : « Voici maintenant ce qui a été réalisé. »",
        tableau={"titre": None, "entetes": ["Tâche ajoutée", "Pourquoi elle est nécessaire", "Fondement"], "largeurs": [9.4, 13.6, 8.2], "taille": 9.5, "sans_code": True,
                 "lignes": [["Contrôler les données de base (effectifs, cotisations, précomptes)", "Un calcul actuariel ne vaut que par ses données. Les fichiers contenaient des lignes de total, des montants doublés et des fichiers vides.", "Offre technique, pôle 4 (p. 70)"],
                            ["Calculer les 36 ratios de la Décision n° 05/2022 de l'UEMOA", "Apprécier les réserves au regard des exigences prudentielles et les comparer aux indicateurs de l'AIRMS.", "Offre technique, pôle 4 (p. 70) et risque R3 (p. 36)"],
                            ["Projeter l'équilibre technique", "Vérifier que les cotisations futures couvriront les prestations et les frais futurs.", "Offre technique, pôle 4 (p. 70)"],
                            ["Volet statistique : qualité des données, décès, santé", "Mesurer la sinistralité, les délais, les dossiers non payés et la mortalité comparée aux tables CIMA.", "Plan d'analyse statistique du 16/10/2026"],
                            ["Revoir les autres provisions (Grenier, compte GNA, retraite du personnel)", "Elles pèsent sur les fonds propres, donc sur la capacité à payer les adhérents.", "Plan de travail, tâche 4.4"],
                            ["Établir le cadre réglementaire applicable", "Le FPPN relève de la mutualité sociale de l'UEMOA. Le choix du Code CIMA comme référence technique doit être justifié.", "Règlement n° 07/2009, règlement d'exécution n° 004/2022, Décision n° 05/2022, PCMS"],
                            ["Décrire les circuits des cotisations par produit", "Repérer où l'argent peut se perdre entre la Solde, la CGRAE, le GNA et la comptabilité.", "Logigrammes du FPPN, entretien du 19/10/2026"]]},
        source="Source : offre technique Adenka ; plan d'analyse du 16/10/2026 ; plan de travail du 23/10/2026")

# ---------------------------------------------------------------- 2. travaux réalisés
separateur(2, "Travaux réalisés", "Des fichiers bruts du FPPN jusqu'aux recommandations, avec des outils qui rendent chaque chiffre vérifiable")
contenu({"titre": "Ce que j'ai fait", "sous_titre": "Les travaux couvrent tout le périmètre, des fichiers bruts du FPPN jusqu'aux recommandations",
         "synthese": "Sept étapes, de la réunion de lancement du 13/10/2026 aux neuf recommandations"},
        "Citer seulement la colonne de droite. « Le rapport fait 291 diapositives parce que chaque chiffre est sourcé et démontré. Le résumé exécutif, aux diapositives 3 à 27, se suffit à lui-même. Le reste sert de preuve. »",
        tableau={"titre": None, "entetes": ["Étape", "Ce qui a été fait", "Chiffre à citer"], "largeurs": [4.0, 18.8, 8.4], "taille": 10.5, "sans_code": True,
                 "lignes": [["Lancement", "Réunions de lancement des 13 et 14/10/2026, liste des documents demandés (PBC)", "90 documents demandés"],
                            ["Collecte", "Sept envois du FPPN du 15/10 au 23/10/2026, pointés un par un", "6 reçus, 41 partiels, 43 non reçus"],
                            ["Entretiens", "Responsable des cotisations, chef comptable, service des prestations (16 et 19/10/2026), avec guides et comptes rendus", "4 entretiens documentés"],
                            ["Données", "Précomptes sur salaire et sur pension, balances, grands livres, états financiers, fichiers du personnel, prêts du Grenier", "4 454 416 lignes de précomptes contrôlées"],
                            ["Contrôles", "Vrai format des fichiers, lignes de total, montants doublés, fichiers vides, rapprochement aller et retour", "7 balances et 506 feuilles de précomptes relues"],
                            ["Analyses", "Effectifs, cotisations, sinistralité, épargne, prêts, retraités, provisions, adossement, placements, 36 ratios, sensibilités, mortalité, projection", "14 parties de rapport"],
                            ["Synthèse", "Constats, matrices des risques et de conformité, recommandations avec horizon", "9 recommandations"]]},
        source="Source : rapport du 23/10/2026 ; PBC du 23/10/2026 ; plan de travail du 23/10/2026")
contenu({"titre": "Outils et livrables", "sous_titre": "Les outils garantissent que chaque chiffre est exact, retraçable et recalculable",
         "synthese": "Les 2 885 chiffres du rapport sont retrouvés un par un dans les calculs, et le classeur Excel les recalcule avec 0 écart"},
        "Point fort à mettre en avant devant un profil finance ou audit. Montrer l'onglet « 0_controle » du classeur si le Chief le souhaite. Transition : « Ces travaux conduisent à quatre conclusions. »",
        tableau={"titre": None, "entetes": ["Outil", "Ce qu'il garantit"], "largeurs": [9.0, 22.2], "taille": 10.5, "sans_code": True,
                 "lignes": [["Python", "Lecture et contrôle des fichiers bruts du FPPN : plus de quatre millions de lignes, plusieurs formats, fichiers scannés lus par reconnaissance de caractères."],
                            ["Empreinte numérique de chaque fichier", "Les fichiers analysés sont identiques à ceux transmis par le FPPN."],
                            ["Classeur Excel de 99 onglets", "Chaque calcul est refait par une formule visible. L'onglet « 0_controle » compare chaque formule au chiffre du rapport : 0 écart."],
                            ["Contrôle automatique du rapport", "Les 2 885 chiffres écrits dans le rapport sont retrouvés un par un dans les calculs (diapositive 239)."],
                            ["Vérification indépendante", "Les chiffres ont été recalculés une seconde fois depuis les fichiers bruts. 18 anomalies ont été relevées et corrigées."],
                            ["Logigrammes modifiables", "Les circuits du FPPN mis à jour, avec les contrôles manquants signalés."]]},
        analyse=["**Livrables produits à ce jour.** Rapport de 291 diapositives (PowerPoint et PDF), classeur de calcul, PBC pointée, liste des 26 questions et 84 documents à transmettre par le FPPN, vérification de cohérence, plan de travail pointé, logigrammes mis à jour, guides et comptes rendus d'entretien."],
        source="Source : rapport du 23/10/2026, diapositives 229 à 239 ; FPPN_Calculs_rapport_20261023.xlsx ; FPPN_Verification_coherence_20261023.xlsx")

# ---------------------------------------------------------------- 3. conclusions et limites
separateur(3, "Conclusions et limites", "Les quatre verdicts, leur démonstration, leurs limites et ce qui reste à faire")
contenu({"titre": "Conclusions", "sous_titre": "Deux réponses sont partiellement démontrées faute de données, deux sont clairement insuffisantes",
         "synthese": "Pour chaque question : le verdict, le chiffre clé et la façon dont il a été obtenu"},
        "Montrer la diapositive 4 du rapport si besoin. « La non-conformité sur deux exercices de suite permet à l'AIRMS de retirer l'agrément (règlement n° 07/2009, art. 84). Ce n'est pas automatique, mais c'est l'enjeu principal pour le FPPN. » Transition : « Ces conclusions ont des limites qu'il faut connaître. »",
        tableau={"titre": None, "entetes": ["Question et verdict", "Conclusion", "Pourquoi et comment"], "largeurs": [5.0, 12.2, 14.0], "taille": 8.5,
                 "lignes": [["1. Engagements\nPartiellement démontré", "L'épargne due vaut 15 743 516 391 FCFA au 31/08/2026. Avec les intérêts promis, l'engagement atteint 19 022 429 629 FCFA.",
                             "Soldes lus dans les balances. Intérêts recalculés sur l'encours moyen de chaque année, à 3,50 % (PEP) et 5,50 % (EFC) : 3 278 913 238 FCFA jamais comptabilisés. Aucun compte individuel transmis."],
                            ["2. Provisions\nPartiellement démontré", "Les provisions maladie et obsèques (4 829 298 290 FCFA) sont figées depuis 2020 et dépassent de 4 809 716 898 FCFA le besoin mesuré (19 581 392 FCFA).",
                             "Besoin mesuré sur les factures de soins et les décès arrivés après la clôture. Le total n'est pas trop faible, il est mal affecté : rien n'est provisionné pour les intérêts de l'épargne."],
                            ["3. Adossement\nInsuffisant", "Les actifs disponibles à moins d'un an couvrent 34,42 % des paiements attendus dans l'année. Le manque cumulé atteint 12 596 052 822 FCFA à cinq ans.",
                             "Actifs classés selon leur délai de mobilisation, dettes par échéance, sorties d'épargne au taux observé. Actifs à leur valeur comptable : estimation favorable au FPPN."],
                            ["4. Placements\nInsuffisant", "En 2025, l'argent réellement placé (763 752 135 FCFA) représente 5,53 % de l'épargne due. Rendement moyen de 4,90 % par an de 2021 à 2025.",
                             "Dépôts à terme et FCP lus jour par jour dans les grands livres. L'épargne a surtout financé des actifs peu liquides. Aucune politique de placement votée n'a été transmise."]],
                 "sans_code": True},
        analyse=["**Constats transversaux.** Les fonds propres sont négatifs à chaque clôture depuis 2020 (-879 498 700 FCFA au 31/08/2026), et 14 ratios de l'UEMOA sur 36 ne sont pas respectés, dont la réserve obligatoire, les fonds propres et la liquidité."],
        source=RAPPORT + ", diapositives 4 à 8, 10, 11 et 20")
contenu({"titre": "Limites des conclusions", "sous_titre": "Les limites changent la précision des montants, pas le sens des verdicts",
         "synthese": "Trois données manquantes bloquent les calculs actuariels centraux"},
        "Être transparent : ces limites expliquent le statut « partiellement démontré ». Transition : « D'où le travail qui reste à faire. »",
        tableau={"titre": None, "entetes": ["Limite", "Effet sur les conclusions"], "largeurs": [12.6, 18.6], "taille": 10, "sans_code": True,
                 "lignes": [["Base des prestations datée, dates de naissance et fichier des ayants droit non reçus", "Ni provision dossier par dossier, ni mortalité par âge, ni engagement santé des retraités"],
                            ["Ni compte individuel d'épargne ni taux servi arrêté", "L'engagement d'épargne est mesuré globalement, pas épargnant par épargnant"],
                            ["Actifs retenus à leur valeur comptable, sans expertise de la clinique ni des terrains", "L'adossement est une estimation favorable au FPPN"],
                            ["Ni courbe de taux, ni échéancier, ni taux servi fiable", "Duration, valeur à risque et modèle stochastique ne sont pas applicables"],
                            ["Taux de cotisation maladie de 3,77 % mesuré sur 8 mois (04 à 10/2024 et 01/2025)", "Le taux ne doit pas être étendu aux autres mois sans vérification"],
                            ["Constats reposant sur des déclarations du FPPN", "Surlignés en jaune dans le rapport, confirmés ou retirés à réception des pièces"],
                            ["PDF du Code CIMA et du référentiel des SFD illisibles", "Seules les pages de la circulaire n° 00230 ont pu être lues"]]},
        source=RAPPORT + ", diapositives 54, 241 et 242 ; annexe « Méthodes non retenues »")
contenu({"titre": "Reste à faire", "sous_titre": "Il reste surtout des calculs qui attendent les données du FPPN, puis la finalisation du rapport",
         "synthese": "Chaque calcul restant est lié à une donnée précise de la liste des demandes"},
        "Transition : « Sans attendre ces données, les premières conclusions permettent déjà de recommander des actions. »",
        tableau={"titre": None, "entetes": ["Calcul restant", "Donnée attendue du FPPN"], "largeurs": [15.6, 15.6], "taille": 11, "sans_code": True,
                 "lignes": [["Engagements par âge et par épargnant", "Dates de naissance, comptes individuels d'épargne"],
                            ["Provisions maladie et décès dossier par dossier, boni et mali par exercice", "Base des prestations datée, liste des dossiers non payés"],
                            ["Adossement par échéance", "Inventaire et échéancier signés des placements"],
                            ["Projection de l'équilibre technique par garantie", "Résultats des deux calculs précédents"],
                            ["Onze indicateurs de gouvernance", "Pièces de gouvernance demandées"],
                            ["Rapport définitif", "Commentaires du FPPN après la restitution"]]},
        source=RAPPORT + ", diapositive 242 ; FPPN_Demandes_au_FPPN_20261023.xlsx")

# ---------------------------------------------------------------- 4. recommandations, avancement et suite
separateur(4, "Actions et suite", "Les recommandations, le taux d'avancement, les prochaines étapes et les points à arbitrer")
contenu({"titre": "Recommandations", "sous_titre": "Neuf recommandations, chacune liée à un constat, dont quatre à engager dans les six mois",
         "synthese": "Quatre actions de court terme conditionnent toutes les autres"},
        "« La matrice de criticité (diapositive 12) croise la gravité de chaque constat et l'horizon de l'action. La feuille de route (diapositive 27) enchaîne quatre étapes : collecter, calculer, vérifier, projeter. » Transition : « Où en est-on aujourd'hui ? »",
        tableau={"titre": None, "entetes": ["Recommandation", "Horizon"], "largeurs": [22.4, 8.8], "taille": 11,
                 "lignes": [["1. Calculer les provisions maladie et décès à chaque clôture", "Court terme"],
                            ["2. Arrêter le taux servi sur l'épargne et tenir les comptes individuels", "Court terme"],
                            ["3. Transmettre les données qui manquent aux calculs actuariels", "Court terme"],
                            ["4. Présenter à l'AIRMS un plan de mise en conformité prudentielle", "Court terme"],
                            ["5. Constituer une réserve de liquidité et adopter une politique de placement", "Moyen terme"],
                            ["6. Rééquilibrer la garantie maladie", "Moyen terme"],
                            ["7. Encadrer les investissements dans la clinique, le Grenier et le projet de Cocody", "Moyen terme"],
                            ["8. Faire réaliser l'étude actuarielle des risques longs et certifier les comptes", "Long terme"],
                            ["9. Mettre en place un dispositif actuariel permanent", "Long terme"]]},
        source=RAPPORT + ", diapositives 25 et 26")
contenu({"titre": "Taux d'avancement", "sous_titre": "L'avancement est bon pour la partie qui dépend d'Adenka, mais à risque pour celle qui dépend du FPPN",
         "synthese": "63,85 % des travaux sont réalisés au 23/10/2026, après 25,71 % du calendrier de la mission"},
        "Montrer l'onglet « Synthèse-PT_Actuariat » du plan de travail si le Chief veut le détail. Transition : « Pour tenir le calendrier, voici les prochaines étapes. »",
        graphique={"titre": "Taux de réalisation par phase au 23/10/2026 (%)", "type": "barres_h",
                   "categories": ["1. Cadrage", "2. Données", "3. Engagements", "4. Provisions", "5. Adossement et placements", "6. Ratios", "7. Projection", "8. Rapport", "Ensemble"],
                   "series": [{"nom": "Taux de réalisation", "valeurs": [87.50, 73.44, 54.37, 71.50, 75.62, 87.50, 40.00, 42.00, 63.85], "couleur": "112A40"}],
                   "format": "0.00", "legende": None, "axe_min": 0, "axe_max": 100},
        analyse=["**Pourquoi c'est un bon avancement.** La mission a démarré le 13/10/2026. Au 23/10/2026, 9 jours ouvrés sur 35 se sont écoulés jusqu'au rapport définitif du 30/11/2026, soit 25,71 % du calendrier.",
                 "**Pourquoi il reste un risque.** Le reste (36,15 %) se concentre dans les engagements (54,37 %) et la projection (40,00 %), qui attendent les données du FPPN, et dans le rapport (42,00 %), prévu en fin de mission.",
                 "**Ce que cela représente.** Environ 14,10 jours de travail sur 39 jours estimés, pour 20 jours ouvrés avant la restitution du 23/11/2026."],
        valable="plan de travail pointé au 23/10/2026, scénario « données dégradées » (39 jours)",
        source="Source : FPPN_Plan_de_travail_Actuariat_20261023.xlsx, onglet « Synthèse-PT_Actuariat », taux pondérés par le temps estimé de chaque tâche")
contenu({"titre": "Prochaines étapes", "sous_titre": "Le calendrier tient si le FPPN transmet ses données avant la mi-novembre",
         "synthese": "Restitution provisoire les 23 et 24/11/2026, rapport définitif le 30/11/2026"},
        "Clôture : « En résumé, le travail est avancé et chaque chiffre est vérifiable. Les conclusions sont déjà solides sur l'adossement et les placements. Les montants des engagements et des provisions seront précisés dès que le FPPN transmettra ses données. »",
        tableau={"titre": None, "entetes": ["Étape", "Qui", "Quand"], "largeurs": [19.4, 6.0, 5.8], "taille": 10, "sans_code": True,
                 "lignes": [["Envoyer au FPPN la liste des 26 questions et 84 documents, avec une date limite de réponse", "Adenka", "Dès validation interne"],
                            ["Relancer le FPPN sur la base des prestations datée, les dates de naissance, les comptes d'épargne et le taux servi", "Adenka et équipe d'audit", "Chaque semaine"],
                            ["Calculer les engagements et les provisions dossier par dossier, puis projeter l'équilibre technique", "Adenka", "À réception des données"],
                            ["Rapprocher les ratios et les placements avec les travaux de l'expert financier", "Adenka et expert financier", "Avant la restitution"],
                            ["Faire relire le rapport par une seconde personne", "Adenka", "Avant la restitution"],
                            ["Restitution provisoire au FPPN", "Adenka", "23 et 24/11/2026"],
                            ["Intégration des commentaires du FPPN", "Adenka", "Du 25 au 27/11/2026"],
                            ["Transmission du rapport définitif", "Adenka", "30/11/2026"]]},
        source="Source : plan de travail du 23/10/2026 ; rapport du 23/10/2026, diapositive 27")
contenu({"titre": "Points à arbitrer", "sous_titre": "Cinq décisions sont attendues avant la restitution au FPPN",
         "synthese": "Les arbitrages portent sur la formulation, le calendrier, le partage avec l'expert financier et le temps passé"},
        "Laisser ensuite le Chief choisir les sujets à approfondir, à l'aide de la carte du rapport en annexe.",
        tableau={"titre": None, "entetes": ["Point à arbitrer", "Pourquoi"], "largeurs": [10.4, 20.8], "taille": 10, "sans_code": True,
                 "lignes": [["Formulation des verdicts « insuffisant »", "L'enjeu de retrait d'agrément doit être présenté avec mesure avant la restitution au FPPN"],
                            ["Envoi de la liste des 26 questions et 84 documents", "La date limite de réponse doit être compatible avec la restitution du 23/11/2026"],
                            ["Date de restitution provisoire", "Si les données n'arrivent pas : maintenir la date avec des verdicts partiels, ou la décaler"],
                            ["Partage avec l'expert financier", "Le rapport calcule les 36 ratios et le rendement des placements pour apprécier les réserves. Les chiffres communs sont à rapprocher"],
                            ["Temps passé", "[[L'offre prévoit 14 hommes-jours pour le pôle 4 (p. 70). Le plan de travail estime 26 jours avec des données correctes et 39 jours avec des données dégradées. Le périmètre exact des 14 hommes-jours est à confirmer.]]"]]},
        source="Source : offre technique Adenka, pôle 4 (p. 70) ; plan de travail du 23/10/2026")

# ---------------------------------------------------------------- annexes
separateur("A", "Annexes", "Carte du rapport, questions probables et glossaire")
SL.append({"type": "annexe", "titre": "Carte du rapport", "notes": "Pour chaque partie du rapport : où elle se trouve, ce qu'elle démontre et la diapositive à ouvrir en cas de question.",
           "tableau": {"titre": None, "entetes": ["Partie", "Diapositives", "Ce qu'elle démontre", "À ouvrir"], "largeurs": [7.0, 3.0, 17.4, 3.8], "taille": 8.5, "sans_code": True,
                       "lignes": [["Résumé exécutif", "3 à 27", "Les réponses, les constats, les matrices et les recommandations", "4, 10, 23, 25"],
                                  ["1. Comprendre le FPPN", "28 à 40", "Garanties, textes, circuits des cotisations par produit", "32, 35 à 38"],
                                  ["2. Contexte et démarche", "41 à 45", "Mission, périmètre, documents reçus", "42, 44"],
                                  ["3. Données et qualité", "46 à 58", "Couverture des fichiers, contrôles, fichiers vides, qualité variable par variable", "48, 53, 54"],
                                  ["4. Effectifs et cotisations", "59 à 71", "Taux réellement prélevés (3,77 % et 1,23 % au lieu de 3,50 % et 1,50 %)", "61, 65"],
                                  ["5. Sinistralité et résultat", "72 à 86", "La sinistralité maladie passe de 6,00 % en 2021 à 73,53 % en 2025", "75"],
                                  ["6. Épargne des mutualistes", "87 à 99", "Intérêts dus non comptabilisés : 3 278 913 238 FCFA", "96, 98"],
                                  ["7. Prêts du Grenier", "100 à 108", "Encours de 1 044 302 107 FCFA, sans dépréciation", "102"],
                                  ["8. Retraités et annexes", "109 à 113", "Retraités précomptés par la CGRAE, flotte automobile, projets", "110"],
                                  ["9. Bilan, provisions et placements", "114 à 167", "Provisions recalculées, adossement, tests de résistance, placements", "123, 145, 155, 162"],
                                  ["10. Conformité prudentielle", "168 à 217", "Ratios par garantie, 36 ratios de l'UEMOA avec leur définition, sensibilités", "175, 178, 205"],
                                  ["11. Tables de mortalité", "218 à 224", "Choix des tables pour les calculs par âge", "220"],
                                  ["12. Projection de l'équilibre", "225 à 228", "Deux scénarios de 2026 à 2030", "227"],
                                  ["13. Outils et traçabilité", "229 à 239", "Traitement des données, extraits de code, contrôle des chiffres", "231, 239"],
                                  ["14. Analyses en attente", "240 à 244", "Données manquantes, questions au FPPN", "241, 243"],
                                  ["Annexes", "245 à 291", "Méthodes écartées, fichiers sources, formules, validité, abréviations", "246, 257"]]},
           "source": "Source : rapport actuariel et statistique du 23/10/2026 (291 diapositives)"})
QR = [["Pouvez-vous conclure sans les données manquantes ?", "Oui pour l'essentiel. Les conclusions reposent sur les balances, les grands livres, les précomptes et les états financiers. Les données manquantes préciseront les montants, pas le sens des verdicts."],
      ["Comment être sûr que les chiffres sont justes ?", "Par trois contrôles : 2 885 chiffres sur 2 885 retrouvés dans les calculs, chaque calcul refait par une formule Excel avec 0 écart, et une vérification indépendante qui a relevé 18 anomalies, toutes corrigées."],
      ["Pourquoi avoir fait plus que les quatre questions ?", "Parce que l'offre technique le prévoit : le pôle 4 demande de contrôler les données de base, d'apprécier les réserves au regard des exigences prudentielles et de projeter l'équilibre technique."],
      ["Le FPPN est-il en faillite ?", "Non. Il dégage des excédents chaque année depuis 2020, sauf sur les huit premiers mois de 2026. Son problème est la liquidité et la mesure de ses engagements d'épargne. Ses fonds propres sont négatifs (-879 498 700 FCFA au 31/08/2026)."],
      ["Les provisions sont supérieures au besoin : est-ce un problème ?", "Oui, parce qu'elles ne couvrent aucun risque identifié, alors que les intérêts dus sur l'épargne (3 278 913 238 FCFA) ne sont pas provisionnés. Le résultat et les fonds propres sont donc mal mesurés."],
      ["Qu'est-ce que l'adossement, et pourquoi est-il insuffisant ?", "C'est la capacité à payer chaque dette à son échéance avec des actifs disponibles à la même date. Au FPPN, l'épargne, exigible à tout moment, finance des actifs longs : clinique, terrains, prêts."],
      ["En quoi votre travail diffère-t-il de l'audit comptable ?", "L'audit comptable vérifie que les comptes sont exacts. Le volet actuariel vérifie que les engagements futurs sont bien mesurés et couverts. Les deux se complètent, et les chiffres partagés ont été rapprochés."],
      ["Pourquoi un rapport aussi long ?", "Pour que chaque conclusion soit démontrée et chaque chiffre retraçable jusqu'au fichier du FPPN. Le résumé exécutif (diapositives 3 à 27) se lit seul. Le reste sert de preuve."],
      ["Tiendrez-vous le délai ?", "Oui pour la partie qui dépend d'Adenka : 63,85 % des travaux sont réalisés après 25,71 % du calendrier. Le délai dépend de la date à laquelle le FPPN transmettra ses données."],
      ["Quels risques pour Adenka ?", "Conclure sur des déclarations non documentées : elles sont surlignées en jaune et seront confirmées ou retirées. Puis le délai, pour la raison ci-dessus."],
      ["Qu'attendez-vous du FPPN maintenant ?", "Les réponses aux 26 questions et les 84 documents de la liste, en priorité la base des prestations datée, les dates de naissance, les comptes individuels d'épargne et la décision fixant le taux servi."]]
for _k, _q in enumerate((QR[:6], QR[6:]), 1):
    SL.append({"type": "annexe", "titre": f"Questions probables et réponses ({_k}/2)", "notes": "Réponses courtes, à adapter.",
               "tableau": {"titre": None, "entetes": ["Question", "Réponse"], "largeurs": [9.4, 21.8], "taille": 9.5, "sans_code": True, "lignes": _q},
               "source": "Source : rapport actuariel et statistique du 23/10/2026"})
SL.append({"type": "annexe", "titre": "Glossaire simple", "notes": "",
           "tableau": {"titre": None, "entetes": ["Terme", "Explication en une phrase"], "largeurs": [7.0, 24.2], "taille": 8.5, "sans_code": True,
                       "lignes": [["Actuariat", "Discipline qui mesure les engagements futurs (prestations, épargne, retraites) à partir des probabilités et du temps."],
                                  ["Adossement actif-passif", "Capacité à payer chaque dette à son échéance avec des actifs disponibles à la même date."],
                                  ["Ayant droit", "Conjoint ou enfant d'un adhérent, couvert par les garanties du FPPN."],
                                  ["Engagement", "Somme que le FPPN devra payer plus tard à ses adhérents (épargne, intérêts, prestations)."],
                                  ["Épargne EFC", "Épargne prélevée chaque mois sur la solde et versée au policier à son départ à la retraite."],
                                  ["Épargne PEP", "Épargne volontaire des policiers, rémunérée."],
                                  ["Fonds propres", "Réserves et résultats cumulés. Négatifs, ils signifient que les pertes passées dépassent les réserves."],
                                  ["Lecture corrigée selon le PCMS", "Les mêmes comptes, avec l'épargne classée en ressource longue et les avances durables en emplois longs."],
                                  ["Précompte", "Retenue de la cotisation directement sur le salaire (Solde) ou sur la pension (CGRAE)."],
                                  ["Provision", "Montant mis de côté à la clôture pour payer des prestations déjà nées mais pas encore payées."],
                                  ["Ratios prudentiels", "Les 36 indicateurs de la Décision n° 05/2022 de l'UEMOA que doivent respecter les mutuelles sociales."],
                                  ["Sinistralité (S/C)", "Part des cotisations consommée par les prestations. La norme de l'UEMOA est comprise entre 70,00 % et 80,00 %."],
                                  ["Table de mortalité", "Tableau qui donne la probabilité de décéder à chaque âge, nécessaire aux calculs par âge."],
                                  ["Test de résistance", "Simulation d'un choc défavorable (retraits massifs, hausse des prestations) pour voir si le FPPN tient."]]},
           "source": "Source : Adenka"})
SL.append({"type": "merci", "notes": ""})
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositives')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "deck_point_chief_fppn (assemblé)", "exec"), g)
