# ---- v9 : méthode de sinistralité du FPPN (communiquée le 22/10/2026) et mortalité brute (partie 5)
SRC_RATIO = "Source : ETUDE DES RATIO MALADIE DECES 2.xlsx (feuilles Tableau, Notes et Historique 2015-2021) et INTERPRETATION DES RATIOS MALADIE MATERNITE ET DECES INVALIDITE.pdf, reçus le 22/10/2026"
_sin26 = RT["SIN"]["valeurs"].get(2026)
contenu({"titre": "Méthode de sinistralité du FPPN", "sous_titre": "Le FPPN rapporte les prestations payées aux cotisations payées, sans provisions, garantie par garantie",
         "synthese": "Au 30/06/2026, le FPPN calcule 12,95 % en décès, 99,23 % en maladie et 78,49 % pour l'ensemble ; sa propre note écrit aussi 19,66 % et 107 %"},
        "La méthode du FPPN figure dans la feuille Notes de son étude des ratios. Elle est comparée à la formule de la Décision n° 05/2022. Les deux méthodes répondent à la même question (part des cotisations consommée par les prestations), mais pas avec les mêmes montants.",
        tableau={"titre": tab("Méthode du FPPN et méthode réglementaire au 30/06/2026"), "entetes": ["Élément", "Méthode du FPPN", "Méthode réglementaire (Décision n° 05/2022)"], "largeurs": [5.2, 13.0, 13.0], "taille": 8.5,
                 "lignes": [["Numérateur", "Prestations payées : allocations décès et frais funéraires au barème ; prestations maladie de l'état des hospitalisations et consultations", "Charges de prestations (comptes 60) et variation des provisions techniques (6901 - 7901)"],
                            ["Dénominateur", "Cotisations payées : policiers en activité, civils de la Direction Générale, retraités", "Cotisations acquises (comptes 70) et variation des provisions sur cotisations (7906 - 6906)"],
                            ["Décès", "55 600 000 / 429 215 718 = 12,95 % (« TSD = 13 % »), la même note écrit aussi « TS = 19, 66 % »", "Pas de taux par garantie dans le texte ; taux par garantie du rapport : partie 5"],
                            ["Maladie", "1 345 906 033 / 1 356 344 706 = 99,23 % (« TSM »), la même note écrit aussi « TS = 107 % »", "Pas de taux par garantie dans le texte"],
                            ["Ensemble", "1 401 506 033 / 1 785 560 424 = 78,49 % (« TSP »)", "Taux global au 31/08/2026 : " + (pc("SIN 2026 (méthode)", _sin26) if _sin26 is not None else "non calculé") + " (balance provisoire, huit mois)"],
                            ["Norme", "Le FPPN compare chaque taux à 70 % à 80 %", "70 % à 80 % pour le seul taux global"]]},
        disposition="dessous",
        analyse=["**La méthode du FPPN ignore les provisions et compare chaque garantie à la norme globale.** La norme de 70 % à 80 % de la Décision s'applique au seul taux global, provisions comprises.",
                 "**La note du FPPN contient deux taux différents pour le décès et pour la maladie.** Les quotients écrits donnent 12,95 % et 99,23 % ; les mentions 19,66 % et 107 % ne correspondent à aucun calcul affiché."],
        implication="La méthode du FPPN est désormais connue : le registre des formules (onglet 11_formules) la reprend. Elle convient au pilotage par garantie, pas à la conformité, qui exige la formule du texte avec les provisions.",
        valable="période du 01/01 au 30/06/2026 pour les taux du FPPN, du 01/01 au 31/08/2026 pour le taux réglementaire",
        source=SRC_RATIO + " ; " + DEC + ", § 1.2.1.2.5 ; " + SRC_COMPTA[9:])
_sc = V9["sinistralite_comparee"]
contenu({"titre": "Sinistralité décès : le FPPN et les comptes", "sous_titre": "De 2022 à 2025, les cotisations et les sinistres du FPPN sont ceux de la comptabilité",
         "synthese": "Le rapport sinistres / cotisations du FPPN égale le taux décès par garantie du rapport de 2022 à 2025 ; 2020 et 2021 diffèrent"},
        "Le tableau compare, année par année, les montants de l'étude du FPPN (feuille Tableau) et ceux des balances : cotisations décès (compte 7021) et prestations décès (comptes 6021 et 6024). En 2020, un assureur portait la garantie : la prime d'assurance décès (compte 65180000) remplace les prestations.",
        tableau={"titre": tab("Sinistralité décès selon le FPPN et selon les comptes"), "entetes": ["Année", "Cotisations (FPPN)", "Sinistres (FPPN)", "S/P du FPPN", "Cotisations décès (comptes)", "Prestations décès (comptes)", "S/C (comptes)"],
                 "largeurs": [2.6, 5.0, 4.8, 3.4, 5.4, 5.4, 3.4], "taille": 9,
                 "lignes": [[str(x["annee"]), F(f"FPPN cot déc {x['annee']}", x["fppn_cotisations"]), F(f"FPPN sin déc {x['annee']}", x["fppn_sinistres"]), pc(f"FPPN S/P {x['annee']}", x["fppn_sp"]),
                             F(f"Compta cot déc {x['annee']}", x["compta_cot_deces"]), F(f"Compta prest déc {x['annee']}", x["compta_prest_deces"] + x["compta_prime_deces"]),
                             pc(f"Compta S/C déc {x['annee']}", x["compta_sc_deces_avec_prime"])] for x in _sc]},
        disposition="dessous",
        analyse=["**De 2022 à 2025, les deux sources donnent les mêmes cotisations, les mêmes sinistres et donc le même taux.** Le taux décès par garantie du rapport est ainsi confirmé par le FPPN.",
                 f"**En 2021, le FPPN retient {F('FPPN sin 2021 (t)', next(x['fppn_sinistres'] for x in _sc if x['annee'] == 2021))} de sinistres, contre {F('Compta prest 2021 (t)', next(x['compta_prest_deces'] for x in _sc if x['annee'] == 2021))} dans les comptes.** En 2020, ses cotisations et ses sinistres diffèrent aussi des comptes, l'assureur portant alors le risque."],
        implication="Pour 2022 à 2025, la garantie décès consomme moins d'un quart de ses cotisations selon les deux sources. L'écart de 2021 entre les deux sources est à expliquer par le FPPN.",
        valable="garantie décès-invalidité, exercices 2020 à 2025",
        source=SRC_RATIO + " ; " + SRC_COMPTA[9:])
_mb = V9["mortalite_brute"]
_m25 = next(x for x in _mb if x["annee"] == 2025)
contenu({"titre": "Mortalité brute des membres participants", "sous_titre": "Entre trois et cinq décès pour mille adhérents chaque année depuis 2021",
         "synthese": f"En 2025 : {N('Décès adh 2025', _m25['deces_adherents'])} décès pour {N('Adh 2025', _m25['adherents'])} adhérents, soit {N('Taux brut 2025', 1000 * _m25['taux_brut'], 2)} pour mille"},
        "Le taux brut rapporte les décès de membres participants de l'année au nombre d'adhérents retenu par le FPPN. Il ne tient pas compte de l'âge : faute de dates de naissance des mutualistes, il ne peut pas être comparé aux décès attendus par la table CIMA H (rapport A/E de la partie 12).",
        tableau={"titre": tab("Décès et taux brut de mortalité des membres participants"), "entetes": ["Année", "Adhérents", "Décès de membres", "Décès de conjoints", "Décès d'enfants", "Taux brut (pour mille)"],
                 "largeurs": [3.0, 5.0, 5.0, 5.0, 5.0, 5.4], "taille": 9,
                 "lignes": [[str(x["annee"]), N(f"Adh {x['annee']}", x["adherents"]), N(f"Déc adh {x['annee']}", x["deces_adherents"]), N(f"Déc conj {x['annee']}", x["deces_conjoints"]),
                             N(f"Déc enf {x['annee']}", x["deces_enfants"]), N(f"Taux brut {x['annee']}", 1000 * x["taux_brut"], 2)] for x in _mb]},
        disposition="dessous",
        analyse=[f"**Le taux brut varie de {N('Taux brut min', 1000 * min(x['taux_brut'] for x in _mb if x['annee'] >= 2021), 2)} à {N('Taux brut max', 1000 * max(x['taux_brut'] for x in _mb if x['annee'] >= 2021), 2)} pour mille de 2021 à 2025.** Le FPPN reporte le même nombre d'adhérents trois années de suite de 2016 à 2018, puis en 2021 et 2022, ce qui fragilise les taux de ces années.",
                 f"**Les décès d'ayants droit (conjoints et enfants) passent de {N('Déc AD 2020', next(x['deces_conjoints'] + x['deces_enfants'] for x in _mb if x['annee'] == 2020))} en 2020 à {N('Déc AD 2025', _m25['deces_conjoints'] + _m25['deces_enfants'])} en 2025.** Le nombre d'ayants droit couverts n'est pas connu : aucun taux ne peut être calculé pour eux."],
        implication="Le taux brut sert d'ordre de grandeur. La comparaison à la table CIMA H, nécessaire pour juger la cotisation décès, demande les dates de naissance des membres et la base des décès (PBC, items 9 à 11 et 28).",
        valable="membres participants recensés par le FPPN, années 2015 à 2025 ; ne vaut ni par âge, ni par grade, ni pour les ayants droit",
        source=SRC_RATIO)
