# ---- v11 : partie 14 : analyses en attente (méthodes retenues), questions au FPPN, documents prioritaires (PBC du 23/10/2026)
def _n(*anciens):
    return ", ".join(str(x) for x in sorted({PBCM[a] for a in anciens}))


contenu({"titre": "Analyses en attente", "sous_titre": "Quatre analyses attendent des données du FPPN, toutes avec une méthode déjà choisie",
         "synthese": "Les conclusions du rapport sont établies sur les documents reçus ; ces analyses les préciseront sans en changer le sens"},
        "Chaque analyse en attente utilise la méthode retenue dans le rapport, pas une méthode écartée. Les numéros renvoient à la PBC du 23/10/2026.",
        tableau={"titre": tab("Analyses en attente, méthode retenue et données nécessaires"), "entetes": ["Analyse", "Méthode retenue", "Ce que l'analyse précisera", "Données nécessaires (PBC n°)"],
                 "largeurs": [5.6, 9.4, 8.4, 7.8], "taille": 9,
                 "lignes": [["Provision maladie", "Soins de l'année facturés après la clôture, sur la base des factures datées (date des soins, du dépôt et du paiement)", f"Montant exact, aujourd'hui borné à {F('AT psap', psap_m_elargie(2024))} au 31/12/2024", "Base des prestations datée (" + _n(71) + ")"],
                            ["Provision décès", "Dossiers connus au compte 41591000 et décès survenus déclarés après la clôture", "Part des dossiers forclos dans la dette de " + F("AT dette", DETTE_DEC26), "Liste des dossiers du 41591000, fichiers des décès (" + _n(115, 92) + ")"],
                            ["Mortalité observée et santé des retraités", "Décès observés comparés aux décès attendus des tables CIMA H et F, par âge", "Écart de mortalité des policiers, engagement de santé à vie des retraités", "Dates de naissance des mutualistes et des retraités (" + _n(9, 120) + ")"],
                            ["Engagement d'épargne par épargnant", "Solde de chaque compte individuel augmenté des intérêts au taux arrêté", "Échéancier des sorties, à la place du taux de sortie observé", "Comptes individuels, taux servi (" + _n(17, 99) + ")"]]},
        disposition="dessous",
        analyse=["**Aucune conclusion du rapport ne dépend de ces analyses pour son sens.** Les provisions du bilan dépassent les besoins mesurés, le non-adossement est démontré et les placements sont insuffisants, quelles que soient les valeurs exactes."],
        implication="Les données listées sont demandées au FPPN avec la priorité « Urgent ». Leur réception permettra une mise à jour chiffrée des diapositives concernées.",
        valable="analyses en attente au 23/10/2026",
        source="Source : PBC actuarielle du 23/10/2026 ; méthodes de la partie 9 et de l'annexe des méthodes non retenues")
QFPPN = [
    ["Cotisations", "Pourquoi la Solde précompte-t-elle 3,77 % et 1,23 % du salaire de base, alors que les statuts (art. 17) fixent 3,5 % et 1,5 % ?", _n(56), "Urgent"],
    ["Cotisations", "Le délai de trois mois de cotisation avant les prestations s'applique-t-il à la maladie et au décès ? L'entretien du 16/10/2026 indiquait l'absence de délai de carence.", _n(59), "Urgent"],
    ["Cotisations", "Quelle décision a porté l'EFC de 6 000 FCFA à 10 000 FCFA en 2021 ?", _n(58), "Moyen"],
    ["Cotisations", "Que représentent les références PTI, RPEP, EP, TI, R-PEP et PTQ précomptées au poste 998 ?", "85", "Moyen"],
    ["Cotisations", "Le FPPN rapproche-t-il les fichiers aller et retour de chaque produit ? Si oui, qui le fait et selon quelle procédure ?", "88", "Faible"],
    ["Épargne", "Quel taux est servi sur l'EFC (3,5 %, 5,5 % ou 14 %), et quelle décision signée le fixe ?", _n(99), "Urgent"],
    ["Épargne", "Le DAT du compte 52210200 est-il un placement du FPPN ou un placement affecté à l'EFC ?", _n(80), "Urgent"],
    ["Provisions", "Quelle décision permettra de reprendre ou de justifier les provisions passées vers 2004 ?", _n(72), "Urgent"],
    ["Provisions", "Pourquoi les sinistres décès 2021 valent-ils 156 020 000 FCFA dans l'étude des ratios et 401 856 666 FCFA dans la balance 2021 ?", "86", "Moyen"],
    ["Provisions", "À quel grade correspond l'allocation de 2 200 000 FCFA des états décès, et quel texte fonde l'exclusion du suicide ?", _n(116), "Moyen"],
    ["Retraités", "La garantie santé des retraités est-elle à vie ou limitée, et quelle est leur cotisation ?", _n(79), "Urgent"],
    ["Personnel", "Quel âge de départ le FPPN applique-t-il à son personnel civil (cinq agents de 55 ans ou plus en poste) ?", _n(113), "Urgent"],
    ["Personnel", "Pourquoi douze dates d'embauche diffèrent-elles entre les DISA 2021 à 2023 et la DISA 2025 ?", _n(113), "Urgent"],
    ["Personnel", "À quelle entité (Direction Générale, clinique, Grenier) sont rattachés les dix-neuf agents recrutés en 2025 ?", _n(113), "Urgent"],
    ["Placements", "Quelle est la nature de la créance de 3 407 519 850 FCFA sur l'hôpital de la Police, et un remboursement est-il prévu ?", _n(87), "Moyen"],
    ["Placements", "Le compte 10620000 enregistre-t-il la réévaluation du terrain du siège ou de matériels, comme l'indique son intitulé ?", _n(86), "Urgent"],
    ["Placements", "Quel est le budget du projet de Cocody, son financement et la part financée par l'épargne ?", _n(118), "Urgent"],
    ["Grenier", "Pourquoi l'essentiel de l'encours était-il en retard fin janvier 2026, et comment sont réglées les échéances non précomptées ?", "90", "Moyen"],
    ["Flotte automobile", "Sur quelle assiette la commission de 20 % versée par GNA est-elle calculée ?", _n(89), "Moyen"],
    ["MEDOC", "Sur quelle période le produit MEDOC a-t-il fonctionné, et quand a-t-il été suspendu ?", _n(69), "Faible"],
    ["Gouvernance", "Quelle version des comptes 2023 a été approuvée, et pourquoi le résultat au 31/08/2026 diffère-t-il entre états financiers et balance ?", _n(110), "Urgent"],
    ["Gouvernance", "Le directeur général a-t-il été désigné ou prolongé après 2023, et qui autorise les placements ?", _n(117), "Urgent"],
    ["Conformité", "Quels sont le numéro, la date et les risques autorisés de l'agrément AIRMS ?", _n(44), "Urgent"],
    ["Conformité", "Le FPPN adhère-t-il au FOMUS-CI, et quel montant a-t-il payé depuis 2022 ?", _n(49), "Moyen"],
    ["Conformité", "Une étude actuarielle des risques longs ou une certification des comptes a-t-elle été réalisée ?", _n(48, 47), "Moyen"],
    ["Conformité", "Quel est l'effectif total de la Police nationale au 31/12 de 2020 à 2025 et au 31/08/2026 ?", "89", "Moyen"],
]
_moit = (len(QFPPN) + 1) // 2
for _k, _lig in enumerate((QFPPN[:_moit], QFPPN[_moit:]), 1):
    contenu({"titre": f"Questions au FPPN ({_k}/2)", "sous_titre": f"{len(QFPPN)} questions simples, chacune reliée à un constat du rapport et à la PBC",
             "synthese": "Réponse écrite attendue avec la pièce qui l'établit ; liste complète des documents dans la PBC du 23/10/2026"},
            "Questions transmises au FPPN dans le fichier « FPPN_Demandes_au_FPPN_20261023.xlsx », avec une colonne de réponse.",
            tableau={"titre": tab(f"Questions au FPPN ({_k}/2)"), "entetes": ["N°", "Thème", "Question", "PBC n°", "Priorité"], "largeurs": [1.0, 3.6, 20.0, 3.6, 3.0], "taille": 8.5,
                     "lignes": [[str(i), *row] for i, row in enumerate(_lig, 1 + (_k - 1) * _moit)]},
            disposition="dessous",
            valable="questions ouvertes au 23/10/2026",
            source="Source : analyses Adenka ; PBC actuarielle du 23/10/2026 ; FPPN_Demandes_au_FPPN_20261023.xlsx")
