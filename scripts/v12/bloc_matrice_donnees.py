# ---- v13 : matrice données × analyses (statut de chaque donnée dans la PBC du 23/10/2026)
_DON = [("Base des prestations datée", 47, "🟠"), ("Dossiers non payés", 7, "🟠"), ("Fichiers des décès", 64, "🟠"), ("Dates de naissance des mutualistes", 4, "🟠"),
        ("Dates de naissance des retraités", 84, "🔴"), ("Fichier des ayants droit", 43, "🔴"), ("Comptes individuels d'épargne", 9, "🟠"),
        ("Taux servi arrêté", 69, "🟠"), ("Inventaire des placements", 55, "🟠")]
_BES = {  # analyse : données nécessaires (indices de _DON)
    "Provision maladie dossier par dossier": (0, 1),
    "Provision décès dossier par dossier": (1, 2),
    "Mortalité observée par âge": (2, 3),
    "Engagement santé des retraités": (0, 4),
    "Engagement d'épargne par épargnant": (6, 7),
    "Adossement par échéance": (6, 8),
    "Projection par garantie": (0, 3, 5),
}
_MAT_D = [[an_] + [_DON[k][2] if k in ks_ else "" for k in range(len(_DON))] for an_, ks_ in _BES.items()]
contenu({"titre": "Matrice des données et des analyses", "sous_titre": "Aucune analyse actuarielle centrale ne dispose de toutes ses données",
         "synthese": "Analyses en lignes, données nécessaires en colonnes, avec leur statut dans la PBC du 23/10/2026"},
        "Une case vide : la donnée n'est pas nécessaire à l'analyse. Une boule : la donnée est nécessaire, et la couleur donne son statut dans la PBC.",
        tableau={"titre": tab("Données nécessaires à chaque analyse et statut dans la PBC"), "entetes": ["Analyse"] + [f"{n_}\nPBC n° {p_}" for n_, p_, _ in _DON],
                 "largeurs": [5.6] + [2.84] * 9, "taille": 8, "sans_code": True, "centre": True, "hauteur_ligne": 0.85,
                 "legende": [("🔴", "Non reçu", "donnée non reçue"), ("🟠", "Partiel", "donnée reçue en partie"), ("🟢", "Reçu", "donnée reçue")], "lignes": _MAT_D},
        disposition="dessous",
        analyse=["**La base des prestations datée et les dates de naissance interviennent dans quatre des sept analyses.** Les dates de naissance des retraités et le fichier des ayants droit n'ont pas été reçus du tout."],
        valable="PBC pointée au 23/10/2026",
        source="Source : PBC du 23/10/2026 (« FPPN_PBC_complete_pointee_20261023.xlsx ») ; méthodes retenues du rapport")
