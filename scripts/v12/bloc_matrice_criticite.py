# ---- v13 : matrice de criticité des constats (criticité × horizon de la recommandation qui y répond)
ROUGE_M, JAUNE_M, VERT_M, ORANGE_M = "C00000", "FFD966", "70AD47", "EFAC62"
_CRIT_L = ["Critique", "Élevé", "Moyen", "Faible"]
_HOR_C = ["Court terme (moins de 6 mois)", "Moyen terme (6 à 18 mois)", "Long terme (plus de 18 mois)"]
_PLACE = {("Critique", 0): "Provisions maladie et décès (rec. 1)\nÉpargne PEP et EFC (rec. 2)\nConformité prudentielle (rec. 4)",
          ("Critique", 1): "Adossement actif-passif (rec. 5)",
          ("Élevé", 0): "Fonds propres (rec. 4)\nDonnées (rec. 3)",
          ("Élevé", 1): "Sinistralité maladie (rec. 6)\nPlacements (rec. 5)\nGouvernance (rec. 7)",
          ("Moyen", 1): "Prêts du Grenier (rec. 7)",
          ("Faible", 2): "Personnel du FPPN (rec. 9)"}
_ZONES_C = {"Critique": [ROUGE_M, ROUGE_M, ORANGE_M], "Élevé": [ROUGE_M, ORANGE_M, JAUNE_M], "Moyen": [ORANGE_M, JAUNE_M, VERT_M], "Faible": [JAUNE_M, VERT_M, VERT_M]}
_BOULE_CRIT = {"Critique": "🔴", "Élevé": "🟠", "Moyen": "🟡", "Faible": "🟢"}
_MAT_C, _FONDS_C = [], {}
for i_, cr_ in enumerate(_CRIT_L):
    _MAT_C.append([_BOULE_CRIT[cr_] + " " + cr_] + [_PLACE.get((cr_, k_), "") for k_ in range(3)])
    for k_ in range(3):
        _FONDS_C[(i_, k_ + 1)] = _ZONES_C[cr_][k_]
contenu({"titre": "Matrice de criticité des constats", "sous_titre": "Trois des quatre constats critiques se traitent dans les six mois",
         "synthese": "Criticité du constat en lignes, horizon de la recommandation qui y répond en colonnes"},
        "Chaque constat de la synthèse est placé selon sa criticité et l'horizon de la recommandation qui le traite. La couleur de la case donne l'ordre de traitement.",
        tableau={"titre": tab("Matrice de criticité des constats et horizon des recommandations"), "entetes": ["Criticité \\ Horizon"] + _HOR_C,
                 "largeurs": [5.2, 8.67, 8.67, 8.66], "taille": 9.5, "sans_code": True, "hauteur_ligne": 1.6, "fonds": _FONDS_C, "centre": True,
                 "legende": [("🔴", "Priorité 1", "à engager immédiatement"), ("🟠", "Priorité 2", "à engager dans l'année"), ("🟡", "Priorité 3", "à planifier"), ("🟢", "Priorité 4", "à surveiller")],
                 "lignes": _MAT_C},
        disposition="dessous",
        analyse=["**Les provisions, l'épargne et la conformité prudentielle sont critiques et se traitent à court terme : ce sont les recommandations 1, 2 et 4.** L'adossement, critique lui aussi, demande une réserve de liquidité et une politique de placement qui se construisent sur 6 à 18 mois (recommandation 5)."],
        valable="constats de la synthèse et recommandations du rapport",
        source="Source : synthèse des constats et recommandations du rapport (analyses Adenka)")
