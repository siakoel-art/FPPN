# ---- v13 : matrices du rapport (risques, criticité des constats, conformité, données et analyses)
ROUGE_M, JAUNE_M, VERT_M, ORANGE_M = "C00000", "FFD966", "70AD47", "EFAC62"


def _grav(montant):
    """Gravité d'un risque chiffré, selon son impact en FCFA."""
    m_ = abs(montant)
    return 1 if m_ < 50e6 else 2 if m_ < 500e6 else 3 if m_ < 2e9 else 4 if m_ <= 5e9 else 5


def _zone(p_, g_):
    sc = p_ * g_
    return ROUGE_M if sc >= 15 else JAUNE_M if sc >= 8 else VERT_M


_INT_AN = (INT55_25 - interets_dus(0.035, 0.055, 2024)) - next(x["produits_financiers"] for x in V11["rendement_gl"] if x["exercice"] == 2025)


def _risques():
    return [  # (code, risque, probabilité, justification, gravité, justification)
        ("R1", "Manque de liquidités face aux sorties d'épargne", 4, f"Sorties observées de {pc('RQ ts', T_MOY)} de l'épargne par an. L'actif à moins d'un an couvre {pc('RQ couv', couverture('31/08/2026', T_MOY)[2])} des paiements de l'année",
         _grav(GAP[4][5]), f"Manque cumulé de {F('RQ gap', -GAP[4][5])} à cinq ans"),
        ("R2", "Retrait d'agrément ou sanctions de l'AIRMS", 4, f"{NB_NC} indicateurs non conformes deux exercices de suite : la condition de l'art. 84 du règlement n° 07/2009 est remplie, la sanction n'est pas automatique",
         5, "Arrêt de l'activité du FPPN"),
        ("R3", "Dette d'épargne sous-évaluée", 5, "Constaté : aucun intérêt comptabilisé de 2020 au 31/08/2026",
         _grav(INT55), f"Intérêts dus non comptabilisés : {F('RQ int', INT55)}"),
        ("R4", "Provisions sans lien avec les risques", 5, "Constaté : montants identiques depuis le 01/09/2020",
         _grav(PROV_MAL + PROV_OBS - _BESOIN), f"Provisions supérieures au besoin mesuré de {F('RQ exc', PROV_MAL + PROV_OBS - _BESOIN)}"),
        ("R5", "Déficit de la garantie maladie", 5, f"Constaté : ratio combiné de {pc('RQ rc25', _RGm[(2025, 'Maladie')]['ratio_combine'])} en 2025 et {pc('RQ rc26', _RGm[(2026, 'Maladie')]['ratio_combine'])} sur huit mois de 2026",
         _grav(_RGm[(2026, "Maladie")]["resultat_technique"]), f"Résultat technique maladie de {F('RQ rt26', _RGm[(2026, 'Maladie')]['resultat_technique'])} sur huit mois de 2026"),
        ("R6", "Absence de coussin de fonds propres", 5, "Constaté : fonds propres négatifs à chaque clôture depuis 2020",
         _grav(b[2026]["fonds_propres"]), f"Fonds propres de {F('RQ fp', b[2026]['fonds_propres'])} au 31/08/2026"),
        ("R7", "Pertes sur les prêts du Grenier", 5, "Constaté : prêts en retard au 31/08/2026, aucune dépréciation passée",
         _grav(_GR26), f"Dépréciation nécessaire de {F('RQ dep', _GR26)}"),
        ("R8", "Produits financiers inférieurs aux intérêts dus", 5, f"Constaté : rendement moyen de {pc('RQ rdm', V11['rendement_moyen_2021_2025'])} par an de 2021 à 2025",
         _grav(_INT_AN), f"Intérêts dus de 2025 supérieurs aux produits financiers de 2025 de {F('RQ int an', _INT_AN)}"),
        ("R9", "Engagements impossibles à mesurer faute de données", 5, "Constaté : ni base des prestations datée, ni dates de naissance des mutualistes",
         3, "Non chiffrable : bloque les provisions dossier par dossier, la mortalité, la santé des retraités et la projection"),
        ("R10", "Placements décidés sans délibération des organes", 4, "[[Le directeur général décide seul des placements]]. Aucune délibération reçue",
         _grav(500000000), f"Placement dans le FCP de {F('RQ fcp', 500000000)} en 2025"),
    ]


# les justifications chiffrées s'affichent sur la diapositive qui suit la matrice
COURANTE = len(SL) + 2
RISQUES = _risques()
COURANTE = None
_ETIQ_P = {5: "5 - Très probable", 4: "4 - Probable", 3: "3 - Possible", 2: "2 - Peu probable", 1: "1 - Très peu probable"}
_ETIQ_G = ["1 - Négligeable", "2 - Mineure", "3 - Modérée", "4 - Majeure", "5 - Catastrophique"]
_MAT_R, _FONDS_R = [], {}
for i_, p_ in enumerate((5, 4, 3, 2, 1)):
    ligne_ = [_ETIQ_P[p_]]
    for g_ in range(1, 6):
        ligne_.append(", ".join(c_ for c_, _, pp_, _, gg_, _ in RISQUES if pp_ == p_ and gg_ == g_))
        _FONDS_R[(i_, g_)] = _zone(p_, g_)
    _MAT_R.append(ligne_)
LEG_RISQ = [("🔴", "Élevé", "score de 15 à 25, traitement immédiat"), ("🟡", "Moyen", "score de 8 à 12, plan d'action"), ("🟢", "Faible", "score de 1 à 6, surveillance")]
contenu({"titre": "Matrice des risques", "sous_titre": f"{sum(1 for r in RISQUES if r[2] * r[4] >= 15)} risques sur {len(RISQUES)} se situent dans la zone rouge",
         "synthese": "Probabilité en lignes, gravité en colonnes ; chaque risque est justifié sur la diapositive suivante"},
        "Score = probabilité × gravité. Zone rouge : score de 15 à 25. Zone jaune : de 8 à 12. Zone verte : de 1 à 6. La gravité d'un risque chiffré suit son impact en FCFA ; la probabilité suit ce qui est constaté dans les comptes et les fichiers.",
        tableau={"titre": tab("Matrice des risques : probabilité et gravité"), "entetes": ["Probabilité \\ Gravité"] + _ETIQ_G,
                 "largeurs": [5.2] * 6, "taille": 10, "sans_code": True, "hauteur_ligne": 1.15, "fonds": _FONDS_R, "centre": True, "legende": LEG_RISQ, "lignes": _MAT_R},
        disposition="dessous",
        encadre=["**Gravité** (impact en FCFA) : 1 négligeable, moins de 50 000 000 FCFA. 2 mineure, de 50 000 000 à 500 000 000 FCFA. 3 modérée, de 500 000 000 à 2 000 000 000 FCFA. 4 majeure, de 2 000 000 000 à 5 000 000 000 FCFA. 5 catastrophique, plus de 5 000 000 000 FCFA ou arrêt de l'activité.",
                 "**Probabilité** : 5 très probable, déjà constaté dans les comptes ou les fichiers. 4 probable, conditions réunies mais risque non encore réalisé. 3 possible. 2 peu probable. 1 très peu probable."],
        encadre_titre="Échelles",
        valable="comptes et fichiers de 2020 au 31/08/2026 ; impacts chiffrés dans le corps du rapport",
        source="Source : analyses Adenka, parties 5 à 10 du rapport")
contenu({"titre": "Matrice des risques : justification", "sous_titre": "Chaque note de probabilité et de gravité repose sur un chiffre du rapport",
         "synthese": "Les risques constatés sont notés 5 en probabilité ; la gravité suit l'impact chiffré en FCFA"},
        "Justification de la place de chaque risque dans la matrice. Niveau : 🔴 élevé (score de 15 à 25), 🟡 moyen (8 à 12), 🟢 faible (1 à 6).",
        tableau={"titre": tab("Risques, probabilité, gravité et score"), "entetes": ["N°", "Risque", "Probabilité et justification", "Gravité et justification", "Score"],
                 "largeurs": [1.2, 6.4, 10.2, 10.2, 3.2], "taille": 7.5, "sans_code": True, "legende": LEG_RISQ,
                 "lignes": [[c_, r_, f"{p_}. {jp_}", f"{g_}. {jg_}", bv(str(p_ * g_), "🔴" if p_ * g_ >= 15 else "🟡" if p_ * g_ >= 8 else "🟢")] for c_, r_, p_, jp_, g_, jg_ in RISQUES]},
        disposition="dessous",
        valable="comptes et fichiers de 2020 au 31/08/2026",
        source="Source : analyses Adenka ; chiffres des parties 5 à 10 du rapport")
