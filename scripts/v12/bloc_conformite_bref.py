# ---- v12 : conformité prudentielle en bref, matrice de 2020 au 31/08/2026 (lecture du FPPN, puis lecture corrigée selon le PCMS)
cles = ["SIN", "MSE", "ROB", "FEP", "LIQ", "EQS", "CHF", "RL1"]


def _nom_court(c):
    return RT[c]["nom"].replace(" (risques longs)", "").replace("Ratio de solvabilité de la branche épargne", "Solvabilité de l'épargne")


def _norme_courte(c):
    return RT[c]["norme"].replace("Au moins 1 (marge disponible au moins égale à la marge exigée)", "Au moins 1").replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "")


def _ligne_bref(c, corr=False):
    x = RT[c]
    sts = x.get("statuts_retraites" if corr else "statuts", {})
    vals = []
    for a in _AN:
        v_ = val_ratio(x, a, corr)
        vals.append(v_ if v_ == "n.d." else bv(v_, boule_statut(sts.get(a, ""))))
    if corr:
        glob = "Non conforme" if any(str(sts.get(a, "")).startswith("Non conforme") for a in range(2022, 2026)) else "Conforme"
        nom = _nom_court(c) + ", lecture corrigée selon le PCMS"
    else:
        glob, nom = x["statut_global"], _nom_court(c)
    return [nom, _norme_courte(c)] + vals + [glob]


_LB = [_ligne_bref(c) for c in cles] + [_ligne_bref(c, True) for c in ("MSE", "LIQ", "EQS")]
contenu({"titre": "Conformité prudentielle en bref", "sous_titre": "La réserve obligatoire, les fonds propres, la liquidité et la couverture de l'épargne ne sont respectées aucune année",
         "synthese": f"Sur 36 indicateurs de la Décision n° 05/2022 : {NB_NC} non conformes, {NB_C} conformes, {NB_ATT} en attente de données ou à apprécier"},
        "Le statut global est jugé sur les exercices 2022 à 2025, la Décision étant entrée en vigueur le 1er juin 2022. 2020 et 2021 sont indicatifs, comme les huit mois de 2026. Détail dans la partie 10.",
        tableau={"titre": tab("Huit ratios prudentiels clés, de 2020 au 31/08/2026"), "entetes": ["Ratio", "Norme"] + _ENT_AN + ["Statut 2022 à 2025"],
                 "largeurs": [7.0, 3.0] + [2.6] * 7 + [3.0], "statut_col": 9, "taille": 7, "lignes": _LB},
        disposition="dessous",
        encadre=["**Lecture du FPPN** : comptes tels qu'ils sont tenus. **Lecture corrigée selon le PCMS** : mêmes montants, avec l'épargne des adhérents (41195100, 41195200) en ressource longue et les avances durables (47111100 à 47111105, 47111201, 485) en emplois longs. Pour la marge, pertes reportées déduites et réévaluation (10620000) exclue."],
        encadre_titre="Deux lectures des comptes",
        analyse=["**Le classement de l'épargne change le verdict sur la liquidité, pas sur la solidité.** Corrigée selon le PCMS, la liquidité devient conforme en 2025 (" + val_ratio(RT["LIQ"], 2025, True) + " contre " + val_ratio(RT["LIQ"], 2025) + "). La réserve obligatoire, les fonds propres et la couverture de l'épargne restent non conformes."],
        valable="ratios de 2020 au 31/08/2026 (balance provisoire), statut global jugé de 2022 à 2025",
        source=SRC_REG)

