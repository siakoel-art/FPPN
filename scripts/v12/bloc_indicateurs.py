# ---- v12 : indicateurs clés de 2020 au 31/08/2026, valeur et boule dans chaque case
_RGm = {(x["exercice"], x["garantie"]): x for x in V11["ratios_garantie"]}
_AN = list(range(2020, 2027))
_ENT_AN = [str(a) if a < 2026 else "31/08/2026" for a in _AN]
_GEa = {int(g["arrete"][-4:]): g["encours"] for g in GE if g["arrete"].startswith("31/12") or g["arrete"] == "31/08/2026"}


def _rt(c, a, en_pc=False):
    """Valeur d'un ratio de la Décision et boule de son statut de l'année."""
    x = RT[c]
    v = x["valeurs"].get(a)
    if v is None:
        return "n.d."
    t_ = pc(f"{c} {a} (%)", v) if en_pc else val_ratio(x, a)
    return bv(t_, boule_statut(x["statuts"].get(a, "")))


def _gar(g, a, cle):
    """Ratio d'une garantie : sinistralité (repère 70 % à 80 %), ratio combiné (seuil 100 %), résultat technique (seuil 0)."""
    x = _RGm.get((a, g))
    if not x:
        return "n.d."
    v = x[cle]
    if cle == "sinistralite":
        return bv(pc(f"RG sin {a} {g}", v), boule_fourchette(v, 0.70, 0.80))
    if cle == "ratio_combine":
        return bv(pc(f"RG rc {a} {g}", v), boule_seuil(v, 1.0, "max"))
    return bv(F(f"RG rt {a} {g}", v), boule_seuil(v, 0.0, "min"))


LEG_MAT = [("🔴", "Non respecté", "norme ou seuil dépassé"), ("🟢", "Respecté", "norme ou seuil respecté"),
           ("🟠", "Sous la fourchette", "sinistralité d'une garantie sous 70 %"), ("⚪", "À apprécier", "norme non chiffrée")]
_IND = [
    ["Cotisations acquises (comptes 70)", "Sans norme"] + [F(f"IC cot {a}", RT["SIN"]["den"][a]) for a in _AN],
    ["Prestations, provisions comprises (comptes 60, 6901, 7901)", "Sans norme"] + [F(f"IC pre {a}", RT["SIN"]["num"][a]) for a in _AN],
    ["Taux de sinistralité réglementaire, toutes garanties", "70 % à 80 %"] + [_rt("SIN", a) for a in _AN],
    ["Sinistralité maladie", "Repère 70 % à 80 %"] + [_gar("Maladie", a, "sinistralite") for a in _AN],
    ["Sinistralité décès", "Repère 70 % à 80 %"] + [_gar("Décès", a, "sinistralite") for a in _AN],
    ["Ratio combiné maladie (prestations et frais)", "Au plus 100 %"] + [_gar("Maladie", a, "ratio_combine") for a in _AN],
    ["Taux des charges de fonctionnement", "Au plus 20 %"] + [_rt("CHF", a) for a in _AN],
    ["Taux d'excédent net", "Au plus 5 %"] + [_rt("EXN", a) for a in _AN],
    ["Épargne due (PEP et EFC)", "Sans norme"] + [F(f"IC ep {a}", b[a]["epargne_totale"]) for a in _AN],
    ["Épargne couverte par des actifs réservés", "Au moins 100 %"] + [_rt("RL1", a, True) for a in _AN],
    ["Fonds propres, résultat compris", "Sans norme"] + [F(f"IC fp {a}", b[a]["fonds_propres"]) for a in _AN],
    ["Fonds propres et assimilés / dettes", "Au moins 1"] + [_rt("FEP", a) for a in _AN],
    ["Ratio de liquidité", "Au moins 1"] + [_rt("LIQ", a) for a in _AN],
    ["Encours des prêts du Grenier", "Sans norme"] + [F(f"IC gre {a}", _GEa[a]) if a in _GEa else "n.d." for a in _AN],
]
contenu({"titre": "Indicateurs clés de 2020 au 31/08/2026", "sous_titre": "La sinistralité monte, l'épargne reste peu couverte et les normes de solidité ne sont jamais atteintes",
         "synthese": f"Sinistralité réglementaire : {pc('IC sin 2021', RT['SIN']['valeurs'][2021])} en 2021, {pc('IC sin 2025', RT['SIN']['valeurs'][2025])} en 2025, {pc('IC sin 2026', RT['SIN']['valeurs'][2026])} sur huit mois de 2026"},
        "Chaque case donne la valeur de l'année et une boule qui la compare à la norme de la Décision n° 05/2022 ou au repère indiqué. 2020 et 2021 sont antérieurs à l'entrée en vigueur de la Décision (01/06/2022) : ils sont indicatifs. 2026 couvre huit mois, sur la balance provisoire.",
        tableau={"titre": tab("Indicateurs clés, de 2020 au 31/08/2026"), "entetes": ["Indicateur", "Norme ou repère"] + _ENT_AN,
                 "largeurs": [6.4, 3.0] + [3.114] * 7, "taille": 7.5, "sans_code": True, "legende": LEG_MAT[:3], "lignes": _IND},
        disposition="dessous",
        analyse=[f"**La sinistralité réglementaire passe de {pc('IC sin 2021 (a)', RT['SIN']['valeurs'][2021])} en 2021 à {pc('IC sin 2025 (a)', RT['SIN']['valeurs'][2025])} en 2025, et la maladie dépasse 80 % sur les huit mois de 2026 ({pc('IC sin mal 2026', _RGm[(2026, 'Maladie')]['sinistralite'])}).** "
                 "Les ratios de solidité (couverture de l'épargne, fonds propres et assimilés, liquidité) sont non conformes chaque année de 2020 au 31/08/2026."],
        implication="Les écarts se répètent chaque année : ils sont structurels. Le plan de mise en conformité doit traiter leurs causes (provisions, épargne, liquidité), pas une année isolée.",
        valable="comptes de 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire) ; 2020 et 2021 indicatifs ; encours du Grenier aux 31/12 et au 31/08/2026",
        source=SRC_COMPTA + " ; " + DEC + " ; " + SRC_GRENIER[9:])
_DEF_IND = [
    ["Cotisations acquises", "Les cotisations de l'année revenant au FPPN (comptes 70, variation des provisions sur cotisations comprise).", "Elles mesurent la ressource technique de l'année."],
    ["Prestations, provisions comprises", "Les prestations de l'année (comptes 60) et la variation des provisions techniques (6901, 7901).", "Elles mesurent le coût des garanties de l'année."],
    ["Taux de sinistralité réglementaire", DEFS["SIN"][0], DEFS["SIN"][1] + " Seul ce taux global a une norme dans la Décision."],
    ["Sinistralité maladie ou décès", "Les prestations d'une garantie divisées par ses cotisations.", "Elle montre quelle garantie consomme ses cotisations. La fourchette de 70 % à 80 % sert de repère."],
    ["Ratio combiné maladie", "Les prestations maladie et les charges de fonctionnement affectées, divisées par les cotisations maladie.", "Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte, frais compris."],
    ["Taux des charges de fonctionnement", DEFS["CHF"][0], DEFS["CHF"][1]],
    ["Taux d'excédent net", DEFS["EXN"][0], DEFS["EXN"][1]],
    ["Épargne couverte par des actifs réservés", DEFS["RL1"][0], DEFS["RL1"][1]],
    ["Fonds propres, résultat compris", "Les ressources propres du FPPN (comptes 10 à 14) : réserves, report à nouveau et résultat.", "Négatifs : les pertes passées dépassent les réserves, le FPPN n'a pas de coussin propre."],
    ["Fonds propres et assimilés / dettes", DEFS["FEP"][0] + " Les provisions techniques comptent dans les fonds assimilés.", DEFS["FEP"][1]],
    ["Ratio de liquidité", DEFS["LIQ"][0], DEFS["LIQ"][1]],
    ["Encours des prêts du Grenier", "Le capital restant dû par les policiers emprunteurs.", "Il mesure le montant de l'épargne et des fonds du FPPN engagé dans le crédit."],
]
contenu({"titre": "Définition et lecture des indicateurs clés", "sous_titre": "Ce que chaque indicateur mesure et comment lire sa valeur",
         "synthese": "Les normes viennent de la Décision n° 05/2022 ; les repères des garanties sont propres à l'analyse d'Adenka"},
        "Définitions des indicateurs de la diapositive précédente et des ratios du résumé exécutif. Formules exactes des ratios réglementaires : partie 10.",
        tableau={"titre": tab("Définition et lecture des indicateurs clés"), "entetes": ["Indicateur", "Ce qu'il mesure", "Comment le lire"],
                 "largeurs": [6.4, 12.0, 12.8], "taille": 8, "sans_code": True, "lignes": _DEF_IND},
        valable="formules et normes en vigueur depuis le 01/06/2022",
        source="Source : " + DEC + " ; lecture Adenka")

