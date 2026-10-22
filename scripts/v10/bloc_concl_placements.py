# ---- v10 : conclusion « adéquation des placements aux engagements » (règlement d'exécution n° 004/2022, art. 17)
_pf_ = {x["exercice"]: x for x in R["rendement_placements"]}
contenu({"titre": "Conclusion : adéquation des placements", "sous_titre": "Les placements ne sont adéquats ni en volume, ni en liquidité, ni en rendement",
         "synthese": "Aucune des quatre conditions de l'article 17 (sécurité, rendement, liquidité, diversification) n'est démontrée"},
        "Le règlement d'exécution n° 004/2022 (art. 17) exige que les placements satisfassent aux conditions de sécurité, de rendement, de liquidité et de diversification. Chaque condition est confrontée aux chiffres du FPPN de 2020 au 31/08/2026.",
        tableau={"titre": tab("Placements et emplois du FPPN confrontés aux conditions de l'article 17"), "entetes": ["Condition", "Ce qu'il faudrait", "Ce que montrent les chiffres", "Appréciation"], "largeurs": [3.6, 8.4, 14.0, 5.2], "taille": 8.5,
                 "lignes": [["Liquidité", "Couvrir les paiements de l'année, sorties d'épargne comprises", "Couverture de " + pc("CPL liq", couverture("31/08/2026", T_MOY)[2]) + " au 31/08/2026 (scénario central)", "Non remplie"],
                            ["Rendement", "Couvrir les intérêts dus sur l'épargne", "Produits financiers 2021 à 2025 : " + F("CPL pf", sum(_pf_[a]["produits_financiers"] for a in range(2021, 2026))) + " ; intérêts dus sur la même période : " + F("CPL int", interets_dus(0.035, 0.055, 2025) - interets_dus(0.035, 0.055, 2020)), "Non remplie"],
                            ["Sécurité", "Actifs évalués et contreparties solides", "Clinique, terrains et créance sur l'hôpital non évalués ; FCP non garanti", "Non démontrée"],
                            ["Diversification", "Pas de concentration sur un emploi", "Plus des trois quarts de l'actif à plus de cinq ans ou à délai incertain", "Non remplie"]]},
        disposition="dessous",
        analyse=["**Le portefeuille est inadéquat d'abord parce que l'épargne finance des actifs qui ne produisent ni revenu financier ni liquidité.** Le rendement des placements financiers eux-mêmes n'est pas le problème principal."],
        implication="Une politique de placement votée par le conseil d'administration doit fixer une réserve de liquidité, un rendement minimal au regard du taux servi arrêté et des limites sur les actifs immobiliers, avant tout nouvel investissement.",
        valable="placements et produits financiers de 2021 à 2025 ; actifs et épargne au 31/08/2026 (provisoire) ; intérêts dus aux taux du manuel d'épargne",
        source="Source : " + REX + ", art. 17 ; " + SRC_COMPTA[9:] + " ; diapositives " + ", ".join(ref(t) for t in ("Adossement : tests de résistance", "Épargne : intérêts dus non comptabilisés", "Rendement des placements")))
