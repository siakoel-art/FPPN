# -*- coding: utf-8 -*-
# v13 (13/11/2026) : corrections du texte issues du registre de vérification de la revue indépendante.
# Exécuté par patch_v13.py (fonctions rep et chaîne s du générateur). Chaque texte nouveau est enregistré pour le bleu.
CORR13 = [
    # adossement (A-B1, D-B1)
    ('"Non-adossement démontré (borne favorable)"', '"Non-adossement démontré"', None),
    ("implication=\"Cet adossement est une borne favorable (diapositive suivante) : le passif à moins d'un an comprend les provisions pour prestations à payer, dont le montant n'est pas recalculé.",
     "implication=\"Les provisions figées gonflent le passif à moins d'un an : la lecture des comptes tels quels est le cas le plus défavorable, et la lecture avec les provisions ramenées au besoin est donnée à côté.",
     "Les provisions figées gonflent le passif à moins d'un an : la lecture des comptes tels quels est le cas le plus défavorable, et la lecture avec les provisions ramenées au besoin est donnée à côté."),
    ("Adossement par échéance établi sur les grands livres : borne favorable", "Adossement par échéance établi sur les grands livres, en deux lectures",
     "Adossement par échéance établi sur les grands livres, en deux lectures"),
    # tendances (B-M5, B-M6)
    ("f\"**La tendance est statistiquement significative.** Elle correspond à {N(",
     "f\"**La hausse arrive par paliers, chaque mois de juillet : aucun test de tendance n'a de valeur sur une série en marches.** En moyenne, cela fait {N(",
     "La hausse arrive par paliers, chaque mois de juillet : aucun test de tendance n'a de valeur sur une série en marches."),
    ("f\"La hausse du S/C maladie est statistiquement significative, d'environ ",
     "f\"La sinistralité maladie est stable de 2020 à 2022, puis monte de 23,31 points par an de 2023 à 2025 : sur six années avec une rupture, aucun test n'a de valeur probante. La pente moyenne est d'environ ",
     "La sinistralité maladie est stable de 2020 à 2022, puis monte de 23,31 points par an de 2023 à 2025"),
    # sinistralité ancienne (B-B3)
    ("de 2020 à 2022, le FPPN rendait moins de 10 % des cotisations en prestations",
     "en 2020 et 2022, le FPPN rendait moins de 10 FCFA de prestations pour 100 FCFA de cotisations (20,77 FCFA en 2021)",
     "en 2020 et 2022, le FPPN rendait moins de 10 FCFA de prestations pour 100 FCFA de cotisations (20,77 FCFA en 2021)"),
    # Grenier (C-M10, C-M3)
    ("retards maîtrisés mais précomptes incomplets", "retards en baisse à expliquer et précomptes incomplets", "retards en baisse à expliquer et précomptes incomplets"),
    ("**Le résultat du Grenier disparaîtrait avec la dépréciation prévue par sa propre politique de crédit** (",
     "**La dépréciation prévue par la politique de crédit réduit le résultat du Grenier sans le rendre négatif** (",
     "La dépréciation prévue par la politique de crédit réduit le résultat du Grenier sans le rendre négatif"),
    # provisions techniques (C-m8)
    ("prévoit quatre provisions techniques ; le FPPN n'en a constitué qu'une", "nomme trois provisions techniques ; le FPPN n'en tient qu'une partie",
     "nomme trois provisions techniques"),
    # sanctions (D-M7)
    ("car l'article 84 du règlement n° 07/2009 prévoit alors le retrait d'agrément.",
     "car le non-respect des ratios de contrôle (tableau 1 et section II de la Décision) entraîne des sanctions, et l'article 84 du règlement n° 07/2009 permet alors le retrait d'agrément.",
     "le non-respect des ratios de contrôle (tableau 1 et section II de la Décision) entraîne des sanctions"),
    # projection (E-B1)
    ("ce sont deux bornes construites sur les seules données agrégées disponibles.",
     "ce sont deux tendances bâties sur 2023 à 2025, déjà dépassées par les huit mois de 2026.",
     "ce sont deux tendances bâties sur 2023 à 2025, déjà dépassées par les huit mois de 2026."),
    ("elle remplacera ces deux bornes par un scénario central expliqué.", "elle remplacera ces deux tendances par un scénario central expliqué, recalé sur 2026.",
     "elle remplacera ces deux tendances par un scénario central expliqué, recalé sur 2026."),
    # documents (A-m8)
    ("lus en entier : quatre apportent des faits nouveaux", "lus en entier : cinq apportent des faits nouveaux", "cinq apportent des faits nouveaux"),
    # traçabilité (A, diapositives 52 et 299)
    ('"FPPN_Calculs_rapport_20261022.xlsx"', '"FPPN_Calculs_rapport_20261113.xlsx"', None),
    ("rapport_v9_fppn.py, donnees_v9_fppn.py,", "rapport_v13_fppn.py (étapes v11 à v13), donnees_v9_fppn.py à donnees_v11_fppn.py,", None),
    # validité (A-M12, E-M14)
    ("jusqu'au 22/10/2026", "jusqu'au 11/11/2026", "jusqu'au 11/11/2026"),
    # mots interdits et ponctuation (contrôle qualité du 13/11/2026)
    ("le FPPN assure la santé des policiers retraités", "le FPPN prend en charge la santé des policiers retraités", None),
    ("Le FPPN assure la santé", "Le FPPN prend en charge la santé", None),
    ("L'écart se creuse chaque année", "L'écart s'élargit chaque année", None),
    ("un rendement inférieur au taux servi creuse la dette d'épargne", "un rendement inférieur au taux servi alourdit la dette d'épargne", None),
    ('("Ce qui nous y a conduits"', '("Ce qui y a conduit"', None),
    ('"—", "Risque de crédi', '"n.d.", "Risque de crédi', None),
]
for _a, _b, _bleu in CORR13:
    rep(_a, _b, n=None)
_BLEU13 = [b for _, _, b in CORR13 if b]
rep("\"bleu_v13.py\"), encoding=\"utf-8\").read())", "\"bleu_v13.py\"), encoding=\"utf-8\").read())\nBLEU_P.extend(" + repr(_BLEU13) + ")")
