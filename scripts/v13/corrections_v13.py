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
     "f\"La sinistralité maladie, stable de 2020 à 2022, monte de 23,31 points par an de 2023 à 2025. Pente moyenne : ",
     "La sinistralité maladie, stable de 2020 à 2022, monte de 23,31 points par an de 2023 à 2025"),
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
    # extourne appliquée aux données de sinistralité figées de 2026 (sinistralite_v15.json)
    ('"assets", "sinistralite_v15.json"), encoding="utf-8")).items()}',
     '"assets", "sinistralite_v15.json"), encoding="utf-8")).items()}\n'
     '_E13 = K13["extourne"]\n'
     '_q = SQ[2026]\n'
     '_q["prest"] -= _E13\n_q["cmp"] -= _E13\n_q["soins"] -= _E13\n_q["sin_num"] -= _E13\n'
     '_q["cats"]["Clinique du FPPN (CMP)"] -= _E13\n'
     '_q.update(sc=_q["prest"] / _q["cot"], sc_soins=_q["soins"] / _q["cot"], part_cmp=_q["cmp"] / _q["soins"], '
     'comb=(_q["prest"] + _q["frais"]) / _q["cot"], comb_soins=(_q["soins"] + _q["frais"]) / _q["cot"], '
     'sin=_q["sin_num"] / _q["sin_den"], sin_soins=(_q["sin_num"] - _q["non_soins"]) / _q["sin_den"])', None),
    ('"31/08/2026 : 75,20 %, dans la norme pour la première fois, budget 2026 calé sur 80 %"',
     '"31/08/2026 : " + n13(K13["regl_26"], 2) + " % après extourne de la provision de 2025, sous la norme, budget 2026 calé sur 80 %"',
     "après extourne de la provision de 2025, sous la norme, budget 2026 calé sur 80 %"),
    # tournures relevées par le contrôle qualité
    ("porte la provision la plus élevée de ce tableau", "porte la plus forte provision de ce tableau", None),
    ("Sous 1 : un retrait massif ne pourrait pas être payé avec l'argent disponible.", "Sous 1 : un retrait massif n'est pas payable avec l'argent disponible.", None),
    ("Plus le solde est élevé, plus le manque à gagner est important.", "Plus le solde est fort, plus le manque à gagner est important.", None),
    ("les frais de gestion sont élevés au regard des ressources.", "les frais de gestion pèsent trop au regard des ressources.", None),
    # intérêts : le montant du manuel est un scénario, la règle retenue est celle des entretiens (13/11/2026)
    ('"Intérêts dus non comptabilisés : " + F("PPP int", INT55) + " au 31/08/2026 (PEP 3,5',
     '"Intérêts selon le manuel, non comptabilisés : " + F("PPP int", INT55) + " au 31/08/2026 (PEP 3,5', "Intérêts selon le manuel, non comptabilisés"),
    ("f\"Intérêts dus non provisionnés : {F('EN int', INT55)}.", "f\"Intérêts selon le manuel, non provisionnés : {F('EN int', INT55)}.", "Intérêts selon le manuel, non provisionnés"),
    ("aucune ne calcule les intérêts dus ni ne produit de relevé par épargnant (\" + F(\"CE int\", IN",
     "aucune ne calcule les intérêts ni ne produit de relevé par épargnant (selon le manuel, \" + F(\"CE int\", IN", "selon le manuel,"),
    ('"Solde comptable + intérêts dus du 01/01/2020 au taux du manuel"', '"Solde comptable + intérêts du 01/01/2020 au taux du manuel (scénario le plus haut)"', "(scénario le plus haut)"),
    # provisions maladie et décès (E-M7, C-M2)
    ("f\"Montant exact, aujourd'hui borné à {F('AT psap'", "f\"Montant exact, aujourd'hui estimé à au moins {F('AT psap'", "aujourd'hui estimé à au moins"),
    ("(états 2021 et 2023)", "(état 2021)", "(état 2021)"),
    # actifs de valeur non établie : SONEB ajouté (D-M11)
    ("Ce calcul reste une sensibilité tant que le CA n'a pas décidé.",
     "En retirant aussi le terrain SONEB (450 000 000 FCFA, sans titre), ils tombent à {n13(K13['fp_sans_contestes'])} FCFA au 31/08/2026. Ce calcul reste une sensibilité tant que le conseil d'administration n'a pas décidé.",
     "En retirant aussi le terrain SONEB (450 000 000 FCFA, sans titre)"),
    # retraits (C-M5)
    ("un retrait de 21,34 % épuiserait toute la trésorerie", "un retrait de 21,34 % épuise toute la trésorerie, et de \" + n13(K13[\"seuil_retrait\"], 2) + \" % une fois les dettes de l'année payées",
     "% une fois les dettes de l'année payées"),
    # sinistralité 2026 écrite en dur : valeurs relues dans SQ après extourne
    ('"Maladie seule : 73,53 % (2025) puis 97,40 % (31/08/2026)"',
     '"Maladie seule : 73,53 % (2025) puis " + n13(100 * SQ[2026]["sc"], 2) + " % (31/08/2026, après extourne)"', "% (31/08/2026, après extourne)"),
    ("141,52 % sur les huit mois de 2026", '" + n13(100 * SQ[2026]["comb"], 2) + " % sur les huit mois de 2026 après extourne', "% sur les huit mois de 2026 après extourne"),
    ('"Les cotisations reviennent davantage aux adhérents : 75,20 % au 31/08/2026, dans la norme"',
     '"Les cotisations reviennent davantage aux adhérents : " + n13(100 * SQ[2026]["sin"], 2) + " % au 31/08/2026 après extourne, sous la norme"', "% au 31/08/2026 après extourne, sous la norme"),
    ("(ratio combiné de 119,07 % en 2026)", '(ratio combiné de " + n13(100 * SQ[2026]["comb_soins"], 2) + " % en 2026 après extourne)', "% en 2026 après extourne)"),
    ("de 74,95 % au 31/08/2026.", "de \" + n13(100 * SQ[2026]['sc_soins'], 2) + \" % au 31/08/2026 après extourne.", "% au 31/08/2026 après extourne."),
    # délai de paiement 2020 : la balance 2020 couvre quatre mois (C-m1)
    ('DELAI = {a: _PM[a]["dettes_cliniques"] / (_PM[a]["factures_enregistrees"] / _PM[a]["mois"]) for a in range(2020, 2027)}',
     'DELAI = {a: _PM[a]["dettes_cliniques"] / (_PM[a]["factures_enregistrees"] / _PM[a]["mois"]) for a in range(2020, 2027)}\nDELAI[2020] = _PM[2020]["dettes_cliniques"] / (_PM[2020]["factures_enregistrees"] / 4)  # balance 2020 : du 01/09 au 31/12', None),
    # provision maladie : plus de rapport « N fois » (C-M1)
    (" et vaut de {N('PMN fois min', min(PV['psap_maladie'][str(c)] / psap_m_elargie(c) for c in (2022, 2023, 2024)), 0)} à {N('PMN fois max', max(PV['psap_maladie'][str(c)] / psap_m_elargie(c) for c in (2022, 2023, 2024)), 0)} fois le besoin mesuré.**",
     " et dépasse largement le besoin estimé par Adenka, qui reste un minimum.**", "et dépasse largement le besoin estimé par Adenka, qui reste un minimum."),
    # projection (E-B1)
    ("La réalité se situe entre les deux scénarios ;", "Les huit mois de 2026 sont déjà sous le scénario B ;", "Les huit mois de 2026 sont déjà sous le scénario B"),
    # listes des figures et des tableaux en 12 points, hauteur de ligne libre
    ('"hauteur_ligne": 0.3}, 1.8, 1.4 + col * 16.0, 7, hauteur_max=15.6)', '}, 1.8, 1.4 + col * 16.0, 12, hauteur_max=15.6)', None),
    ('"largeurs": [1.6, 12.4, 1.4]', '"largeurs": [2.3, 11.3, 1.8]', None),
]
for _a, _b, _bleu in CORR13:
    rep(_a, _b, n=None)
_BLEU13 = [b for _, _, b in CORR13 if b]
rep("\"bleu_v13.py\"), encoding=\"utf-8\").read())", "\"bleu_v13.py\"), encoding=\"utf-8\").read())\nBLEU_P.extend(" + repr(_BLEU13) + ")")
