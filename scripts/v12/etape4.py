# -*- coding: utf-8 -*-
# v12, étape 4 (02/11/2026) : âge de départ du personnel porté à 60 ans (ordonnance n° 2012-03 du 11/01/2012, CNPS),
# la convention collective du 19/07/1977 (art. 40, 55 ans) devient une sensibilité. Exécuté par patch_v12.py.
ORD = "ordonnance n° 2012-03 du 11/01/2012 (CNPS)"
rep("plafond de 25 fois le SMIG annuel et départ à 55 ans, plus un an par enfant à charge dans la limite de trois (convention collective interprofessionnelle du 19/07/1977, art. 40).",
    "plafond de 25 fois le SMIG annuel (convention collective interprofessionnelle du 19/07/1977, art. 40). Départ à 60 ans dans le secteur privé depuis l'" + ORD + " ; la convention de 1977 fixait 55 ans, plus un an par enfant à charge dans la limite de trois.")
rep('''["Âge de départ", N("PIDR âge", PP["age_depart"]) + " ans", "Convention collective interprofessionnelle du 19/07/1977, art. 40 : 55 ans, plus un an par enfant à charge (au plus 3) ; les enfants ne figurent pas dans les fichiers reçus."],''',
    '''["Âge de départ", N("PIDR âge", PP["age_depart"]) + " ans", "Ordonnance n° 2012-03 du 11/01/2012 (CNPS) : 60 ans dans le secteur privé. La convention collective du 19/07/1977 (art. 40 : 55 ans, plus un an par enfant à charge) est dépassée sur ce point et testée en sensibilité. [[Texte officiel et âge appliqué par le FPPN à confirmer.]]"],''')
rep("Seuls l'âge de départ, faute d'enfants déclarés, et le taux d'actualisation, faute de marché profond, relèvent d'un choix justifié.",
    "Seuls l'âge de départ, que le FPPN n'a pas confirmé, et le taux d'actualisation, faute de marché profond, relèvent d'un choix justifié.")
rep('''_ag55 = [x for x in _P25 if x["age"] >= PP["age_depart"]]''', '''_ag55 = [x for x in _P25 if x["age"] >= 55]''')
rep('''analyse=[f"**{N('PIDR âge 55 atteint', len(_ag55))} agents ont déjà {N('PIDR âge (t)', PP['age_depart'])} ans ou plus et sont toujours en poste en 2025, dont un de {N('PIDR âge max', max(x['age'] for x in _ag55))} ans.** [[Le FPPN n'applique donc pas l'âge de 55 ans]] : la sensibilité ci-dessus chiffre un départ à 58 et à 60 ans.",''',
    '''analyse=[f"**{N('PIDR âge 55 atteint', len(_ag55))} agents ont déjà 55 ans ou plus et sont toujours en poste en 2025, dont un de {N('PIDR âge max', max(x['age'] for x in _ag55))} ans.** Le FPPN n'applique donc pas l'âge de 55 ans de la convention de 1977, ce qui concorde avec l'âge de {N('PIDR âge (t)', PP['age_depart'])} ans de l'ordonnance de 2012. La sensibilité ci-dessus chiffre un départ à 55 et à 58 ans.",''')
rep("valable=\"personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel chiffré à part\",\n        source=SRC_DISA + \" ; IAS 19, § 145 ; convention collective interprofessionnelle du 19/07/1977, art. 40 ;",
    "valable=\"personnel civil de la DISA 2025, évaluation au 31/12/2025 ; plafond conventionnel chiffré à part\",\n        source=SRC_DISA + \" ; IAS 19, § 145 ; ordonnance n° 2012-03 ; convention collective du 19/07/1977, art. 40 ;")
rep('source=SRC_DISA + " ; décret n° 2017-210 du 30/03/2017 ; convention collective du 19/07/1977, art. 40 ; IAS 19, § 67 et 83 ; Investing.com ; table CIMA F")',
    'source=SRC_DISA + " ; décret n° 2017-210 du 30/03/2017 ; ' + ORD + ' ; convention collective du 19/07/1977, art. 40 ; IAS 19, § 67 et 83 ; Investing.com ; table CIMA F")')
