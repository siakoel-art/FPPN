# ---- v13 : circuits des cotisations par produit, d'après les logigrammes du FPPN, complétés des contrôles absents (cadre rouge)
_SRC_FLOW = "Source : « FLOW_CHART_COTISATION_ET_PRODUITS_DEPARGNE_1.pptx » (FPPN, reçu le 23/10/2026) ; contrôles absents : analyses Adenka"
contenu({"titre": "Circuit EFC et PEP : du bulletin à la comptabilisation", "sous_titre": "Le FPPN produit le fichier de l'épargne, mais ne rapproche pas l'aller et le retour",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Forme arrondie bleue : début ou fin. Rectangle : étape. Losange : décision. Forme ondulée : document. Cadre rouge : contrôle absent.",
        logigramme="epargne",
        analyse=["**Le circuit s'arrête à la comptabilisation des précomptes.** Aucune étape ne rapproche le fichier aller du fichier retour, aucune ne calcule les intérêts dus ni ne produit de relevé par épargnant (" + F("CE int", INT55) + " d'intérêts dus au 31/08/2026)."],
        source=_SRC_FLOW + " ; " + SRC_ENT19[9:])
contenu({"titre": "Circuit de la flotte automobile (GNA)", "sous_titre": "Le FPPN encaisse une commission de 20 % dont l'assiette n'est pas documentée",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Cadre rouge : contrôle absent ou fichier manquant.",
        logigramme="gna",
        analyse=["**Le FPPN collecte les primes pour le GNA et gère les trop-perçus, mais la convention d'apporteur d'affaires n'a pas été transmise.** L'assiette de la commission de 20 % ne peut pas être vérifiée, et le fichier retour de janvier 2026 manque."],
        source=_SRC_FLOW)
contenu({"titre": "Circuit des retraités (CGRAE)", "sous_titre": "La cotisation suit le barème de la CGRAE, dans la limite de la quotité cessible",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Barème de la CGRAE : 2,5 % de (pension de base - MUGEFCI) + 5 000 à 15 000 FCFA selon le grade.",
        logigramme="cgrae",
        analyse=["**Seuls les retraités qui disposent d'une quotité cessible sont précomptés : la liste des autres souscripteurs n'a pas été transmise.** Les fichiers retour de janvier et de juin 2024 manquent."],
        source=_SRC_FLOW + " ; « BAREME_CGRAE.xlsx » ; CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf")
contenu({"titre": "Circuit de la maladie et du décès", "sous_titre": "Le fichier aller est produit hors du FPPN : le recouvrement ne peut pas être mesuré",
         "synthese": "Logigramme du FPPN redessiné ; les cadres rouges signalent les contrôles absents constatés par Adenka"},
        "Couloirs repris du logigramme du FPPN. Sans carte de santé FPPN, l'adhérent n'a pas de prestation.",
        logigramme="maladie_deces",
        analyse=["**Le service du personnel calcule et transmet le fichier aller sans le communiquer au FPPN.** Le FPPN ne peut ni vérifier les taux retenus (3,77 % et 1,23 % au lieu de 3,5 % et 1,5 %), ni mesurer le taux de recouvrement exigé par la Décision n° 05/2022 (au moins 90 %)."],
        source=_SRC_FLOW + " ; statuts du FPPN édités le 15/06/2023, art. 17 ; fichiers de précomptes 936 et 8Z6")
