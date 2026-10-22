# ---- v9 : règles et taux de l'épargne, intérêts servis, épargne présentée en Assemblée générale (partie 6)
SRC_EP = "Source : CAHIER DE CHARGEEFC & PEP 2021.pdf ; MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx ; CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx (reçus le 22/10/2026) ; entretien du chef comptable du 19/10/2026"
contenu({"titre": "EFC et PEP : trois taux dans trois documents", "sous_titre": "Le taux servi sur l'EFC vaut 3,5 %, 5,5 % ou 14 % selon le document",
         "synthese": "Aucun document ne fixe un taux unique, signé et daté : l'engagement d'épargne ne peut pas être évalué sans ce taux"},
        "Le taux servi est l'intérêt que le FPPN doit ajouter à l'épargne de chaque mutualiste. Il fixe la valeur de l'engagement : à épargne égale, un taux plus élevé fait payer davantage au départ du mutualiste.",
        tableau={"titre": tab("Taux servis sur l'épargne selon les documents du FPPN"), "entetes": ["Document (fichier exact)", "Produit", "Taux écrit", "Portée"], "largeurs": [10.4, 3.0, 7.2, 10.6], "taille": 8.5,
                 "lignes": [["CAHIER DE CHARGEEFC & PEP 2021.pdf, « B- Taux d'intérêt applicable »", "EFC", "14 % d'intérêt simple, sans période indiquée", "Cotisations au-delà de 2024 et souscriptions postérieures au cahier des charges ; bonus de 300 000 FCFA si à jour"],
                            ["MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx", "EFC", "5,5 % l'an", "Rémunération de l'EFC ; certificat d'adoption non signé"],
                            ["MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx", "PEP", "3,5 % tous les ans", "Rémunération du PEP"],
                            ["CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx", "EFC", "5,5 % à partir de 2024", "Taux d'intérêt de l'EFC"],
                            ["Entretien du chef comptable, 19/10/2026", "Épargne", "3,5 %", "Taux servi déclaré, sans document"]]},
        disposition="dessous",
        analyse=["**Pour l'EFC, trois taux coexistent : 3,5 %, 5,5 % l'an et 14 % d'intérêt simple.** Le cahier des charges n'indique pas si les 14 % s'appliquent une fois ou chaque année.",
                 "**Aucun des trois documents n'est accompagné d'une décision d'adoption signée.** Le manuel porte un certificat d'adoption non signé."],
        implication="Le taux applicable doit être arrêté par une décision des organes du FPPN, puis appliqué de manière identique aux comptes, aux relevés individuels et au calcul de l'engagement (PBC, item 99).",
        valable="documents du FPPN reçus le 22/10/2026 et entretien du 19/10/2026 ; le taux réellement appliqué aux mutualistes n'est pas connu",
        source=SRC_EP)
_gl = V9["gl2024_epargne"]
_nat = {(x["compte"], x["nature"]): x for x in _gl}
_lig_gl = [[("PEP (41195100)" if x["compte"] == "41195100" else "EFC (41195200)"), x["nature"], N(f"GL écr déb {x['compte']} {x['nature']}", x["ecritures_debit"]),
            F(f"GL déb {x['compte']} {x['nature']}", x["debit"]), N(f"GL écr cré {x['compte']} {x['nature']}", x["ecritures_credit"]), F(f"GL cré {x['compte']} {x['nature']}", x["credit"])] for x in _gl]
contenu({"titre": "Intérêts servis : aucune écriture dans les comptes", "sous_titre": "Les comptes d'épargne ne reçoivent que des cotisations et ne versent que des remboursements",
         "synthese": "En 2024, aucune écriture d'intérêts n'est passée sur les comptes PEP et EFC, et aucune balance de 2020 à 2026 ne porte de charge d'intérêts servis aux épargnants"},
        "Le grand livre 2024 détaille chaque écriture des deux comptes d'épargne. Les écritures sont classées par nature d'après leur libellé : précomptes et cotisations, règlements EFC, remboursements PEP, rachats, autres. Une écriture d'intérêts servis porterait un libellé « intérêts » ou une contrepartie en charge (classe 67).",
        tableau={"titre": tab("Écritures 2024 des comptes d'épargne par nature"), "entetes": ["Compte", "Nature", "Écritures au débit", "Débit", "Écritures au crédit", "Crédit"], "largeurs": [4.4, 6.0, 3.6, 6.0, 3.6, 6.0], "taille": 9,
                 "lignes": _lig_gl},
        disposition="dessous",
        analyse=["**Aucune écriture d'intérêts ne figure sur les comptes d'épargne en 2024, et aucun compte de charge d'intérêts servis n'existe dans les balances de 2020 au 31/08/2026.** En 2024, le solde comptable de l'épargne n'évolue donc que par les cotisations, les remboursements et les rachats.",
                 f"**En 2024, {N('Règlements EFC 2024', _nat[('41195200', 'Règlements EFC')]['ecritures_debit'])} règlements EFC ont été payés pour {F('Montant règlements EFC 2024', _nat[('41195200', 'Règlements EFC')]['debit'])}.** Le grand livre ne dit pas si ces règlements comprennent des intérêts calculés hors comptabilité."],
        implication="Si un taux est dû aux mutualistes, la dette d'épargne inscrite au bilan est inférieure à l'engagement réel du montant des intérêts acquis et non comptabilisés. Ce montant ne peut être chiffré qu'avec le taux applicable et l'historique individuel des versements.",
        valable="grand livre 2024 pour le détail des écritures ; balances de 2020 au 31/08/2026 pour l'absence de charge d'intérêts",
        source="Source : GRAND LIVRE FPPN 2024.xlsx (comptes 41195100 et 41195200) ; " + SRC_COMPTA[9:])
_ef = V9["efc_taux"]
contenu({"titre": "Sensibilité : intérêts annuels dus sur l'EFC", "sous_titre": "Au taux de 5,5 %, les intérêts dus chaque année dépassent les produits financiers du FPPN",
         "synthese": f"Encours EFC au 31/12/2025 : {F('EFC encours 2025 (t)', next(x['encours_efc'] for x in _ef if x['exercice'] == 2025))} ; intérêts d'une année à 5,5 % : {F('EFC int 5,5 2025', next(x['interets_5_5'] for x in _ef if x['exercice'] == 2025))}"},
        "Le calcul applique chaque taux écrit à l'encours EFC de fin d'année (compte 41195200), une hypothèse à la fois. La colonne à 14 % suppose un taux annuel : c'est une borne haute, car le cahier des charges ne précise pas la période. Les produits financiers sont ceux de la partie 9.",
        tableau={"titre": tab("Intérêts d'une année sur l'encours EFC selon le taux servi"), "entetes": ["Exercice", "Encours EFC (41195200)", "Intérêts à 3,5 %", "Intérêts à 5,5 %", "Intérêts à 14 % (borne haute)", "Produits financiers"],
                 "largeurs": [3.0, 6.0, 5.2, 5.2, 6.0, 5.6], "taille": 9,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "31/08/2026", F(f"EFC enc {x['exercice']}", x["encours_efc"]), F(f"EFC i3,5 {x['exercice']}", x["interets_3_5"]),
                             F(f"EFC i5,5 {x['exercice']}", x["interets_5_5"]), F(f"EFC i14 {x['exercice']}", x["interets_14"]),
                             F(f"Prod fin (EFC) {x['exercice']}", x["produits_financiers"]) if x["produits_financiers"] is not None else "-"] for x in _ef]},
        disposition="dessous",
        analyse=[f"**Même au taux le plus bas (3,5 %), les intérêts d'une année sur l'EFC ({F('EFC i3,5 2025 (t)', next(x['interets_3_5'] for x in _ef if x['exercice'] == 2025))} en 2025) dépassent, chaque année de 2021 à 2025, les produits financiers de l'année.**",
                 "**L'épargne n'est couverte qu'en faible part par des actifs qui lui sont réservés (partie 6, adossement de l'épargne).** Les produits de ces actifs ne suffisent pas à servir un taux, quel qu'il soit, sans recourir aux cotisations des garanties."],
        implication="Le taux promis aux mutualistes doit être cohérent avec le rendement des actifs qui couvrent l'épargne. À défaut, chaque point de taux servi crée une perte récurrente que les excédents des garanties maladie et décès devront combler.",
        valable="encours EFC comptable au 31/12 de 2020 à 2025 et au 31/08/2026, une année d'intérêts, sans projection ni capitalisation",
        source=SRC_EP + " ; " + SRC_COMPTA[9:])
_e25 = next(x for x in V9["epargne_mouvements"] if x["exercice"] == 2025)
_e24 = next(x for x in V9["epargne_mouvements"] if x["exercice"] == 2024)
contenu({"titre": "Rapport de gestion 2025 : l'épargne présentée est la collecte de l'année", "sous_titre": "Le rapport annonce 4,39 milliards de FCFA d'épargne, alors que la dette d'épargne atteint 13,81 milliards",
         "synthese": "Les montants du rapport de gestion 2025 sont les crédits de l'année des comptes 41195100 et 41195200, présentés comme « le montant de ces épargnes au 31/12/2025 »"},
        "Le rapport de gestion détaillé de l'exercice 2025 (page 98, tableau 2 (III)) présente l'épargne PEP et EFC. Ses montants sont comparés aux mouvements et aux soldes des mêmes comptes dans la balance 2025.",
        tableau={"titre": tab("Épargne au 31/12/2025 : rapport de gestion et balance"), "entetes": ["Compte", "Rapport de gestion 2025", "Crédits de l'année (balance)", "Solde au 31/12/2025 (balance)", "Solde au 31/12/2024 (balance)"],
                 "largeurs": [4.6, 6.4, 6.4, 6.8, 6.8], "taille": 9,
                 "lignes": [["PEP (41195100)", F("RG PEP 2025", 2038935495), F("Crédits PEP 2025", _e25["mvt_c_41195100"]), F("Solde PEP 2025 (RG)", _e25["sol_c_41195100"]), F("Solde PEP 2024 (RG)", _e24["sol_c_41195100"])],
                            ["EFC (41195200)", F("RG EFC 2025", 2348888140), F("Crédits EFC 2025", _e25["mvt_c_41195200"]), F("Solde EFC 2025 (RG)", _e25["sol_c_41195200"]), F("Solde EFC 2024 (RG)", _e24["sol_c_41195200"])],
                            ["Total", F("RG total 2025", 4387823635), F("Crédits total 2025", _e25["mvt_c_41195100"] + _e25["mvt_c_41195200"]), F("Solde total 2025 (RG)", _e25["sol_c_41195100"] + _e25["sol_c_41195200"]), F("Solde total 2024 (RG)", _e24["sol_c_41195100"] + _e24["sol_c_41195200"])]]},
        disposition="dessous",
        analyse=["**Les montants du rapport de gestion sont exactement les crédits de l'année dans la balance**, c'est-à-dire les cotisations d'épargne collectées en 2025, et non l'épargne due aux mutualistes.",
                 "**La dette d'épargne réelle au 31/12/2025 est plus de trois fois supérieure au montant présenté.** Le rapport compare aussi la collecte 2025 à la collecte 2024 en la qualifiant de « variation » de l'épargne."],
        implication="L'information donnée à l'Assemblée générale minore l'engagement d'épargne du FPPN. Le rapport de gestion doit présenter le solde dû aux mutualistes, distinct de la collecte de l'année (PBC, item 112).",
        valable="exercice 2025 ; les rapports de gestion des autres années n'ont pas été contrôlés sur ce point",
        source="Source : Rapport de Gestion  2025 Final (2).pdf, pages 97 et 98, tableau 2 (III) ; " + SRC_COMPTA[9:])
