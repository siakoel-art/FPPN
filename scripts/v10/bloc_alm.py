# ---- v10 : adossement actif-passif (méthodes justifiées, écoulement de l'épargne, écarts de liquidité, résistance, gouvernance)
SRC_ALM = {"ONT": "« ALM_ONT.pdf »", "BAN": "« Gestion des Risques Bancaires(Actif PAssif).pdf » (Sciences-Po, 2009-2010)", "IAA": "« Ch13_Gestion_Actif-Passif_French.pdf » (Association actuarielle internationale, 2016)",
           "TIC": "« Construction d'un modèle ALM.pdf »", "LAV": "« 2024081723_jerome-lavoie-pilote.pdf » (Institut des actuaires)"}
contenu({"titre": "Choix des méthodes d'adossement", "sous_titre": "Des méthodes de liquidité, adaptées à un passif court et à une épargne sans échéancier",
         "synthese": "Chaque méthode est retenue parce que ses conditions d'emploi sont remplies par les données du FPPN ; les méthodes écartées sont démontrées en annexe"},
        "La gestion actif-passif vérifie que les actifs produisent l'argent nécessaire au moment où les engagements doivent être payés. Le choix de la méthode dépend du profil du passif et des données disponibles.",
        tableau={"titre": tab("Méthodes d'adossement retenues et justification"), "entetes": ["Question", "Méthode retenue", "Pourquoi (document de référence, page)", "Donnée du FPPN qui le permet"], "largeurs": [5.0, 6.8, 11.0, 8.4], "taille": 8,
                 "lignes": [["Quel est l'enjeu pour la maladie et le décès ?", "Liquidité à court terme", "Contrats santé de durée courte : complexité actif-passif « faible », l'enjeu est la gestion du portefeuille d'actifs (" + SRC_ALM["ONT"] + ", p. 8)", f"Délai moyen de paiement des cliniques de {N('ALM délai', DELAI[2025], 2)} mois en 2025"],
                            ["Quand l'épargne sortira-t-elle ?", "Écoulement conventionnel", "Les comptes d'épargne sans échéance contractuelle ont un amortissement « à modéliser » (" + SRC_ALM["BAN"] + ", p. 33)", "Sorties annuelles des comptes 41195100 et 41195200 de 2021 à 2025"],
                            ["L'actif suit-il les sorties ?", "Écart de liquidité par échéance", "Écart entre passifs et actifs au cours du temps après écoulement des éléments du bilan (" + SRC_ALM["BAN"] + ", p. 86)", "Actifs classés par délai de mobilisation aux bilans 2025 et 2026"],
                            ["Quels actifs sont mobilisables ?", "Classes de liquidité", "Les titres sont classés selon leur degré de liquidité (" + SRC_ALM["BAN"] + ", p. 95)", "Nature comptable de chaque actif (banques, DAT, FCP, prêts, immeubles)"],
                            ["Que se passe-t-il en cas de choc ?", "Tests de résistance calibrés sur le FPPN", "Le rachat massif de 40 % de la formule standard européenne (" + SRC_ALM["TIC"] + ", p. 110) n'est pas calibré sur le FPPN", f"Taux de sortie maximal observé : {pc('ALM tmax', T_MAX)} ; hausse maximale des prestations maladie de 2024 et 2025 : {pc('ALM pic', PIC)}"],
                            ["Qui décide ?", "Gouvernance actif-passif", "Le conseil d'administration fixe les limites, un comité suit les écarts (" + SRC_ALM["IAA"] + ", p. 15 et 19 ; " + SRC_ALM["LAV"] + ", p. 8 et 11)", "Décision de placement par le seul directeur général (entretien du 19/10/2026)"]]},
        disposition="dessous",
        analyse=["**Les méthodes de taux (duration, immunisation, modèles stochastiques, VaR) ne sont pas retenues** : elles supposent des flux de passif connus et un portefeuille obligataire, absents au FPPN (annexe « Méthodes non retenues et pourquoi »)."],
        implication="Les résultats de l'adossement sont des ordres de grandeur fondés sur le comportement observé des mutualistes ; ils deviendront des flux contractuels dès que l'épargne aura un échéancier.",
        valable="choix de méthode fondé sur les documents de référence cités et sur les données du FPPN de 2020 au 31/08/2026",
        source="Source : " + " ; ".join(SRC_ALM.values()))
_EP10 = {e["exercice"]: e for e in V10["epargne"]}
contenu({"titre": "Adossement : à quel rythme l'épargne sort-elle ?", "sous_titre": f"De 2021 à 2025, il sort chaque année en moyenne {pc('ALM t moy', T_MOY)} de l'épargne du début d'année",
         "synthese": f"Taux de sortie compris entre {pc('ALM tmin', min(x['total_taux'] for x in SORT10))} et {pc('ALM tmax (s)', T_MAX)} ; le PEP sort plus vite que l'EFC"},
        "Taux de sortie de l'année = sorties de l'année (débits des comptes 41195100 et 41195200) / épargne au début de l'année. C'est l'écoulement conventionnel : faute d'échéancier, le comportement passé des mutualistes sert à prévoir les sorties futures.",
        tableau={"titre": tab("Sorties annuelles de l'épargne et taux de sortie"), "entetes": ["Exercice", "Épargne au début", "Sorties de l'année", "Taux PEP", "Taux EFC", "Taux total"], "largeurs": [3.0, 5.4, 5.4, 3.6, 3.6, 3.6], "taille": 9,
                 "lignes": [[str(x["exercice"]), F(f"ALM début {x['exercice']}", _EP10[x["exercice"]]["pep_debut"] + _EP10[x["exercice"]]["efc_debut"]),
                             F(f"ALM sorties {x['exercice']}", _EP10[x["exercice"]]["pep_sorties"] + _EP10[x["exercice"]]["efc_sorties"]),
                             pc(f"ALM tp {x['exercice']}", x["pep_taux"]), pc(f"ALM te {x['exercice']}", x["efc_taux"]), pc(f"ALM tt {x['exercice']}", x["total_taux"])] for x in SORT10]},
        disposition="droite",
        analyse=[f"**Le taux moyen ({pc('ALM t moy (a)', T_MOY)}) sert au scénario central, le taux maximal ({pc('ALM tmax (a)', T_MAX)}, en {max(SORT10, key=lambda x: x['total_taux'])['exercice']}) au test de résistance.**",
                 "**Les sorties de l'EFC suivent les départs à la retraite**, que seules les dates de naissance permettraient de prévoir."],
        implication="Les sorties d'épargne d'une année représentent plusieurs milliards de FCFA : c'est le premier besoin de liquidité du FPPN, devant les prestations.",
        valable="débits et soldes des comptes 41195100 et 41195200, balances 2020 à 2025 ; l'année 2020 sert de point de départ ; 2026 non retenue (huit mois)",
        source=SRC_COMPTA + " ; " + SRC_ALM["BAN"] + ", p. 33 à 36 ; onglet 24a du classeur des calculs")
# écarts de liquidité année par année au 31/08/2026 (écoulement statique, sans nouveaux versements)
_m = MOB10["31/08/2026"]
_E0 = _m["epargne_echeance_a_etablir"]
GAP = []
_cum = 0.0
for k in range(1, 6):
    act = (_m["immediat"] + _m["moins_1_an"]) if k == 1 else _m["1_a_5_ans"] / 4
    sor = _E0 * (1 - T_MOY) ** (k - 1) * T_MOY
    pas_ = (_m["passif_moins_1_an"] if k == 1 else 0.0) + sor
    _cum += act - pas_
    GAP.append((f"Année {k}", act, _m["passif_moins_1_an"] if k == 1 else 0.0, sor, act - pas_, _cum))
_res_ep = _E0 * (1 - T_MOY) ** 5
GAP.append(("Au-delà de 5 ans", _m["plus_5_ans_ou_incertain"], 0.0, _res_ep, _m["plus_5_ans_ou_incertain"] - _res_ep, _cum + _m["plus_5_ans_ou_incertain"] - _res_ep))
contenu({"titre": "Adossement : écarts de liquidité année par année", "sous_titre": f"Le manque de liquidité cumulé atteint {F('GAP an5', -GAP[4][5])} à la fin de la cinquième année",
         "synthese": f"Même en comptant à leur valeur comptable les terrains, la clinique et les créances longues, il manque {F('GAP fin', -GAP[5][5])}"},
        "Écoulement au 31/08/2026, sans nouveaux versements (écart statique). Sorties d'épargne de l'année k = épargne × (1 - taux moyen) puissance (k - 1) × taux moyen. Actifs à moins d'un an en année 1 ; actifs de 1 à 5 ans répartis à parts égales sur les années 2 à 5, faute d'échéancier. Écart = actifs disponibles - paiements.",
        tableau={"titre": tab("Écarts de liquidité au 31/08/2026, scénario central"), "entetes": ["Échéance", "Actifs disponibles", "Dettes à payer", "Sorties d'épargne", "Écart de l'année", "Écart cumulé"], "largeurs": [3.8, 5.4, 5.0, 5.0, 5.4, 5.4], "taille": 9,
                 "lignes": [[g[0], F(f"GAP act {g[0]}", g[1]), F(f"GAP det {g[0]}", g[2]), F(f"GAP sor {g[0]}", g[3]), F(f"GAP éc {g[0]}", g[4]), F(f"GAP cum {g[0]}", g[5])] for g in GAP]},
        disposition="dessous",
        analyse=[f"**Dès la première année, les paiements dépassent les actifs disponibles de {F('GAP an1', -GAP[0][4])}.** L'écart se creuse chaque année : les actifs de 1 à 5 ans (prêts du Grenier, créances) sont trop faibles pour suivre les sorties.",
                 f"**Même en ajoutant les {F('GAP 5+', _m['plus_5_ans_ou_incertain'])} d'actifs longs ou incertains (terrains, clinique, créance sur l'hôpital) à leur valeur comptable, l'écart cumulé reste négatif de {F('GAP fin (a)', -GAP[5][5])}.** Leur valeur de vente n'est d'ailleurs pas établie."],
        implication="En pratique, l'écart est comblé chaque année par les nouvelles cotisations d'épargne et les excédents : le FPPN dépend de la poursuite des versements pour payer les sorties.",
        valable="bilan provisoire au 31/08/2026, actifs classés par nature comptable ; taux de sortie moyen 2021 à 2025 ; sans nouveaux versements ni intérêts dus",
        source=SRC_COMPTA + " ; " + SRC_ALM["BAN"] + ", p. 86 et 87 ; onglet 24b du classeur des calculs")
_SC = [("Central", T_MOY, 0.0, "Taux de sortie moyen 2021 à 2025"), ("Résistance", T_MAX, 0.0, "Taux de sortie maximal observé"),
       ("Résistance sévère", T_MAX, CHOC_MAL, f"Taux maximal et hausse des prestations maladie de {pc('ALM pic (sc)', PIC)}")]
_lig_sc = []
for cl in ("31/12/2025", "31/08/2026"):
    for nom, t, ch, desc in _SC:
        a_, b_, c_ = couverture(cl, t, ch)
        _lig_sc.append([cl, nom, desc, F(f"SC act {cl} {nom}", a_), F(f"SC bes {cl} {nom}", b_), F(f"SC éc {cl} {nom}", a_ - b_), pc(f"SC couv {cl} {nom}", c_)])
contenu({"titre": "Adossement : tests de résistance", "sous_titre": f"Même dans le scénario central, l'actif à moins d'un an couvre {pc('SC c26', couverture('31/08/2026', T_MOY)[2])} des paiements de l'année au 31/08/2026",
         "synthese": "Les chocs sont calibrés sur les extrêmes observés au FPPN, pour rester cohérents avec le reste du rapport"},
        "Besoin de l'année = dettes à moins d'un an + épargne × taux de sortie + choc de prestations. Couverture = actif à moins d'un an / besoin. Choc de prestations = prestations maladie 2025 × hausse annuelle la plus forte de 2024 et 2025 ; les hausses de 2021 à 2023 sont écartées car elles reflètent le rattrapage des factures impayées.",
        tableau={"titre": tab("Couverture des paiements de l'année selon trois scénarios"), "entetes": ["Clôture", "Scénario", "Hypothèse", "Actif à moins d'un an", "Besoin de l'année", "Écart", "Couverture"], "largeurs": [3.0, 3.6, 7.4, 4.4, 4.4, 4.4, 3.0], "taille": 8.5,
                 "lignes": _lig_sc},
        disposition="dessous",
        analyse=[f"**La couverture varie entre {pc('SC min', min(couverture(c, T_MAX, CHOC_MAL)[2] for c in ('31/12/2025', '31/08/2026')))} et {pc('SC max', max(couverture(c, T_MOY)[2] for c in ('31/12/2025', '31/08/2026')))} selon le scénario et la clôture.** Aucun scénario n'approche 100 %."],
        implication="Une réserve de liquidité égale aux sorties d'épargne d'une année au taux maximal observé est la cible minimale d'une politique actif-passif.",
        valable="bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; taux de sortie 2021 à 2025 ; prestations maladie 2023 à 2025",
        source=SRC_COMPTA + " ; onglet 24b du classeur des calculs")
contenu({"titre": "Adossement : gouvernance à mettre en place", "sous_titre": "Aucune des pratiques de gouvernance actif-passif de référence n'est en place au FPPN",
         "synthese": "Le conseil d'administration doit fixer les limites ; un comité doit suivre les écarts et en rendre compte"},
        "Les pratiques de référence sont celles de l'Association actuarielle internationale et de l'Institut des actuaires. Elles sont rapprochées de la situation constatée au FPPN.",
        tableau={"titre": tab("Gouvernance actif-passif : pratiques de référence et situation du FPPN"), "entetes": ["Pratique de référence", "Source", "Situation du FPPN", "Action proposée"], "largeurs": [7.6, 5.4, 9.0, 9.2], "taille": 8.5,
                 "lignes": [["Le conseil d'administration s'engage et approuve les limites de risque", "AAI, ch. 13, p. 15", "Aucune limite votée ; le directeur général décide seul des placements", "Faire voter une politique actif-passif et de placement"],
                            ["Un comité actif-passif, avec un mandat écrit", "AAI, ch. 13, p. 19 ; Lavoie, p. 11", "Pas de comité", "Créer un comité : direction, comptabilité, actuaire"],
                            ["Une politique qui encadre la gestion actif-passif", "Lavoie, p. 8", "Pas de politique écrite", "Fixer réserve de liquidité, actifs admis, rendement minimal"],
                            ["Un suivi régulier des écarts", "Gestion des risques bancaires, p. 86", "Pas d'échéancier de l'épargne, pas de tableau de suivi", "Écarts de liquidité trimestriels, présentés au comité"],
                            ["Des engagements connus", "AAI, ch. 13, p. 3", "Taux servi non arrêté, comptes individuels non transmis", "Arrêter le taux, tenir les comptes individuels"]]},
        disposition="dessous",
        analyse=["**L'absence de gouvernance explique les constats chiffrés** : sans limite ni suivi, l'épargne a financé des actifs longs sans que personne ne mesure l'effet sur la liquidité."],
        implication="La gouvernance est le levier le moins coûteux : elle ne demande ni vente d'actif ni hausse de cotisation.",
        valable="pratiques des documents de référence cités ; situation du FPPN constatée en entretien le 19/10/2026 et sur les documents reçus jusqu'au 22/10/2026",
        source="Source : " + SRC_ALM["IAA"] + " ; " + SRC_ALM["LAV"] + " ; " + SRC_ALM["BAN"] + " ; " + SRC_ENT14[9:])
contenu({"titre": "Conclusion : adossement actif-passif", "sous_titre": "Le FPPN n'est pas adossé : l'épargne exigible finance des actifs longs et peu liquides",
         "synthese": f"Couverture des paiements de l'année au 31/08/2026 : {pc('CA c', couverture('31/08/2026', T_MOY)[2])} (central) à {pc('CA s', couverture('31/08/2026', T_MAX, CHOC_MAL)[2])} (sévère)"},
        "Synthèse de l'adossement actif-passif : ce qui est démontré, avec quelle méthode, et ce qui permettrait de passer à un adossement en flux contractuels.",
        tableau={"titre": tab("Adossement actif-passif : réponse et niveau de preuve"), "entetes": ["Question", "Réponse", "Méthode", "Pour aller plus loin"], "largeurs": [6.0, 10.4, 5.6, 9.2], "taille": 8.5,
                 "lignes": [["La liquidité couvre-t-elle l'année qui vient ?", "Non : " + pc("CA c (t)", couverture("31/08/2026", T_MOY)[2]) + " dans le scénario central", "Écoulement conventionnel", "Échéancier de l'épargne"],
                            ["L'écart se résorbe-t-il dans les cinq ans ?", "Non : manque cumulé de " + F("CA gap", -GAP[4][5]) + " à cinq ans", "Écarts par année", "Classement contractuel des actifs"],
                            ["Le FPPN résiste-t-il à un choc ?", "Non : " + pc("CA s (t)", couverture("31/08/2026", T_MAX, CHOC_MAL)[2]) + " dans le scénario sévère", "Tests de résistance", "Données de sorties par motif"],
                            ["Le FPPN pilote-t-il son adossement ?", "Non : ni politique, ni comité, ni limites", "Pratiques de référence", "Décision du conseil d'administration"]]},
        disposition="dessous",
        analyse=["**Le non-adossement est démontré avec les données du FPPN**, quel que soit le scénario retenu ; seule l'ampleur exacte dépend de l'échéancier de l'épargne."],
        implication="Le plan de redressement doit traiter en priorité la liquidité de l'épargne : réserve de liquidité, arrêt du financement d'actifs longs par l'épargne, politique actif-passif votée.",
        valable="bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; comportements de sortie observés de 2021 à 2025",
        source="Source : diapositives " + ", ".join(ref(t) for t in ("Adossement : écarts de liquidité année par année", "Adossement : tests de résistance", "Adossement : gouvernance à mettre en place")))
