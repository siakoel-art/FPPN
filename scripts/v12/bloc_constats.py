# ---- v12 : synthèse des constats au format fait, norme ou référence, écart chiffré, effet pour le FPPN, criticité
_GR26 = GRD.get("2026-08-31", 0.0)
_PLG = V11["placements_gl"]
_RG = {(x["exercice"], x["garantie"]): x for x in V11["ratios_garantie"]}
_PP = V11["pidr_plafond"]
_BESOIN = psap_m_elargie(2024) + TARD_DEC
_ACT, _PAI, _COUV = couverture("31/08/2026", T_MOY)
SYN1 = [
    ["❑ Provisions\nmaladie et décès",
     f"Les provisions maladie (**{F('SY prov mal', PROV_MAL)}**) et obsèques (**{F('SY prov obs', PROV_OBS)}**) ont le même montant depuis le 01/09/2020. Seul mouvement : un changement de compte en 2024.",
     "Statuts, art. 89 : provisions recalculées à chaque clôture. Règlement d'exécution n° 004/2022, art. 10 : calcul dossier par dossier",
     f"Besoin mesuré : **{F('SY besoin', _BESOIN)}**, soit le montant que les provisions devaient couvrir à la clôture. Il comprend les soins de 2024 facturés en 2025, estimés à {F('SY besoin mal', psap_m_elargie(2024))} à partir des factures de 2024 repérées dans le journal SIN de 2025, et {F('SY besoin dec', TARD_DEC)} d'allocations de six décès de 2021 enregistrés en 2022. Les provisions dépassent ce besoin de **{F('SY exces', PROV_MAL + PROV_OBS - _BESOIN)}**",
     f"Résultat et fonds propres sous-évalués d'autant. Les allocations décès dues ({F('SY dette dec', DETTE_DEC26)} au 31/08/2026) sont déjà en dette aux comptes 41591000 et 41592000 : la provision obsèques les compte une seconde fois",
     "Critique"],
    ["❑ Épargne\nPEP et EFC",
     "Aucun intérêt n'est comptabilisé sur l'épargne du 01/01/2020 au 31/08/2026. Trois taux figurent dans trois documents (3,5 %, 5,5 % et 14 %), aucun n'est arrêté par une décision signée.",
     "Manuel des produits d'épargne et cahier des charges EFC-PEP de 2021 : épargne rémunérée. Règlement d'exécution n° 004/2022, art. 11 : provision des risques longs",
     f"Intérêts dus non comptabilisés : **{F('SY int', INT55)}** au 31/08/2026, aux taux de 3,5 % (PEP) et 5,5 % (EFC), soit **{pc('SY int part', INT55 / b[2026]['epargne_totale'])}** de l'épargne inscrite au bilan ({F('SY ep', b[2026]['epargne_totale'])})",
     "La dette envers les épargnants est sous-évaluée d'autant. Le FPPN ne peut pas remettre à chaque épargnant un relevé de ses droits",
     "Critique"],
    ["❑ Adossement\nactif-passif",
     f"Au 31/08/2026, les actifs disponibles à moins d'un an valent **{F('SY act', _ACT)}**. Les paiements attendus dans l'année, sorties d'épargne au taux observé de {pc('SY tsort', T_MOY)} par an comprises, valent **{F('SY pai', _PAI)}**.",
     "Règlement d'exécution n° 004/2022 : placements sûrs et liquides (art. 17), réserve de trésorerie d'un trimestre pour les risques longs (art. 19)",
     f"Couverture des paiements de l'année : **{pc('SY couv', _COUV)}**. Manque cumulé à cinq ans : **{F('SY gap', -GAP[4][5])}**. Réserve de trésorerie constituée : 0 FCFA",
     "Pour payer des retraits supérieurs à la trésorerie, le FPPN devrait vendre ou mobiliser des actifs longs : terrains, clinique, créance sur l'Hôpital de la Police",
     "Critique"],
    ["❑ Conformité\nprudentielle",
     f"**{NB_NC}** des 36 indicateurs de la Décision n° 05/2022 sont non conformes de 2022 à 2025, dont la réserve obligatoire ({val_ratio(RT['ROB'], 2025)} en 2025), la liquidité ({val_ratio(RT['LIQ'], 2025)}) et les fonds propres et assimilés ({val_ratio(RT['FEP'], 2025)}).",
     "Décision n° 05/2022 : réserve obligatoire de 6 à 9 mois de prestations, liquidité et fonds propres et assimilés au moins égaux à 1",
     f"En 2025 : il manque {N('SY rob ec', 6 - RT['ROB']['valeurs'][2025], 2)} mois de réserve obligatoire, {N('SY liq ec', 1 - RT['LIQ']['valeurs'][2025], 2)} point de liquidité et {N('SY fep ec', 1 - RT['FEP']['valeurs'][2025], 2)} point de fonds propres et assimilés",
     "Deux exercices de suite non conformes : retrait d'agrément possible (règlement n° 07/2009, art. 84) et sanctions de l'arrêté n° 00115/MEPS/CAB du 12/12/2023",
     "Critique"],
    ["❑ Fonds propres",
     f"Les fonds propres, résultat compris, sont négatifs à chaque clôture depuis 2020 : **{F('SY fp', b[2025]['fonds_propres'])}** au 31/12/2025 et **{F('SY fp26', b[2026]['fonds_propres'])}** au 31/08/2026.",
     "Décision n° 05/2022 : fonds propres et assimilés au moins égaux aux dettes",
     "Les pertes reportées dépassent l'ensemble des réserves",
     "Aucun coussin pour absorber un choc de prestations, une perte sur les prêts du Grenier ou une baisse de valeur des actifs",
     "Élevé"],
]


def _syn_diapo(_k, _lig, _suite):
    contenu({"titre": f"Synthèse des constats ({_k}/2)", "sous_titre": "Cinq constats critiques portent sur les engagements, la liquidité et la conformité" if _k == 1 else "Les autres constats touchent la sinistralité, les placements, les données et la gouvernance",
             "synthese": "Chaque constat donne le fait chiffré, la norme qui s'applique, l'écart mesuré et son effet pour le FPPN"},
            "Tableau de synthèse : le fait constaté et depuis quand, la norme ou la référence qui s'applique, l'écart chiffré, l'effet concret pour le FPPN et le niveau de criticité. Chiffres identiques à ceux du corps du rapport.",
            tableau={"titre": tab(f"Synthèse des constats ({_k}/2) : {_suite}"), "entetes": ["Section", "Constat", "Norme ou référence", "Écart chiffré", "Effet pour le FPPN", "Criticité"],
                     "largeurs": [2.8, 7.6, 5.6, 7.0, 6.2, 2.0], "taille": 7.5, "lignes": _lig},
            disposition="dessous",
            valable="chiffres du corps du rapport, de 2020 au 31/08/2026",
            source="Source : analyses Adenka, parties 3 à 10 du rapport. Provisions : balances de 2020 au 31/08/2026, journal SIN des grands livres des tiers 2023 à 2025, états de fin d'année des allocations décès. Épargne : manuel et cahier des charges des produits d'épargne")


_syn_diapo(1, SYN1, "provisions, épargne, adossement, conformité, fonds propres")
SYN2 = [
    ["❑ Sinistralité\nmaladie",
     f"La sinistralité maladie passe de **{pc('SY sc21', s[2021]['sc_maladie'])}** en 2021 à **{pc('SY sc25', s[2025]['sc_maladie'])}** en 2025 et **{pc('SY sc26', s[2026]['sc_maladie'])}** sur les huit mois de 2026.",
     "Repère : fourchette de 70 % à 80 % du taux réglementaire, qui est global (Décision n° 05/2022, § 1.2.1.2.5)",
     f"Ratio combiné maladie, frais compris : **{pc('SY rcm', _RG[(2025, 'Maladie')]['ratio_combine'])}** en 2025. Résultat technique maladie 2025 : **{F('SY rtm', _RG[(2025, 'Maladie')]['resultat_technique'])}**",
     "La cotisation maladie de 3,5 % du salaire de base ne couvre plus le coût complet de la garantie. Le décès, excédentaire, finance la maladie",
     "Élevé"],
    ["❑ Placements",
     f"En 2025, l'encours moyen réellement placé (dépôts à terme et FCP, lu jour par jour dans les grands livres) vaut **{F('SY pl', _PLG['2025']['placements_moyens'])}**.",
     "Règlement d'exécution n° 004/2022, art. 17 : placements sûrs, rentables, liquides et diversifiés, selon une politique votée",
     f"Soit **{pc('SY pl part', _PLG['2025']['placements_moyens'] / EP_DUE25)}** de l'épargne due au 31/12/2025 ({F('SY ep25', EP_DUE25)}). Rendement moyen de 2021 à 2025 : **{pc('SY rdm', V11['rendement_moyen_2021_2025'])}** par an, sous le taux de 5,5 % annoncé sur l'EFC",
     "Les produits financiers ne peuvent pas payer les intérêts dus sur l'épargne. [[Le directeur général décide seul des placements]]",
     "Élevé"],
    ["❑ Données",
     "Ni la base des prestations datée, ni les dates de naissance des mutualistes ne sont reçues. Les fichiers maladie et décès de décembre 2024 sont vides (0 octet).",
     f"PBC du 23/10/2026, n° {PBCM[9]}, {PBCM[62]} et {PBCM[71]}. Règlement d'exécution n° 004/2022, art. 7 à 10 : provisions calculées dossier par dossier",
     f"Date de naissance absente des fichiers de précomptes maladie et décès. Connue pour {N('SY rdn', RDN['nb'])} retraités seulement, dans les fichiers de la CGRAE de février à avril 2024",
     "Mortalité observée, engagement santé des retraités et projection par âge impossibles à calculer",
     "Élevé"],
    ["❑ Gouvernance",
     "[[Le directeur général décide seul des placements]]. Le projet de Cocody n'a ni budget ni plan de financement transmis. La désignation du directeur général n'est documentée que pour 2020 à 2023.",
     "Statuts, art. 85 : délibération du conseil d'administration pour chaque placement",
     "Aucune délibération de placement reçue",
     "L'épargne des adhérents est employée sans décision des organes",
     "Élevé"],
    ["❑ Prêts du Grenier",
     f"L'encours des prêts vaut **{F('SY gre', GP['encours'])}** au 31/08/2026. Aucune dépréciation n'est passée depuis 2022.",
     "Politique de crédit du Grenier : dépréciation de 40 %, 80 % ou 100 % de l'encours selon l'ancienneté du retard",
     f"Dépréciation nécessaire au 31/08/2026 : **{F('SY dep', _GR26)}**",
     "Actif et fonds propres surévalués de ce montant. L'épargne des adhérents finance l'essentiel du crédit",
     "Moyen"],
    ["❑ Personnel\ndu FPPN",
     f"Le compte 196 porte **{F('SY 196', _PP['provision_196'])}** au 31/12/2025 pour les indemnités de départ à la retraite du personnel du FPPN.",
     "Décret n° 2017-210, art. 4 et 5 : indemnité de départ. Convention collective, art. 40 : plafond de 25 fois le SMIG annuel",
     f"Recalcul agent par agent : **{F('SY pidr', _PP['dbo_sans_plafond'])}** sans plafond, **{F('SY pidr cap', _PP['dbo_avec_plafond'])}** avec le plafond. Le compte 196 dépasse le recalcul de {F('SY pidr ec', _PP['provision_196'] - _PP['dbo_sans_plafond'])}",
     "Provision suffisante. La méthode du FPPN n'est pas documentée",
     "Faible"],
]
