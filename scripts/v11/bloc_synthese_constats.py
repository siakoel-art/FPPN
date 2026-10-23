# ---- v11 : synthèse des constats (résumé exécutif) : section, constat, impact, origine et sources, criticité
_GR26 = GRD.get("2026-08-31", 0.0)
_PLG = V11["placements_gl"]
_RG = {(x["exercice"], x["garantie"]): x for x in V11["ratios_garantie"]}
_PP = V11["pidr_plafond"]
SYN1 = [
    ["❑ Provisions\nmaladie et décès",
     f"• Maladie **{F('SY prov mal', PROV_MAL)}** et obsèques **{F('SY prov obs', PROV_OBS)}**, identiques de 2020 au 31/08/2026\n• Besoin mesuré : **{F('SY besoin', psap_m_elargie(2024) + TARD_DEC)}**, hors dettes déjà comptabilisées",
     "Provisions non calculées : statuts, art. 89, et règlement d'exécution n° 004/2022, art. 10, non respectés. La provision obsèques double la dette du compte 41591000",
     "Balances de 2020 au 31/08/2026 (comptes 19011000, 19012000, 15801000, 15802000). Journal SIN des grands livres des tiers 2023 à 2025. États de fin d'année des allocations décès 2021 à 2025", "Critique"],
    ["❑ Épargne\nPEP et EFC",
     f"• Intérêts dus non comptabilisés : **{F('SY int', INT55)}** du 01/01/2020 au 31/08/2026\n• Trois taux écrits (3,5 %, 5,5 %, 14 %), aucun arrêté par décision signée",
     "La dette d'épargne inscrite au bilan est sous-évaluée. L'engagement envers chaque épargnant n'est pas démontrable sans comptes individuels",
     "« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx », « CAHIER DE CHARGEEFC & PEP 2021.pdf », « GRAND LIVRE FPPN 2024.xlsx » (comptes 41195100 et 41195200), balances 2020 à 2026", "Critique"],
    ["❑ Adossement\nactif-passif",
     f"• L'actif à moins d'un an couvre **{pc('SY couv', couverture('31/08/2026', T_MOY)[2])}** des paiements de l'année (scénario central)\n• Manque cumulé de **{F('SY gap', -GAP[4][5])}** à cinq ans",
     "Le FPPN ne peut pas payer les sorties d'épargne sans mobiliser des actifs longs (terrains, clinique, créances)",
     "Balance provisoire au 31/08/2026. Débits des comptes d'épargne de 2021 à 2025 (sorties observées)", "Critique"],
    ["❑ Conformité\nprudentielle",
     f"• **{NB_NC}** indicateurs non conformes sur 36, dont la réserve obligatoire, les fonds propres, la liquidité et les frais\n• Deux exercices de suite non conformes pour plusieurs ratios",
     "Retrait d'agrément possible après deux exercices non conformes (règlement n° 07/2009, art. 84) et sanctions de l'arrêté n° 00115/MEPS/CAB",
     "Décision n° 05/2022/COM/UEMOA, annexe. Balances 2022 à 2025. Arrêté n° 00115/MEPS/CAB du 12/12/2023", "Critique"],
    ["❑ Fonds propres",
     f"• Fonds propres de **{F('SY fp', b[2025]['fonds_propres'])}** au 31/12/2025\n• Pertes reportées supérieures à l'ensemble des réserves",
     "Aucun coussin pour absorber un choc de prestations ou une perte sur actifs",
     "Balances 2020 à 2025 (comptes 10 à 13)", "Élevé"],
]
def _syn_diapo(_k, _lig, _suite):
    contenu({"titre": f"Synthèse des constats ({_k}/2)", "sous_titre": "Cinq constats critiques portent sur les engagements, la liquidité et la conformité" if _k == 1 else "Les autres constats touchent la sinistralité, les placements, les données et la gouvernance",
             "synthese": "Chaque constat est démontré dans le corps du rapport à partir des fichiers et des textes cités en colonne « Origine et sources »"},
            "Tableau de synthèse au format du rapport de diagnostic : section, constat chiffré, impact, origine du constat et niveau de criticité. Chiffres identiques à ceux du corps du rapport.",
            tableau={"titre": tab(f"Synthèse des constats ({_k}/2) : {_suite}"), "entetes": ["Section", "Constat", "Impact", "Origine et sources", "Niveau de criticité"],
                     "largeurs": [3.6, 10.6, 7.4, 7.0, 2.6], "taille": 8, "lignes": _lig},
            disposition="dessous",
            valable="chiffres du corps du rapport, de 2020 au 31/08/2026 ; criticité appréciée selon la grille ci-dessous",
            source="Source : analyses Adenka, parties 3 à 10 du rapport ; fichiers et textes cités dans chaque ligne")


_syn_diapo(1, SYN1, "provisions, épargne, adossement, conformité, fonds propres")
SYN2 = [
    ["❑ Sinistralité\nmaladie",
     f"• S/C maladie : **{pc('SY sc25', s[2025]['sc_maladie'])}** en 2025 et **{pc('SY sc26', s[2026]['sc_maladie'])}** sur huit mois de 2026\n• Résultat technique maladie 2025 : **{F('SY rtm', _RG[(2025, 'Maladie')]['resultat_technique'])}**",
     "La garantie maladie est déficitaire, frais compris. Le décès, excédentaire, finance la maladie",
     "Balances 2020 au 31/08/2026 (comptes 601 à 6024, 7010, 702, 612 à 68)", "Élevé"],
    ["❑ Placements",
     f"• Encours moyen placé en 2025 : **{F('SY pl', _PLG['2025']['placements_moyens'])}**, soit **{pc('SY pl part', _PLG['2025']['placements_moyens'] / EP_DUE25)}** de l'épargne due\n• Rendement moyen 2021 à 2025 : **{pc('SY rdm', V11['rendement_moyen_2021_2025'])}**, aucune politique de placement",
     "Les produits financiers ne couvrent pas les intérêts dus sur l'épargne. Aucune des conditions de l'article 17 du règlement d'exécution n° 004/2022 n'est démontrée",
     "Grands livres des comptes 2021 au 31/08/2026 (comptes 5221, 5330, 77). [[Entretien du chef comptable du 19/10/2026]]", "Élevé"],
    ["❑ Données",
     "• Base des prestations datée et dates de naissance des mutualistes non reçues\n• Décembre 2024 vide pour la maladie et le décès",
     "Mortalité observée, provision santé des retraités et projection par âge non calculables",
     f"PBC du 23/10/2026, n° {PBCM[9]}, {PBCM[62]} et {PBCM[71]}", "Élevé"],
    ["❑ Gouvernance",
     "• [[Le directeur général décide seul des placements]]\n• Projet de Cocody sans budget ni financement reçus. Désignation du directeur général reçue pour 2020 à 2023 seulement",
     "Les statuts (art. 85) exigent une délibération du conseil d'administration pour chaque placement",
     "[[Entretien du chef comptable du 19/10/2026]]. Dérogation d'urbanisme de janvier 2025. Attestation de désignation du 04/09/2020", "Élevé"],
    ["❑ Prêts du Grenier",
     f"• Encours de **{F('SY gre', GP['encours'])}** au 31/08/2026, aucune dépréciation passée\n• Dépréciation selon la politique du Grenier : **{F('SY dep', _GR26)}**",
     "Actif et fonds propres légèrement surévalués. L'épargne finance l'essentiel du crédit",
     "Exports « BalanceAgeeDetaille … .xls » du logiciel de prêts. « POLITIK DE CREDIT DU GRENIER VFVF.doc »", "Moyen"],
    ["❑ Personnel\ndu FPPN",
     f"• Compte 196 : **{F('SY 196', _PP['provision_196'])}**. Recalcul agent par agent : **{F('SY pidr', _PP['dbo_sans_plafond'])}**, **{F('SY pidr cap', _PP['dbo_avec_plafond'])}** avec le plafond conventionnel",
     "Provision suffisante. Méthode du FPPN non documentée",
     "DISA 2021 à 2025 (CNPS). Décret n° 2017-210 du 30/03/2017. Convention collective du 19/07/1977, art. 40. Décret n° 2022-986 du 21/12/2022 (SMIG)", "Faible"],
]
_syn_diapo(2, SYN2, "sinistralité, placements, données, gouvernance, Grenier, personnel")
