# ---- v11 : ajouts au format du rapport de diagnostic : méthode du FPPN et recalcul, enjeux du provisionnement, dispositif actuariel
_PP = V11["pidr_plafond"]
_GR26 = GRD.get("2026-08-31", 0.0)
_RECALC = [
    ("Provision maladie (15801000)", "Montant hérité, inchangé depuis 2020", PROV_MAL, psap_m_elargie(2024), "Soins facturés l'année suivante, journal SIN"),
    ("Provision obsèques (15802000)", "Montant hérité, inchangé depuis 2020", PROV_OBS, TARD_DEC, "Décès déclarés après la clôture (dette connue au 41591000)"),
    ("Départ à la retraite du personnel (196)", "Reprise et dotation annuelles, méthode non transmise", _PP["provision_196"], _PP["dbo_sans_plafond"], "Unités de crédit projetées, agent par agent"),
    ("Intérêts dus sur l'épargne", "Aucune écriture", 0.0, INT55, "Taux du manuel d'épargne, du 01/01/2020 au 31/08/2026"),
    ("Dépréciation des prêts du Grenier", "Aucune écriture depuis 2022", 0.0, _GR26, "Taux de la politique de crédit du Grenier"),
]


def _ecart_rel(fppn, recalc):
    return pc(f"MR rel {fppn} {recalc}", (fppn - recalc) / recalc) if recalc else "Sans objet"


contenu({"titre": "Méthode du FPPN et recalcul", "sous_titre": "Le FPPN a provisionné trop pour la maladie et le décès, et rien pour les intérêts et les prêts",
         "synthese": f"Écart total sur les cinq postes : {F('MR total', sum(f - r for _, _, f, r, _ in _RECALC))}, montant du FPPN moins recalcul"},
        "Pour chaque provision, le tableau compare le montant du FPPN au 31/08/2026 au recalcul du rapport, avec la méthode utilisée. Écart absolu = FPPN - recalcul. Écart relatif = écart absolu / recalcul.",
        tableau={"titre": tab("Montants du FPPN, recalcul et écarts"), "entetes": ["Poste", "Méthode du FPPN", "Montant du FPPN", "Recalcul", "Méthode du recalcul", "Écart absolu", "Écart relatif"],
                 "largeurs": [5.6, 5.0, 4.0, 4.0, 6.0, 4.2, 2.4], "taille": 8.5,
                 "lignes": [[p, m, F(f"MR fppn {p}", f), F(f"MR rec {p}", r), mr, F(f"MR abs {p}", f - r), _ecart_rel(f, r)] for p, m, f, r, mr in _RECALC]},
        disposition="dessous",
        analyse=["**Les provisions maladie et obsèques dépassent largement les besoins mesurés, alors que les intérêts dus et la dépréciation des prêts ne sont pas provisionnés.** Le total des provisions n'est pas insuffisant : il est mal affecté."],
        implication="Le reclassement des montants hérités vers les intérêts dus sur l'épargne suppose une décision des organes et l'information de l'AIRMS (recommandation 1).",
        valable="bilan provisoire au 31/08/2026 ; recalculs des diapositives de la partie 9 ; provision du personnel au 31/12/2025",
        source=SRC_COMPTA + " ; " + ", ".join(ref(t) for t in ("Provision maladie : montant nécessaire", "Allocations décès : décès connus après la clôture", "Provision du personnel par entité de 2021 à 2025", "Épargne : intérêts dus non comptabilisés", "Prêts du Grenier : dépréciation selon la politique de crédit")))
contenu({"titre": "Enjeux d'un provisionnement mal calibré", "sous_titre": "Une provision trop forte masque les vrais engagements, une provision trop faible surévalue les fonds propres",
         "synthese": "Au FPPN, les deux erreurs coexistent : maladie et décès surprovisionnés, épargne et prêts sous-provisionnés"},
        "Le tableau décrit l'effet de chaque erreur de provisionnement sur les comptes et les ratios, puis la situation constatée au FPPN.",
        tableau={"titre": tab("Effets d'une provision mal calibrée et situation du FPPN"), "entetes": ["Erreur", "Effet sur les comptes", "Effet sur les ratios et les décisions", "Situation au FPPN"],
                 "largeurs": [4.6, 8.4, 9.0, 9.2], "taille": 8.5,
                 "lignes": [["Provision trop forte", "Résultat et fonds propres sous-évalués. Charges futures déjà couvertes deux fois", "Sinistralité et frais mal mesurés. Cotisations jugées insuffisantes à tort", f"Maladie et obsèques : {F('EN trop', PROV_MAL + PROV_OBS - psap_m_elargie(2024) - TARD_DEC)} au-dessus des besoins mesurés"],
                            ["Provision trop faible", "Dettes futures absentes du bilan. Résultat et fonds propres surévalués", "Marge de solvabilité et réserve obligatoire surévaluées. Tarifs et placements décidés sur une base fausse", f"Intérêts dus non provisionnés : {F('EN int', INT55)}. Prêts non dépréciés : {F('EN gre', _GR26)}"],
                            ["Provision non recalculée", "Montant sans lien avec les engagements de l'année", "Impossible de démontrer la suffisance à l'AIRMS ou au commissaire aux comptes", "Montants identiques de 2020 au 31/08/2026 (statuts, art. 89, non appliqué)"]]},
        disposition="dessous",
        analyse=["**Le risque principal du FPPN n'est pas un manque global de provisions, mais leur affectation.** Les engagements d'épargne, les plus importants, sont les moins provisionnés."],
        implication="Une note de calcul annuelle, signée et présentée à l'Assemblée générale, corrige les trois erreurs à la fois (recommandation 1).",
        valable="provisions et recalculs au 31/08/2026",
        source="Source : analyses Adenka, partie 9 ; " + REX + ", art. 7 à 11 ; statuts du FPPN édités le 15/06/2023, art. 89")
