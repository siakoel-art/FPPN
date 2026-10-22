# ---- v10 : intérêts dus sur l'épargne depuis le 01/01/2020 (partie 6)
_EI = [e for e in V10["epargne"]]


def _ip(e, t):
    return (e["pep_debut"] + e["pep_fin"]) / 2 * t * e["mois"] / 12


def _ie(e, t):
    return (e["efc_debut"] + e["efc_fin"]) / 2 * t * e["mois"] / 12


contenu({"titre": "Épargne : intérêts dus non comptabilisés", "sous_titre": f"Du 01/01/2020 au 31/08/2026, les intérêts dus selon le manuel d'épargne atteignent {F('EI total', INT55)}",
         "synthese": f"Engagement d'épargne au 31/08/2026 : {F('EI dû', EP_DUE26)} au bilan, {F('EI engagement', EP_DUE26 + INT55)} avec les intérêts dus (PEP 3,5 %, EFC 5,5 %)"},
        "Intérêts de l'année = épargne moyenne de l'année (moyenne du début et de la fin) × taux servi × nombre de mois / 12. Taux du manuel des produits d'épargne : PEP 3,5 %, EFC 5,5 %. Les intérêts ne sont pas capitalisés d'une année sur l'autre : le montant est donc un minimum. Pour 2020, le premier solde disponible est celui du 01/09/2020 (à-nouveaux de la balance 2020) : la moyenne du 01/09 et du 31/12/2020 est appliquée à toute l'année.",
        tableau={"titre": tab("Intérêts dus sur l'épargne, par exercice"), "entetes": ["Exercice", "PEP moyen", "EFC moyenne", "Intérêts PEP (3,5 %)", "Intérêts EFC (5,5 %)", "Total de l'année", "Total si EFC à 3,5 %"],
                 "largeurs": [3.0, 4.8, 4.8, 4.4, 4.4, 4.6, 4.6], "taille": 8.5,
                 "lignes": [[str(e["exercice"]) if e["exercice"] < 2026 else "2026 (8 mois)", F(f"EI pep moy {e['exercice']}", (e["pep_debut"] + e["pep_fin"]) / 2), F(f"EI efc moy {e['exercice']}", (e["efc_debut"] + e["efc_fin"]) / 2),
                             F(f"EI ip {e['exercice']}", _ip(e, 0.035)), F(f"EI ie {e['exercice']}", _ie(e, 0.055)), F(f"EI tot {e['exercice']}", _ip(e, 0.035) + _ie(e, 0.055)),
                             F(f"EI tot35 {e['exercice']}", _ip(e, 0.035) + _ie(e, 0.035))] for e in _EI]
                 + [["Total", "", "", F("EI ip tot", sum(_ip(e, 0.035) for e in _EI)), F("EI ie tot", sum(_ie(e, 0.055) for e in _EI)), F("EI tot (t)", INT55), F("EI tot35 (t)", INT35)]]},
        disposition="dessous",
        analyse=[f"**Selon le taux retenu pour l'EFC, les intérêts dus depuis le 01/01/2020 sont compris entre {F('EI bas', INT35)} et {F('EI haut', INT55)}.** Aucun de ces montants n'est comptabilisé : il n'existe ni charge d'intérêts ni dette d'intérêts dans les balances.",
                 f"**Au 31/12/2025, les intérêts dus atteignent déjà {F('EI 2025', INT55_25)}**, soit {N('EI fois pf', INT55_25 / sum(x['produits_financiers'] for x in R['rendement_placements']), 1)} fois les produits financiers du FPPN de 2021 à 2025 ({F('EI pf', sum(x['produits_financiers'] for x in R['rendement_placements']))})."],
        implication="Les intérêts dus sont une dette envers les mutualistes. Le taux servi doit être arrêté par une décision signée, puis les intérêts calculés compte par compte et comptabilisés chaque année.",
        valable="soldes des comptes 41195100 (PEP) et 41195200 (EFC) du 01/09/2020 au 31/08/2026 (balances 2020 à 2025 et provisoire 2026), 2020 approché par les soldes du 01/09 et du 31/12 ; taux du manuel d'épargne ; intérêts antérieurs au 01/01/2020 et capitalisation non comptés",
        source=SRC_EP + " ; " + SRC_COMPTA[9:] + " ; onglet 21 du classeur des calculs")
