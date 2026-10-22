# ---- v10 : données de la revue actuarielle complète et renvois entre diapositives
V10 = json.load(open(V10J, encoding="utf-8"))


def ref(titre):
    """Renvoi vers une diapositive, résolu en numéro une fois toutes les diapositives créées."""
    return "§§" + titre + "§§"


_E10 = {e["exercice"]: e for e in V10["epargne"]}


def interets_dus(t_pep, t_efc, jusqu_a=2026):
    """Intérêts dus non comptabilisés du 01/01/2020 à la fin de l'exercice indiqué (non capitalisés, encours moyen)."""
    tot = 0.0
    for a in range(2020, jusqu_a + 1):
        e = _E10[a]
        tot += ((e["pep_debut"] + e["pep_fin"]) / 2 * t_pep + (e["efc_debut"] + e["efc_fin"]) / 2 * t_efc) * e["mois"] / 12
    return tot


EP_DUE26 = _E10[2026]["pep_fin"] + _E10[2026]["efc_fin"]
EP_DUE25 = _E10[2025]["pep_fin"] + _E10[2025]["efc_fin"]
INT55 = interets_dus(0.035, 0.055)
INT35 = interets_dus(0.035, 0.035)
INT55_25 = interets_dus(0.035, 0.055, 2025)
DST = {x["exercice"]: x for x in V10["deces_stock"]}
DETTE_DEC26 = (DST[2026]["ouv_41591000"] + DST[2026]["mvt_c_41591000"] - DST[2026]["mvt_d_41591000"]) + (DST[2026]["ouv_41592000"] + DST[2026]["mvt_c_41592000"] - DST[2026]["mvt_d_41592000"])
PROV_OBS = PV["psap_deces"]["2026"]
PROV_MAL = PV["psap_maladie"]["2026"]
TARD_DEC = sum(x["montant"] for x in V10["deces_tardifs"])
_PM = {x["exercice"]: x for x in V10["psap_maladie_delais"]}
TARDM = {x["cloture"]: x for x in V10["psap_maladie_tardives"]}


def psap_m_elargie(cl):
    x, pa = TARDM[cl], _PM[cl + 1]
    return x["references_annee_precedente"] / (x["references_identifiees"] / x["factures_sinistres"]) * (pa["prestations_maladie"] / pa["factures_enregistrees"])


DELAI = {a: _PM[a]["dettes_cliniques"] / (_PM[a]["factures_enregistrees"] / _PM[a]["mois"]) for a in range(2020, 2027)}
GRD = {}
for _g in V10["grenier_depreciation"]:
    GRD[_g["arrete"]] = GRD.get(_g["arrete"], 0.0) + _g["encours"] * _g["taux"]
SORT10 = V10["epargne_sorties"]
T_MOY = sum(x["total_taux"] for x in SORT10) / len(SORT10)
T_MAX = max(x["total_taux"] for x in SORT10)
MOB10 = {m["cloture"]: m for m in V10["mobilisation"]}


def couverture(cl, taux, choc=0.0):
    m = MOB10[cl]
    actif = m["immediat"] + m["moins_1_an"]
    besoin = m["passif_moins_1_an"] + m["epargne_echeance_a_etablir"] * taux + choc
    return actif, besoin, actif / besoin


PIC = V10["pic_prestations"]
CHOC_MAL = s[2025]["prest_maladie"] * PIC
PDV = PV["pidr"]
PIDR25 = V9["pidr_synthese"][-1]
SRC_LOT4 = "« PROV ALLOCATION DECES 2021.xlsx », « PROV ALLOCATION DECES 2022.xlsx », « ALLOC DECES 2023 PAYER ET A PROVISIONNER .xlsx », « PROV ALLOCATION DECES 2024.xlsx », « PROV ALLOCATIONS DECES ET FRAIS FUNERAIRES 2025.xlsx »"
SRC_LOT4C = "états de fin d'année des allocations décès 2021 à 2025 (« PROV ALLOCATION DECES 2021.xlsx » et suivants, voir l'annexe « Fichiers sources »)"
SRC_GLT = "grands livres des tiers 2023 à 2025 (« GRAND LIVRE DES TIERS FPPN 2023.xlsx » et suivants)"
SRC_RI = "règlement intérieur du FPPN édité le 15/06/2023"
SRC_ST = "statuts du FPPN édités le 15/06/2023"
RDN = V10["retraites_naissance"]
