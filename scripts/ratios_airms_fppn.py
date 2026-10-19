# -*- coding: utf-8 -*-
"""
FPPN - Ratios prudentiels de la mutualité sociale UEMOA appliqués aux comptes du FPPN
=====================================================================================

Textes appliqués (documentation AIRMS fournie le 18/10/2026) :
- Règlement n° 07/2009/CM/UEMOA du 26 juin 2009 (art. 60 à 63, 83, 84) ;
- Règlement d'exécution n° 004/2022/COM/UEMOA du 1er juin 2022 (art. 3, 7 à 27) ;
- Décision n° 05/2022/COM/UEMOA du 1er juin 2022 et son annexe (indicateurs et ratios, comptes PCMS) ;
- Plan comptable des mutuelles sociales (PCMS, mai 2012) ;
- Arrêté n° 00115/MEPS/CAB du 12 décembre 2023 (sanctions administratives).

Méthode en deux temps, comme demandé :
1. « Comme le FPPN » : les comptes sont pris tels qu'ils sont tenus et classés selon les numéros du PCMS
   que la Décision n° 05/2022 cite pour chaque ratio.
2. « Retraité » : certains comptes sont reclassés selon leur nature économique (épargne des adhérents,
   avances au CMP et au Grenier, frais répartis par risque), chaque retraitement étant motivé.

Utilisation : python ratios_airms_fppn.py <dossier_base> <resultats.json> <sortie_dossier>
"""
import json
import os
import sys

import numpy as np
import pandas as pd

BASE, RES, SORTIE = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(SORTIE, exist_ok=True)
EX = list(range(2020, 2027))
R = json.load(open(RES, encoding="utf-8"))

bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal["nature"] == "comptes"].copy()
bal["net"] = bal["sol_d"] - bal["sol_c"]
LIB = bal.groupby("compte")["intitule"].last().str.strip()


def solde(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["net"].sum())


def debiteurs(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["net"].clip(lower=0).sum())


def crediteurs(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float((-g["net"]).clip(lower=0).sum())


def mvt_d(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["mvt_d"].sum())


# ---------------------------------------------------------------------------
# 1. Agrégats selon les comptes PCMS cités par la Décision n° 05/2022 (annexe, § 1.2)
# ---------------------------------------------------------------------------
A = {}
for ex in EX:
    a = {}
    a["cotisations_acquises"] = -solde(ex, r"^70") - solde(ex, r"^7906") + solde(ex, r"^6906")
    a["cot_maladie"] = -solde(ex, r"^701")
    a["cot_deces"] = -solde(ex, r"^702")
    a["charges_prestations"] = solde(ex, r"^60") + solde(ex, r"^6901") + solde(ex, r"^7901")
    a["prest_maladie"] = solde(ex, r"^(601|6013|6018|60800007)")
    a["prest_deces"] = solde(ex, r"^(6021|6024)")
    a["prime_assurance_6518"] = solde(ex, r"^6518")
    a["charges_fonctionnement"] = solde(ex, r"^(612|613|618|62|63|64|65|66|67|68)")
    a["charges_fonctionnement_brut"] = solde(ex, r"^(61|62|63|64|65|66|67|68|691|697)")
    a["charges_personnel"] = solde(ex, r"^66")
    a["charges_personnel_641"] = solde(ex, r"^641")
    a["cout_formation"] = solde(ex, r"^6391")
    a["total_produits"] = -sum(solde(ex, m) for m in (r"^70", r"^71", r"^72", r"^73", r"^75", r"^77", r"^781",
                                                     r"^782", r"^791", r"^797", r"^82", r"^84", r"^86", r"^88")) \
        - solde(ex, r"^7906") + solde(ex, r"^6906")
    a["subventions"] = -solde(ex, r"^(75[1-8]|88)")
    a["total_charges"] = solde(ex, r"^6") + solde(ex, r"^(81|83|85)")
    a["produits_7"] = -solde(ex, r"^7")
    a["charges_6"] = solde(ex, r"^6")
    a["resultat_hao"] = -solde(ex, r"^8")
    a["resultat_net"] = a["produits_7"] - a["charges_6"] + a["resultat_hao"]
    # Bilan
    a["fonds_mutualistes_10"] = -solde(ex, r"^10")
    a["reserve_disponible_106"] = -solde(ex, r"^106")
    a["reserves_legales_111"] = -solde(ex, r"^111")
    a["reserves_facultatives_118"] = -solde(ex, r"^118")
    a["ran_crediteur_121"] = -solde(ex, r"^121")
    a["ran_debiteur_129"] = -solde(ex, r"^129")
    a["subventions_invest_14"] = -solde(ex, r"^14")
    a["provisions_techniques"] = -solde(ex, r"^(190|1580[12])")
    a["provisions_reglementees_15"] = -solde(ex, r"^15")
    a["provisions_190"] = -solde(ex, r"^190")
    a["provision_pensions_196"] = -solde(ex, r"^196")
    a["emprunts_16"] = -solde(ex, r"^16")
    a["epargne_pep"] = -solde(ex, r"^411951")
    a["epargne_efc"] = -solde(ex, r"^411952")
    a["epargne"] = a["epargne_pep"] + a["epargne_efc"]
    a["sorties_epargne"] = mvt_d(ex, r"^41195[12]")
    a["actif_immobilise_net"] = solde(ex, r"^2")
    a["emplois_longs_tiers"] = solde(ex, r"^(47111100|47111101|47111102|47111104|47111105|47111201|485)")
    a["creance_hopital"] = solde(ex, r"^47111100")
    a["cmp"] = solde(ex, r"^(47111101|47111104|47111105)")
    a["grenier"] = solde(ex, r"^47111102")
    a["tresorerie_actif"] = debiteurs(ex, r"^5")
    a["tresorerie_passif"] = crediteurs(ex, r"^5")
    a["disponibilites"] = debiteurs(ex, r"^(521|531|532|57)")
    a["placements_tresorerie"] = debiteurs(ex, r"^(522|526|533|50)")
    a["comptes_courants_521"] = solde(ex, r"^521")
    a["creances_ct"] = debiteurs(ex, r"^(3|4)") - a["emplois_longs_tiers"]
    a["dettes_ct_pcms"] = crediteurs(ex, r"^4") + a["tresorerie_passif"]
    a["dettes_ct_hors_epargne"] = a["dettes_ct_pcms"] - a["epargne"]
    a["dettes_fournisseurs_cliniques"] = -solde(ex, r"^40110")
    a["actifs_dedies_epargne"] = solde(ex, r"^(5211020|5221020|5211030|5221030|526120)") + solde(ex, r"^5330")
    a["prets_pep"] = solde(ex, r"^(2711|2712|4111031)")
    a["interets_dat_fcp"] = -solde(ex, r"^77")
    A[ex] = a
for ex in EX:
    A[ex]["var_provisions_techniques"] = (A[ex]["provisions_techniques"] - A[ex - 1]["provisions_techniques"]) if ex > 2020 else 0.0

# Fonds propres « comme le FPPN » (somme des classes 10 à 14 et du résultat de l'exercice)
for ex in EX:
    a = A[ex]
    a["fonds_propres"] = (a["fonds_mutualistes_10"] + a["reserves_legales_111"] + a["reserves_facultatives_118"]
                          + a["ran_crediteur_121"] + a["ran_debiteur_129"] + a["subventions_invest_14"] + a["resultat_net"])
    # contrôle : identique au calcul de la v2 (analyses_fppn.py)
    v2 = next(b for b in R["bilan"] if b["exercice"] == ex)["fonds_propres"]
    assert abs(a["fonds_propres"] - v2) < 1, (ex, a["fonds_propres"], v2)

# Contrôles de cohérence avec les agrégats déjà publiés (v2)
for s in R["sinistralite"]:
    ex = s["exercice"]
    assert abs(A[ex]["resultat_net"] - s["resultat"]) < 1
    assert abs(A[ex]["prest_maladie"] - s["prest_maladie"]) < 1

# ---------------------------------------------------------------------------
# 2. Effectifs : fidélisation et croissance (fichiers du poste 936, 2024 à 2026)
# ---------------------------------------------------------------------------
pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
pre = pre[pre.retenu & (pre.poste == "936") & (pre.sens == "retour") & ((pre.annee < 2026) | (pre.mois <= 8))]
mat = {a: set(pre[pre.annee == a].mat6.dropna()) for a in (2024, 2025, 2026)}
EFF = {"adherents_2024": len(mat[2024]), "adherents_2025": len(mat[2025]), "adherents_2026": len(mat[2026]),
       "fidelisation_2025": len(mat[2024] & mat[2025]) / len(mat[2024]),
       "renouvellements_2025": len(mat[2024] & mat[2025]),
       "fidelisation_2026": len(mat[2025] & mat[2026]) / len(mat[2025]),
       "renouvellements_2026": len(mat[2025] & mat[2026]),
       "croissance_2025": len(mat[2025]) / len(mat[2024]) - 1}
cg = next(x for x in R["aller_retour"] if x["poste"] == "CGRAE")
EFF["recouvrement_cgrae"] = cg["taux_recouvrement"]
EFF["recouvrement_cgrae_mois"] = cg["mois_communs"]

# ---------------------------------------------------------------------------
# 3. Ratios : définition, formule, comptes, norme, source, priorité
# ---------------------------------------------------------------------------
DEC = "Décision n° 05/2022/COM/UEMOA, annexe"
RE = "Règlement d'exécution n° 004/2022/COM/UEMOA"


def r_(n, d):
    return n / d if d not in (0, None) and n is not None else None


RATIOS = []


def ratio(code, nom, famille, priorite, formule, comptes, norme_txt, test, source, valeurs, num=None, den=None,
          unite="%", qualitatif=None):
    RATIOS.append({"code": code, "nom": nom, "famille": famille, "priorite": priorite, "formule": formule,
                   "comptes": comptes, "norme": norme_txt, "test": test, "source": source, "unite": unite,
                   "valeurs": valeurs, "num": num or {}, "den": den or {}, "qualitatif": qualitatif})


def serie(f):
    return {ex: f(A[ex], ex) for ex in EX}


# --- Priorité actuarielle : solvabilité, provisions, sinistralité, adossement ---
ratio("SIN", "Taux de sinistralité", "Actuariel", 1,
      "Charges de prestations (avec provisions) / Cotisations acquises",
      "60 + 6901 - 7901 / 70 + 7906 - 6906", "Entre 70 % et 80 %", ("entre", 0.70, 0.80),
      DEC + ", tableau 1, ligne 5 et § 1.2.1.2.5",
      serie(lambda a, ex: r_(a["charges_prestations"], a["cotisations_acquises"])),
      serie(lambda a, ex: a["charges_prestations"]), serie(lambda a, ex: a["cotisations_acquises"]))
ratio("MSE", "Marge de solvabilité disponible / exigée", "Actuariel", 2,
      "(Réserve obligatoire + fonds d'établissement + réserves libres + droits d'adhésion + dons et legs) / max(18 % des cotisations acquises n-1 ; 26 % des prestations maladie moyennes des 3 derniers exercices)",
      "111 + 1188 + 106 / 70 (n-1) et 601, 608", "Au moins 1 (marge disponible au moins égale à la marge exigée)", ("min", 1.0),
      DEC + ", § 1.1.1 (fonds d'établissement) ; " + RE + ", art. 23", {}, unite="x")
ratio("ROB", "Ratio de réserve obligatoire", "Actuariel", 3,
      "Réserves obligatoires x 12 / Charges de prestations", "111 x 12 / 60 + 6901 - 7901",
      "Entre 6 et 9 mois (tableau 1) ; 9 mois au § 1.2.2.5.4", ("entre", 6.0, 9.0),
      DEC + ", tableau 1, ligne 10 et § 1.2.2.5.4",
      serie(lambda a, ex: r_(a["reserves_legales_111"] * (12 if ex < 2026 else 8), a["charges_prestations"])),
      serie(lambda a, ex: a["reserves_legales_111"]), serie(lambda a, ex: a["charges_prestations"]), unite="mois")
ratio("FEP", "Ratio des fonds propres et assimilés", "Actuariel", 4,
      "Fonds propres et assimilés / Dettes (y compris provisions techniques)",
      "10 + 11 + 12 + 13 + 14 + 190 + 15 / classe 4 créditrice + 16 + 190 + 15 + 5 créditeur", "Au moins 1", ("min", 1.0),
      DEC + ", tableau 2, ligne 9 et § 1.2.2.5.3",
      serie(lambda a, ex: r_(a["fonds_propres"] + a["provisions_190"] + a["provisions_reglementees_15"],
                             a["dettes_ct_pcms"] + a["emprunts_16"] + a["provisions_190"] + a["provisions_reglementees_15"]
                             + a["provision_pensions_196"])),
      serie(lambda a, ex: a["fonds_propres"] + a["provisions_190"] + a["provisions_reglementees_15"]),
      serie(lambda a, ex: a["dettes_ct_pcms"] + a["emprunts_16"] + a["provisions_190"] + a["provisions_reglementees_15"]
            + a["provision_pensions_196"]), unite="x")
ratio("LIQ", "Ratio de liquidité", "Actuariel", 5,
      "Trésorerie actif / (Dettes à court terme + variation des provisions techniques)",
      "classe 5 débitrice / classe 4 créditrice + classe 5 créditrice + variation de 190 et 158", "Au moins 1", ("min", 1.0),
      DEC + ", tableau 2, ligne 8 et § 1.2.2.5.2",
      serie(lambda a, ex: r_(a["tresorerie_actif"], a["dettes_ct_pcms"] + a["var_provisions_techniques"])),
      serie(lambda a, ex: a["tresorerie_actif"]), serie(lambda a, ex: a["dettes_ct_pcms"] + a["var_provisions_techniques"]),
      unite="x")
ratio("EQS", "Ratio d'équilibre stable", "Actuariel", 6,
      "Emplois moyens et longs / Ressources moyennes et longues", "classe 2 nette / classes 10 à 19",
      "Au plus 1", ("max", 1.0), DEC + ", tableau 1, ligne 11 et § 1.2.3.1",
      serie(lambda a, ex: r_(a["actif_immobilise_net"], a["fonds_propres"] + a["provisions_techniques"]
                             + a["provision_pensions_196"] + a["emprunts_16"])),
      serie(lambda a, ex: a["actif_immobilise_net"]),
      serie(lambda a, ex: a["fonds_propres"] + a["provisions_techniques"] + a["provision_pensions_196"] + a["emprunts_16"]),
      unite="x")
ratio("AUT", "Ratio d'autonomie financière", "Actuariel", 7,
      "Cotisations acquises / Total des charges", "70 + 7906 - 6906 / classe 6 + 81 + 83 + 85", "Au moins 1", ("min", 1.0),
      DEC + ", tableau 1, ligne 9 et § 1.2.2.5.1",
      serie(lambda a, ex: r_(a["cotisations_acquises"], a["total_charges"])),
      serie(lambda a, ex: a["cotisations_acquises"]), serie(lambda a, ex: a["total_charges"]), unite="x")
ratio("EXN", "Taux d'excédent net", "Actuariel", 8,
      "Excédent net de l'exercice / Cotisations acquises", "résultat net / 70 + 7906 - 6906", "Au plus 5 %", ("max", 0.05),
      DEC + ", tableau 2, ligne 10 et § 1.2.2.5.5",
      serie(lambda a, ex: r_(a["resultat_net"], a["cotisations_acquises"])),
      serie(lambda a, ex: a["resultat_net"]), serie(lambda a, ex: a["cotisations_acquises"]))
ratio("TFP", "Taux de financement propre", "Actuariel", 9,
      "Total des produits avant subventions / Total des charges", "produits (liste de l'annexe) - 75 et 88 / charges",
      "Au moins 100 %", ("min", 1.0), DEC + ", tableau 2, ligne 11 et § 1.2.3.2",
      serie(lambda a, ex: r_(a["total_produits"] - a["subventions"], a["total_charges"])),
      serie(lambda a, ex: a["total_produits"] - a["subventions"]), serie(lambda a, ex: a["total_charges"]))
# --- Risques longs (section II de l'annexe ; règlement d'exécution art. 19 et 21) ---
ratio("RL1", "Ratio de solvabilité de la branche épargne (risques longs)", "Actuariel - risques longs", 10,
      "Actifs admis / Engagements réglementés", "banques et DAT dédiés, FCP / 41195100 + 41195200", "Au moins 1",
      ("min", 1.0), DEC + ", section II, § 2.1",
      serie(lambda a, ex: r_(a["actifs_dedies_epargne"], a["epargne"])),
      serie(lambda a, ex: a["actifs_dedies_epargne"]), serie(lambda a, ex: a["epargne"]), unite="x")
ratio("RL2", "Ratio de trésorerie à échéance (risques longs)", "Actuariel - risques longs", 11,
      "(Valeurs réalisables + disponibles) / Dettes à court terme", "créances court terme + classe 5 débitrice / dettes court terme",
      "Supérieur à 1", ("min_strict", 1.0), DEC + ", section II, § 2.2",
      serie(lambda a, ex: r_(a["creances_ct"] + a["tresorerie_actif"], a["dettes_ct_pcms"])),
      serie(lambda a, ex: a["creances_ct"] + a["tresorerie_actif"]), serie(lambda a, ex: a["dettes_ct_pcms"]), unite="x")
ratio("RL3", "Ratio de trésorerie immédiate (risques longs)", "Actuariel - risques longs", 12,
      "Valeurs disponibles / Dettes à court terme", "classe 5 débitrice / dettes court terme", "Au moins 1", ("min", 1.0),
      DEC + ", section II, § 2.3",
      serie(lambda a, ex: r_(a["tresorerie_actif"], a["dettes_ct_pcms"])),
      serie(lambda a, ex: a["tresorerie_actif"]), serie(lambda a, ex: a["dettes_ct_pcms"]), unite="x")
ratio("RL4", "Réserve de trésorerie des risques longs (un trimestre d'engagements et de frais)", "Actuariel - risques longs", 13,
      "Réserve de trésorerie constituée / (sorties d'épargne + frais de fonctionnement au prorata) / 4",
      "aucun compte de réserve de trésorerie identifié", "Au moins 1 trimestre", ("min", 1.0),
      RE + ", art. 19, alinéa 3", serie(lambda a, ex: 0.0),
      serie(lambda a, ex: 0.0), serie(lambda a, ex: (a["sorties_epargne"] + a["charges_fonctionnement"]
                                                   * a["epargne"] / (a["epargne"] + a["cotisations_acquises"])) / 4),
      unite="x")
ratio("RL5", "Fonds propres / prestations des risques longs des deux derniers exercices", "Actuariel - risques longs", 14,
      "Fonds propres / sorties d'épargne des exercices n et n-1", "10 à 14 + résultat / débits de 41195100 et 41195200",
      "Au moins 1", ("min", 1.0), RE + ", art. 21",
      {ex: (r_(A[ex]["fonds_propres"], A[ex]["sorties_epargne"] + A[ex - 1]["sorties_epargne"]) if ex > 2020 else None) for ex in EX},
      serie(lambda a, ex: a["fonds_propres"]),
      {ex: (A[ex]["sorties_epargne"] + A[ex - 1]["sorties_epargne"]) if ex > 2020 else None for ex in EX}, unite="x")
ratio("RL6", "Taux brut de rentabilité des placements", "Actuariel - risques longs", 15,
      "Intérêts perçus / Capital investi", "77 / DAT et FCP moyens", "Au moins le taux moyen du marché", None,
      DEC + ", section II, § 2.4", {}, qualitatif="Comparé au taux directeur de la BCEAO (partie 9) : appréciation, pas de seuil chiffré.")
ratio("RL7", "Solde moyen des comptes courants non rémunérés", "Actuariel - risques longs", 16,
      "Somme des soldes des comptes courants / nombre de comptes", "521 (soldes aux clôtures)", "Avoisinant zéro", None,
      DEC + ", section II, § 2.5", serie(lambda a, ex: a["comptes_courants_521"]), unite="FCFA",
      qualitatif="Norme non chiffrée : appréciation.")
# --- Gestion ---
ratio("CHF", "Taux des charges de fonctionnement", "Gestion", 17, "Charges de fonctionnement / Cotisations acquises",
      "612, 613, 618, 62 à 68 / 70 + 7906 - 6906", "Au plus 20 %", ("max", 0.20), DEC + ", tableau 1, ligne 2 et § 1.2.1.2.1",
      serie(lambda a, ex: r_(a["charges_fonctionnement"], a["cotisations_acquises"])),
      serie(lambda a, ex: a["charges_fonctionnement"]), serie(lambda a, ex: a["cotisations_acquises"]))
ratio("CHB", "Taux des charges de fonctionnement brut", "Gestion", 18, "Charges de fonctionnement brut / Total des produits",
      "61 à 68, 691, 697 / produits (liste de l'annexe)", "Au plus 15 %", ("max", 0.15), DEC + ", tableau 1, ligne 3 et § 1.2.1.2.2",
      serie(lambda a, ex: r_(a["charges_fonctionnement_brut"], a["total_produits"])),
      serie(lambda a, ex: a["charges_fonctionnement_brut"]), serie(lambda a, ex: a["total_produits"]))
ratio("CUP", "Coût par unité monétaire prestée", "Gestion", 19, "Charges de fonctionnement / Charges de prestations",
      "612, 613, 618, 62 à 68 / 60 + 6901 - 7901", "Au plus 15 %", ("max", 0.15), DEC + ", tableau 2, ligne 5 et § 1.2.1.2.3",
      serie(lambda a, ex: r_(a["charges_fonctionnement"], a["charges_prestations"])),
      serie(lambda a, ex: a["charges_fonctionnement"]), serie(lambda a, ex: a["charges_prestations"]))
ratio("CHP", "Taux de charges de personnel", "Gestion", 20, "Charges de personnel / Charges de fonctionnement brut",
      "66 / 61 à 68, 691, 697", "Au plus 85 %", ("max", 0.85), DEC + ", tableau 1, ligne 4 et § 1.2.1.2.4",
      serie(lambda a, ex: r_(a["charges_personnel"], a["charges_fonctionnement_brut"])),
      serie(lambda a, ex: a["charges_personnel"]), serie(lambda a, ex: a["charges_fonctionnement_brut"]))
ratio("REC", "Taux de recouvrement des cotisations", "Gestion", 21, "Cotisations recouvrées / Cotisations émises",
      "fichiers retour / fichiers aller", "Au moins 90 %", ("min", 0.90), DEC + ", tableau 2, ligne 13 et § 1.2.2.3.1", {},
      qualitatif="Fichiers aller des postes 936 et 8Z6 non reçus. Calculable seulement pour les retraités CGRAE.")
ratio("FID", "Taux de fidélisation", "Gestion", 22, "Renouvellements d'adhésion n / Adhérents n-1",
      "matricules du poste 936 présents en n et n-1", "Au moins 80 %", ("min", 0.80), DEC + ", tableau 2, ligne 4 et § 1.2.1.1.4",
      {2025: EFF["fidelisation_2025"], 2026: EFF["fidelisation_2026"]})
ratio("DPP", "Délai moyen de paiement des prestataires de soins", "Gestion", 23,
      "Dettes envers les prestataires x 365 / Prestations prises en charge", "40110000 x 365 / 60", "Délai contractuel",
      None, DEC + ", tableau 2, ligne 7 et § 1.2.2.3.2",
      serie(lambda a, ex: r_(a["dettes_fournisseurs_cliniques"] * (365 if ex < 2026 else 243), a["charges_prestations"])),
      unite="jours", qualitatif="Délai contractuel des conventions non transmis.")
ratio("FOR", "Taux d'investissement brut en formation", "Gestion", 24, "Coût total de formation / Total des produits",
      "6391 / produits", "Au plus 3 %", ("max", 0.03), DEC + ", tableau 2, ligne 6 et § 1.2.2.1.1.2",
      serie(lambda a, ex: r_(a["cout_formation"], a["total_produits"])))
ratio("CRB", "Taux de croissance brute des bénéficiaires", "Gestion", 25, "(Bénéficiaires n - n-1) / Bénéficiaires n-1",
      "adhérents du poste 936 (ayants droit non transmis)", "Objectif du plan de développement stratégique", None,
      DEC + ", tableau 2, ligne 1", {2025: EFF["croissance_2025"]}, qualitatif="Plan de développement stratégique non transmis.")
ratio("RBA", "Ratio bénéficiaire-adhérent", "Gestion", 26, "Bénéficiaires / Adhérents", "fichier des ayants droit",
      "Objectif du plan de développement stratégique", None, DEC + ", tableau 2, ligne 2", {},
      qualitatif="Fichier des ayants droit non transmis.")
ratio("PEN", "Taux de pénétration", "Gestion", 27, "Bénéficiaires / Population cible", "effectif de la Police nationale",
      "Objectif du plan de développement stratégique", None, DEC + ", tableau 2, ligne 3", {},
      qualitatif="Effectif total de la Police nationale non transmis.")
ratio("IND", "Taux de prestations indues", "Gestion", 28, "Prestations indûment payées / Prestations payées",
      "rapports de contrôle a posteriori", "Au plus 2 %", ("max", 0.02), DEC + ", tableau 2, ligne 12 et § 1.2.2.3.3", {},
      qualitatif="Rapports de contrôle a posteriori non transmis.")
ratio("CAP", "Ratio de contrôle a posteriori", "Gestion", 29, "Factures contrôlées / Factures traitées",
      "registre des contrôles", "Au moins 5 % (10 % pour les risques longs)", ("min", 0.05),
      DEC + ", § 1.2.2.4 et section II, § 2.9", {}, qualitatif="Registre des contrôles non transmis.")
ratio("DLQ", "Délai de liquidation des premières demandes (risques longs)", "Gestion", 30, "Durée maximale de liquidation",
      "dossiers avec dates", "Au plus 30 jours", ("max", 30), DEC + ", section II, § 2.7", {}, unite="jours",
      qualitatif="Dossiers de prestations avec dates non transmis.")
ratio("RAT", "Restes à traiter en fin de période (risques longs)", "Gestion", 31, "Dossiers en instance / Dossiers reçus",
      "dossiers avec statut", "Au plus 5 %", ("max", 0.05), DEC + ", section II, § 2.8", {},
      qualitatif="Dossiers de prestations avec statut non transmis.")
ratio("DRB", "Délai de remise en banque des titres de paiement (risques longs)", "Gestion", 32, "Délai de dépôt",
      "registre des encaissements", "Au plus 1 jour", ("max", 1), DEC + ", section II, § 2.6", {}, unite="jours",
      qualitatif="Registre des encaissements non transmis.")
# --- Gouvernance ---
ratio("FNG", "Taux de règlement de la cotisation au Fonds national de garantie", "Gouvernance", 33,
      "Montant payé à l'échéance / Cotisation due", "44100000 (aucun compte au FPPN)", "100 %", ("min", 1.0),
      DEC + ", tableau 1, ligne 1 ; Règlement n° 07/2009, art. 63", {},
      qualitatif="Aucun compte FNG (FOMUS-CI) dans les balances : adhésion et cotisation à confirmer.")
ratio("TTR", "Taux de tenue des réunions des organes", "Gouvernance", 34, "Réunions tenues / Réunions statutaires",
      "procès-verbaux", "100 %", ("min", 1.0), DEC + ", tableau 1, ligne 6", {}, qualitatif="Procès-verbaux non transmis.")
ratio("OSA", "Existence et tenue des outils de suivi administratif", "Gouvernance", 35, "Contrôle sur place et sur pièce",
      "registres", "Existence et tenue", None, DEC + ", tableau 1, ligne 7", {},
      qualitatif="Contrôle sur place, hors du champ des données reçues.")
ratio("OSF", "Existence et tenue des outils de suivi comptable et financier", "Gouvernance", 36, "Contrôle sur place et sur pièce",
      "registres", "Existence et tenue", None, DEC + ", tableau 1, ligne 8", {},
      qualitatif="Contrôle sur place, hors du champ des données reçues.")

# Marge de solvabilité (calcul détaillé)
MS = []
for ex in EX:
    a = A[ex]
    exig_cot = 0.18 * A[ex - 1]["cotisations_acquises"] if ex > 2020 else None
    trois = [A[e]["prest_maladie"] for e in (ex, ex - 1, ex - 2) if e in A]
    exig_prest = 0.26 * np.mean(trois) if len(trois) == 3 else None
    exigee = max(x for x in (exig_cot, exig_prest) if x is not None) if (exig_cot or exig_prest) else None
    dispo_fppn = a["reserves_legales_111"] + a["reserves_facultatives_118"] + a["reserve_disponible_106"]
    dispo_nette = a["fonds_propres"]
    MS.append({"exercice": ex, "exigee_18pc_cotisations_n_1": exig_cot, "exigee_26pc_prest_maladie_3ans": exig_prest,
               "marge_exigee": exigee, "reserve_obligatoire_111": a["reserves_legales_111"],
               "reserves_libres_1188": a["reserves_facultatives_118"], "reserve_disponible_106": a["reserve_disponible_106"],
               "fonds_etablissement": 0.0, "marge_disponible_comme_fppn": dispo_fppn,
               "pertes_reportees_129": a["ran_debiteur_129"], "marge_disponible_nette": dispo_nette,
               "ratio_comme_fppn": r_(dispo_fppn, exigee), "ratio_net": r_(dispo_nette, exigee),
               "fonds_etablissement_minimum": exigee / 3 if exigee else None})
for x in RATIOS:
    if x["code"] == "MSE":
        x["valeurs"] = {m["exercice"]: m["ratio_comme_fppn"] for m in MS}
        x["valeurs_retraitees"] = {m["exercice"]: m["ratio_net"] for m in MS}
        x["num"] = {m["exercice"]: m["marge_disponible_comme_fppn"] for m in MS}
        x["den"] = {m["exercice"]: m["marge_exigee"] for m in MS}

# Valeurs retraitées : épargne des adhérents classée en ressource longue (risque long, art. 3 du règlement
# d'exécution) et avances au CMP, au Grenier et à l'Hôpital de la Police classées en emplois longs.
RETR = {
    "LIQ": {ex: r_(A[ex]["tresorerie_actif"], A[ex]["dettes_ct_hors_epargne"] + A[ex]["var_provisions_techniques"]) for ex in EX},
    "RL2": {ex: r_(A[ex]["creances_ct"] + A[ex]["tresorerie_actif"], A[ex]["dettes_ct_hors_epargne"]) for ex in EX},
    "RL3": {ex: r_(A[ex]["tresorerie_actif"], A[ex]["dettes_ct_hors_epargne"]) for ex in EX},
    "EQS": {ex: r_(A[ex]["actif_immobilise_net"] + A[ex]["emplois_longs_tiers"],
                   A[ex]["fonds_propres"] + A[ex]["provisions_techniques"] + A[ex]["provision_pensions_196"]
                   + A[ex]["emprunts_16"] + A[ex]["epargne"]) for ex in EX},
}
for x in RATIOS:
    if x["code"] in RETR:
        x["valeurs_retraitees"] = RETR[x["code"]]

# Rentabilité des placements (reprise de la v2)
x = next(x for x in RATIOS if x["code"] == "RL6")
x["valeurs"] = {int(r["exercice"]): r.get("rendement") for r in R.get("rendement_placements", []) if r.get("rendement") is not None}
# Recouvrement CGRAE
x = next(x for x in RATIOS if x["code"] == "REC")
x["valeurs"] = {"2024-2026 (CGRAE)": EFF["recouvrement_cgrae"]}


# ---------------------------------------------------------------------------
# 4. Statut : règle écrite
# ---------------------------------------------------------------------------
def statut(test, v, ex):
    """Conforme si la valeur respecte la norme du texte ; Non conforme sinon ; En attente si donnée absente ;
    À apprécier si la norme n'est pas chiffrée. 2020-2021 : indicatif (Décision du 1er juin 2022) ;
    2026 : indicatif (8 mois, balance provisoire)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "En attente"
    if test is None:
        return "À apprécier"
    t = test[0]
    if isinstance(v, float) and v < 0 and t == "max":
        return "Non conforme (ressources négatives)" + (" (indicatif)" if isinstance(ex, int) and (ex < 2022 or ex == 2026) else "")
    ok = {"entre": lambda: test[1] <= v <= test[2], "min": lambda: v >= test[1], "min_strict": lambda: v > test[1],
          "max": lambda: v <= test[1]}[t]()
    if t == "entre" and not ok and v > test[2]:
        s = "Au-dessus de la fourchette"
    else:
        s = "Conforme" if ok else "Non conforme"
    if isinstance(ex, int) and (ex < 2022 or ex == 2026):
        s += " (indicatif)"
    return s


for x in RATIOS:
    x["statuts"] = {ex: statut(x["test"], v, ex) for ex, v in x["valeurs"].items()}
    if "valeurs_retraitees" in x:
        x["statuts_retraites"] = {ex: statut(x["test"], v, ex) for ex, v in x["valeurs_retraitees"].items()}
        nc_r = [e for e in (2022, 2023, 2024, 2025) if x["statuts_retraites"].get(e, "").startswith(("Non", "Au-dessus"))]
        x["statut_retraite"] = "Non conforme" if nc_r else "Conforme"
    if not x["valeurs"]:
        x["statut_global"] = "En attente" if x["test"] is not None or x["qualitatif"] else "À apprécier"
    else:
        jug = [x["statuts"][e] for e in (2022, 2023, 2024, 2025) if e in x["statuts"]]
        if not jug:
            jug = list(x["statuts"].values())
        nc = [e for e in (2022, 2023, 2024, 2025) if e in x["statuts"] and x["statuts"][e].startswith(("Non", "Au-dessus"))]
        consec = any(e in nc and e + 1 in nc for e in nc)
        if x["test"] is None:
            x["statut_global"] = "À apprécier"
        elif not nc and all(j.startswith("Conforme") for j in jug):
            x["statut_global"] = "Conforme"
        elif nc:
            x["statut_global"] = "Non conforme"
        else:
            x["statut_global"] = "En attente"
        x["deux_exercices_consecutifs"] = consec

# ---------------------------------------------------------------------------
# 5. Affectation des excédents et excédent par risque
# ---------------------------------------------------------------------------
AFF = []
for ex in (2023, 2024, 2025):
    res_prec = A[ex - 1]["resultat_net"]
    dot_111 = A[ex]["reserves_legales_111"] - A[ex - 1]["reserves_legales_111"]
    dot_118 = A[ex]["reserves_facultatives_118"] - A[ex - 1]["reserves_facultatives_118"]
    apur = A[ex]["ran_debiteur_129"] - A[ex - 1]["ran_debiteur_129"]
    AFF.append({"exercice_affectation": ex, "resultat_n_1": res_prec, "ran_crediteur_n_1": A[ex - 1]["ran_crediteur_121"],
                "dotation_reserve_legale": dot_111, "dotation_reserves_facultatives": dot_118,
                "apurement_pertes_129": apur, "taux_dotation_legale": r_(dot_111, res_prec)})
EXR = []
for ex in EX:
    a = A[ex]
    cot = a["cot_maladie"] + a["cot_deces"]
    pm, pd_ = a["prest_maladie"], a["prest_deces"] + a["prime_assurance_6518"]
    frais = a["charges_fonctionnement"] - a["prime_assurance_6518"]  # la prime décès est rattachée au risque décès
    fm, fd = (frais * a["cot_maladie"] / cot, frais * a["cot_deces"] / cot) if cot else (None, None)
    autres = a["resultat_net"] - (a["cot_maladie"] - pm - fm) - (a["cot_deces"] - pd_ - fd)
    em, ed = a["cot_maladie"] - pm - fm, a["cot_deces"] - pd_ - fd
    EXR.append({"exercice": ex, "cot_maladie": a["cot_maladie"], "prest_maladie": pm, "frais_maladie": fm,
                "excedent_maladie": em, "cot_deces": a["cot_deces"], "prest_deces_et_prime": pd_, "frais_deces": fd,
                "excedent_deces": ed, "autres_elements_non_affectes": autres, "resultat_net": a["resultat_net"],
                "part_maladie_cotisations": a["cot_maladie"] / cot if cot else None,
                "dotation_minimale": 0.30 * max(em, 0) + 0.50 * max(ed, 0)})
for af in AFF:
    e = next(x for x in EXR if x["exercice"] == af["exercice_affectation"] - 1)
    af["dotation_minimale_par_risque"] = e["dotation_minimale"]
    af["ecart_dotation"] = af["dotation_reserve_legale"] - e["dotation_minimale"]

# ---------------------------------------------------------------------------
# 6. Classement des comptes (tiroirs) et contrôle PCMS limité aux comptes des calculs
# ---------------------------------------------------------------------------
PCMS_RUB = {"10": "Fonds mutualistes", "106": "Écarts de réévaluation des immeubles", "11": "Réserves", "111": "Réserves légales",
            "118": "Autres réserves", "12": "Report à nouveau", "129": "Report à nouveau débiteur", "121": "Report à nouveau créditeur",
            "15": "Provisions réglementées (caractère fiscal)", "158": "Autres provisions réglementées", "16": "Emprunts",
            "190": "Provisions techniques", "19011": "Provision pour prestations à payer - maladie",
            "19012": "Provision pour prestations à payer - obsèques", "196": "Provisions pour pensions",
            "2": "Actif immobilisé", "27": "Autres immobilisations financières (prêts)", "40": "Fournisseurs",
            "41195": "Autres comptes créditeurs des adhérents", "471": "Comptes d'attente débiteurs", "485": "Créances sur cessions",
            "52": "Banques", "53": "Établissements financiers", "57": "Caisse", "60": "Charges de prestations",
            "601": "Prestations maladie", "602": "Prestations obsèques", "608": "Autres prestations", "61": "Achats",
            "641": "Impôts et taxes sur salaires", "651": "Primes d'assurances", "66": "Charges de personnel",
            "70": "Cotisations", "701": "Cotisations maladie", "702": "Cotisations décès", "704": "Autres produits techniques", "77": "Revenus financiers"}


def rubrique(c):
    for k in sorted(PCMS_RUB, key=len, reverse=True):
        if c.startswith(k):
            return k + " " + PCMS_RUB[k]
    return c[:2]


def tiroir(c):
    if c.startswith(("10", "11", "12", "13", "14")):
        return ("Ressources propres", "Ressource longue", "Ressource longue")
    if c.startswith(("190", "15801", "15802")):
        return ("Provisions techniques", "Ressource longue", "Engagement à moins d'un an (prestations à payer)")
    if c.startswith(("15", "19")):
        return ("Autres provisions", "Ressource longue", "Ressource longue")
    if c.startswith("16"):
        return ("Emprunts", "Ressource longue", "Ressource longue")
    if c.startswith("41195"):
        return ("Épargne des adhérents", "Dette à court terme (classe 4)", "Engagement long (risque long, échéance à établir)")
    if c.startswith(("47111100", "47111101", "47111102", "47111104", "47111105", "47111201", "485")):
        return ("Avances et créances de long terme", "Créance à court terme (classe 4)", "Emploi long ou incertain")
    if c.startswith("2"):
        return ("Actif immobilisé", "Emploi long", "Emploi long")
    if c.startswith(("3", "4")):
        return ("Tiers à court terme", "Court terme", "Court terme")
    if c.startswith("5"):
        return ("Trésorerie", "Trésorerie", "Trésorerie")
    if c.startswith("6"):
        return ("Charges", "Charges", "Charges")
    if c.startswith("7"):
        return ("Produits", "Produits", "Produits")
    return ("Hors activités ordinaires", "HAO", "HAO")


USAGE = [(r"^70", "SIN, AUT, EXN, CHF, MSE"), (r"^60", "SIN, ROB, CUP, DPP"), (r"^(612|613|618|62|63|64|65|66|67|68)", "CHF, CUP"),
         (r"^66", "CHP"), (r"^111", "ROB, MSE"), (r"^(1188|106)", "MSE"), (r"^(190|158)", "FEP, LIQ, EQS"),
         (r"^41195", "LIQ, RL1, RL2, RL3, RL4, RL5"), (r"^5", "LIQ, RL1, RL2, RL3"), (r"^2", "EQS"), (r"^4711110", "EQS (retraité)")]
cl = []
for c, lib in LIB.items():
    if not bal[(bal.compte == c) & (bal.net.abs() > 0.5)].shape[0]:
        continue
    t = tiroir(c)
    us = ", ".join(u for m, u in USAGE if pd.Series([c]).str.match(m)[0])
    cl.append({"compte": c, "intitule_fppn": lib, "rubrique_pcms": rubrique(c), "tiroir": t[0],
               "classement_pcms": t[1], "classement_retraite": t[2], "ratios": us,
               **{f"solde_{ex}": round(A_ex) for ex, A_ex in ((e, solde(e, "^" + c + "$")) for e in EX)}})
CL = pd.DataFrame(cl)

PCMS_CTRL = [
    {"point": "Provisions pour prestations à payer", "comptes_fppn": "19011000, 19012000 (2020 à 2023) puis 15801000, 15802000 (2024 à 2026)",
     "attendu_pcms": "19011 et 19012 (classe 19, provisions techniques). La rubrique 15 est réservée aux provisions réglementées à caractère fiscal.",
     "constat": "Conforme de 2020 à 2023. Depuis 2024, les provisions techniques sont classées en provisions réglementées (158), sans changement de montant.",
     "incidence": "Aucune sur les montants. Les provisions techniques n'apparaissent plus comme telles dans le bilan, ce qui gêne leur lecture et le contrôle de l'article 7 du règlement d'exécution.",
     "source": "PCMS, liste des comptes (classe 1) et fiche n° 6, rubrique 15"},
    {"point": "Réserve disponible", "comptes_fppn": "10620000 « Réserve disponible » : " + f"{A[2025]['reserve_disponible_106']:,.0f}".replace(",", " ") + " FCFA, inchangé de 2020 à 2026",
     "attendu_pcms": "106 = écarts de réévaluation des immeubles (fonds mutualistes) ; les réserves vont en 11.",
     "constat": "Libellé et rubrique ne concordent pas. L'origine du montant (réévaluation ou excédents mis en réserve) n'est pas documentée.",
     "incidence": "Le montant compte ou non dans la marge de solvabilité selon sa nature : à confirmer par le FPPN.",
     "source": "PCMS, liste des comptes et fiche n° 1, rubrique 10"},
    {"point": "Réserves légales et facultatives", "comptes_fppn": "11100000, 11880000", "attendu_pcms": "111 réserves légales, 118 autres réserves",
     "constat": "Conforme.", "incidence": "Base du ratio de réserve obligatoire et de la marge de solvabilité.",
     "source": "PCMS, fiche n° 2, rubrique 11"},
    {"point": "Épargne des adhérents (PEP, EFC)", "comptes_fppn": "41195100, 41195200",
     "attendu_pcms": "41195 « Autres comptes créditeurs des adhérents » (classe 4, tiers).",
     "constat": "Compte conforme au PCMS. Le PCMS ne prévoit pas de compte propre aux risques longs ; le règlement d'exécution exige des états financiers par risque long (art. 13).",
     "incidence": "Classée en dette à court terme, l'épargne pèse sur les ratios de liquidité. Son échéance réelle reste à établir avec les âges.",
     "source": "PCMS, fiche n° 25, rubrique 41 ; règlement d'exécution n° 004/2022, art. 13"},
    {"point": "Avances au CMP, au Grenier et créance sur l'Hôpital de la Police", "comptes_fppn": "47111100, 47111101, 47111102, 47111104, 47111105",
     "attendu_pcms": "471 = comptes d'attente débiteurs (opérations à régulariser). Les prêts et placements durables relèvent des classes 26 et 27.",
     "constat": "Des emplois durables de " + f"{A[2026]['emplois_longs_tiers']:,.0f}".replace(",", " ") + " FCFA au 31/08/2026 sont suivis en comptes d'attente ou de tiers.",
     "incidence": "Ils sont comptés en court terme dans les ratios de liquidité et absents des emplois longs du ratio d'équilibre stable.",
     "source": "PCMS, liste des comptes (classes 2 et 4)"},
    {"point": "Cotisations et prestations", "comptes_fppn": "70100000, 70210000 ; 601, 602, 608",
     "attendu_pcms": "701 maladie, 702 décès ; 601 maladie, 602 obsèques, 608 autres prestations", "constat": "Conforme.",
     "incidence": "Base de la sinistralité réglementaire.", "source": "PCMS, liste des comptes (classes 6 et 7)"},
    {"point": "Primes, indemnités et supplément familial du personnel", "comptes_fppn": "64130000, 64140000, 64150000",
     "attendu_pcms": "641 = impôts et taxes sur salaires ; les rémunérations relèvent de 66.",
     "constat": "Ces montants sont classés en impôts et taxes.", "incidence": "Faible (moins de 8 000 000 FCFA par an) : le taux de charges de personnel est légèrement sous-estimé.",
     "source": "PCMS, liste des comptes (classe 6)"},
]

# ---------------------------------------------------------------------------
# 7. Actif par délai de mobilisation (étapes 1 à 3 de la méthode d'adossement)
# ---------------------------------------------------------------------------
MOB = []
for ex in (2025, 2026):
    a = A[ex]
    MOB.append({"cloture": "31/12/2025" if ex == 2025 else "31/08/2026",
                "immediat": a["disponibilites"],
                "moins_1_an": a["placements_tresorerie"] + debiteurs(ex, r"^(3|4)") - a["emplois_longs_tiers"],
                "1_a_5_ans": a["prets_pep"] + a["grenier"],
                "plus_5_ans_ou_incertain": a["actif_immobilise_net"] - a["prets_pep"] + a["cmp"] + a["creance_hopital"]
                + solde(ex, r"^(47111201|485)"),
                "passif_moins_1_an": a["dettes_ct_hors_epargne"] + a["provisions_techniques"],
                "epargne_echeance_a_etablir": a["epargne"]})
for m in MOB:
    m["total_actif"] = m["immediat"] + m["moins_1_an"] + m["1_a_5_ans"] + m["plus_5_ans_ou_incertain"]

# ---------------------------------------------------------------------------
# 8. Sorties
# ---------------------------------------------------------------------------
OUT = {"agregats": {str(k): v for k, v in A.items()}, "effectifs": EFF, "ratios": RATIOS, "marge_solvabilite": MS,
       "affectation": AFF, "excedent_par_risque": EXR, "pcms": PCMS_CTRL, "mobilisation": MOB,
       "classement_resume": CL.groupby(["tiroir", "classement_pcms", "classement_retraite"]).size().reset_index(name="comptes").to_dict("records")}


def conv(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return float(o)
    if isinstance(o, tuple):
        return list(o)
    return str(o)


json.dump(OUT, open(os.path.join(SORTIE, "ratios.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=conv)

with pd.ExcelWriter(os.path.join(SORTIE, "FPPN_Calculs_ratios.xlsx")) as w:
    rows = []
    for x in RATIOS:
        base = {"Priorité": x["priorite"], "Code": x["code"], "Famille": x["famille"], "Ratio": x["nom"], "Formule": x["formule"],
                "Comptes ou données": x["comptes"], "Norme": x["norme"], "Source": x["source"]}
        for ex in EX:
            v = x["valeurs"].get(ex)
            base[str(ex)] = v
            base[f"Statut {ex}"] = x["statuts"].get(ex, "En attente")
        base["Statut global (2022 à 2025)"] = x["statut_global"]
        if "valeurs_retraitees" in x:
            for ex in EX:
                base[f"Retraité {ex}"] = x["valeurs_retraitees"].get(ex)
            base["Statut retraité"] = x["statut_retraite"]
        base["Non conforme deux exercices consécutifs"] = "Oui" if x.get("deux_exercices_consecutifs") else "Non"
        base["Observation"] = x["qualitatif"] or ""
        rows.append(base)
    pd.DataFrame(rows).to_excel(w, sheet_name="10_ratios", index=False)
    det = []
    for x in RATIOS:
        for ex in EX:
            if ex in x["num"] or ex in x["den"]:
                det.append({"Code": x["code"], "Exercice": ex, "Numérateur": x["num"].get(ex), "Dénominateur": x["den"].get(ex),
                            "Valeur": x["valeurs"].get(ex), "Statut": x["statuts"].get(ex)})
    pd.DataFrame(det).to_excel(w, sheet_name="10_ratios_detail", index=False)
    CL.to_excel(w, sheet_name="10_classement_comptes", index=False)
    pd.DataFrame(MS).to_excel(w, sheet_name="10_marge_solvabilite", index=False)
    pd.DataFrame(AFF).to_excel(w, sheet_name="10_affectation_excedents", index=False)
    pd.DataFrame(EXR).to_excel(w, sheet_name="10_excedent_par_risque", index=False)
    pd.DataFrame(PCMS_CTRL).to_excel(w, sheet_name="10_controle_pcms", index=False)
    pd.DataFrame(MOB).to_excel(w, sheet_name="10_actif_mobilisation", index=False)
    pd.DataFrame([{"exercice": k, **v} for k, v in A.items()]).to_excel(w, sheet_name="10_agregats", index=False)
print(len(RATIOS), "ratios ;", len(CL), "comptes classés")
for x in RATIOS:
    print(f"{x['priorite']:>2} {x['code']} {x['statut_global']:<14}", {k: (round(v, 4) if isinstance(v, float) else v) for k, v in x["valeurs"].items()})
