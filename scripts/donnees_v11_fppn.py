# -*- coding: utf-8 -*-
"""
FPPN - Données de la version 11 du rapport (vérification de cohérence du 23/10/2026)
====================================================================================
1. Placements lus dans les grands livres : encours moyen pondéré par les jours, pic, soldes de fin de mois du DAT EFC,
   rendement 2021 à 2025 recalculé sur l'encours réellement placé (remplace la moyenne des soldes de clôture).
2. Provision pour départ à la retraite avec le plafond de la convention collective du 19/07/1977 (art. 40) :
   25 fois le SMIG annuel (base de 2 080 heures), SMIG de 75 000 FCFA par mois depuis le 01/01/2023 (décret n° 2022-986 du 21/12/2022).
3. Ratios par garantie (maladie, décès) : sinistralité, taux de frais affecté, ratio combiné, résultat technique.
Utilisation : python donnees_v11_fppn.py <placements_gl.json> <v9.json> <ratios.json> <resultats.json> <sortie.json>
"""
import json
import sys

PGL, V9J, RATJ, RESJ, SORTIE = sys.argv[1:6]
G = {int(k): v for k, v in json.load(open(PGL, encoding="utf-8")).items()}
V9 = json.load(open(V9J, encoding="utf-8"))
Q = json.load(open(RATJ, encoding="utf-8"))
R = json.load(open(RESJ, encoding="utf-8"))
OUT = {}

# 1. placements
OUT["placements_gl"] = {a: {"dat_efc_moyen": g["dat_efc"]["moyenne"], "dat_efc_pic": g["dat_efc"]["pic"], "dat_pep_moyen": g["dat_pep"]["moyenne"],
                            "dat_orabank_moyen": g["dat_orabank"]["moyenne"], "fcp_moyen": g["fcp"]["moyenne"], "placements_moyens": g["placements"]["moyenne"],
                            "placements_cloture": g["placements"]["cloture"], "produits_financiers": g["produits_77"], "dat_efc_fin_de_mois": g["dat_efc_fin_de_mois"],
                            "fichier": g["fichier"]} for a, g in G.items()}
OUT["rendement_gl"] = [{"exercice": a, "produits_financiers": G[a]["produits_77"], "placements_moyens": G[a]["placements"]["moyenne"],
                        "rendement": G[a]["produits_77"] / G[a]["placements"]["moyenne"]} for a in range(2021, 2026)]
OUT["rendement_moyen_2021_2025"] = sum(x["produits_financiers"] for x in OUT["rendement_gl"]) / sum(x["placements_moyens"] for x in OUT["rendement_gl"])
OUT["dat_pep"] = {"montant": 505383858, "debouclage_2025": 638047119, "interets_2025": 132663262, "date_debouclage": "30/10/2025"}
anc = {x["exercice"]: x for x in R["rendement_placements"]}
OUT["rendement_clotures"] = {a: anc[a]["placements_moyens"] for a in anc}

# 2. provision du personnel plafonnée
SMIG_MENSUEL = 75000
PLAFOND = 25 * SMIG_MENSUEL * 12          # 2 080 heures par an = 12 mois de 173,33 heures
lig = [x for x in V9["pidr_individuel"] if x["annee"] == 2025]
cap = [x["dbo"] * min(1.0, PLAFOND / x["indemnite_projetee"]) if x["indemnite_projetee"] > 0 else 0.0 for x in lig]
OUT["pidr_plafond"] = {"smig_mensuel": SMIG_MENSUEL, "smig_annuel": SMIG_MENSUEL * 12, "plafond": PLAFOND, "agents": len(lig),
                       "agents_au_dessus": sum(1 for x in lig if x["indemnite_projetee"] > PLAFOND),
                       "dbo_sans_plafond": sum(x["dbo"] for x in lig), "dbo_avec_plafond": sum(cap),
                       "max_projete": max(x["indemnite_projetee"] for x in lig),
                       "provision_196": V9["pidr_synthese"][-1]["provision_196"]}

# 3. ratios par garantie
rg = []
for e in Q["excedent_par_risque"]:
    for gar, cot, pre, fr, ex in (("Maladie", "cot_maladie", "prest_maladie", "frais_maladie", "excedent_maladie"),
                                  ("Décès", "cot_deces", "prest_deces_et_prime", "frais_deces", "excedent_deces")):
        c = e[cot]
        rg.append({"exercice": e["exercice"], "garantie": gar, "cotisations": c, "prestations": e[pre], "frais": e[fr], "resultat_technique": e[ex],
                   "sinistralite": e[pre] / c if c else None, "taux_frais": e[fr] / c if c else None, "ratio_combine": (e[pre] + e[fr]) / c if c else None})
OUT["ratios_garantie"] = rg
s20 = next(x for x in R["sinistralite"] if x["exercice"] == 2020)
OUT["deces_2020"] = {"obseques": s20["prest_deces"], "prime_assurance": next(x for x in Q["excedent_par_risque"] if x["exercice"] == 2020)["prest_deces_et_prime"] - s20["prest_deces"]}
json.dump(OUT, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("rendement moyen 2021-2025 :", round(OUT["rendement_moyen_2021_2025"], 4), "| PIDR plafonnée :", round(OUT["pidr_plafond"]["dbo_avec_plafond"]),
      "| décès 2020 :", OUT["deces_2020"])
for x in rg[-4:]:
    print(x)
