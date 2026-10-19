# -*- coding: utf-8 -*-
"""
FPPN - Provisions, engagements et choix des formules de sinistralité (version 5 du rapport)
=========================================================================================

Calculs faisables avec les données reçues le 15/10/2026 :
1. inventaire des provisions et dépréciations au bilan, 2020 au 31/08/2026, et des comptes de dotations et
   reprises techniques prévus par la Décision n° 05/2022 (6901, 7901, 7906, 6906) ;
2. rapprochement de la provision pour sinistres à payer (PSAP) maladie avec les dettes envers les cliniques
   (compte 40110000) et avec les prestations maladie de l'année (nombre de mois de prestations couverts) ;
3. encours du Grenier par tranche de jours de retard au 31/12/2025 et au 31/08/2026 (base d'une dépréciation,
   taux du barème de la BCEAO en attente du texte) ;
4. compte de tiers GNA (47111106) : soldes et mouvements ;
5. provision pour indemnités de départ à la retraite du personnel (compte 196) : soldes, dotations, reprises ;
6. taux de sinistralité selon la formule de la Décision n° 05/2022 et selon ses variantes (par garantie,
   sans MEDOC, sans bons de prise en charge), de 2020 au 31/08/2026.

Utilisation : python provisions_fppn.py <dossier_base> <sortie_dossier>
Sorties : provisions.json et FPPN_Calculs_provisions.xlsx (onglets « 12_ »).
"""
import json
import os
import sys

import pandas as pd

BASE, SORTIE = sys.argv[1:3]
os.makedirs(SORTIE, exist_ok=True)
EX = list(range(2020, 2027))

bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal["nature"] == "comptes"].copy()
bal["net"] = bal["sol_d"] - bal["sol_c"]
LIB = bal.groupby("compte")["intitule"].last().str.strip()


def solde(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["net"].sum())


def mvt(ex, motif, sens):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["mvt_d" if sens == "d" else "mvt_c"].sum())


def existe(motif):
    return bool(bal["compte"].str.match(motif).any())


T, R = {}, {}

# 1. Inventaire des provisions et dépréciations ------------------------------------------------
comptes_prov = sorted(c for c in bal["compte"].unique() if pd.Series([c]).str.match(r"^(15|19|29|39|49|59)").iloc[0])
inv = []
for c in comptes_prov:
    ligne = {"compte": c, "intitule": LIB[c]}
    for ex in EX:
        ligne[str(ex)] = -solde(ex, "^" + c + "$")
    inv.append(ligne)
T["12_provisions_inventaire"] = pd.DataFrame(inv)
R["inventaire"] = inv
R["comptes_techniques_absents"] = {m: not existe("^" + m) for m in ("6901", "7901", "7906", "6906")}
psap_m = {ex: -solde(ex, r"^(19011|15801)") for ex in EX}
psap_d = {ex: -solde(ex, r"^(19012|15802)") for ex in EX}
R["psap_maladie"], R["psap_deces"] = psap_m, psap_d
R["psap_inchangee"] = len(set(round(v) for v in psap_m.values())) == 1 and len(set(round(v) for v in psap_d.values())) == 1

# 2. PSAP maladie, dettes envers les cliniques et prestations de l'année --------------------------
rap = []
for ex in EX:
    prest = solde(ex, r"^(601|6013|6018|60800007)")
    mois = 8 if ex == 2026 else 12
    clin = -solde(ex, r"^40110000$")
    rap.append({"exercice": ex, "psap_maladie": psap_m[ex], "dettes_cliniques_40110000": clin,
                "prestations_maladie": prest, "mois": mois,
                "psap_en_mois_de_prestations": psap_m[ex] / (prest / mois) if prest else None,
                "dettes_cliniques_en_mois": clin / (prest / mois) if prest else None,
                "psap_sur_dettes_cliniques": psap_m[ex] / clin if clin else None,
                "prestations_deces": solde(ex, r"^(6021|6024)"),
                "psap_deces_en_mois": psap_d[ex] / (solde(ex, r"^(6021|6024)") / mois) if solde(ex, r"^(6021|6024)") else None})
T["12_psap_rapprochement"] = pd.DataFrame(rap)
R["psap_rapprochement"] = rap

# 3. Grenier : encours par tranche de retard (base d'une dépréciation) ----------------------------
gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"))
gg = gb[gb.portefeuille == "global"]
TR = [(-1, 0, "À jour"), (0, 30, "1 à 30 jours"), (30, 90, "31 à 90 jours"), (90, 180, "91 à 180 jours"),
      (180, 365, "181 à 365 jours"), (365, 100000, "Plus de 365 jours")]
gre = []
for d in ("2025-12-31", "2026-08-31"):
    g = gg[gg.arrete == d]
    tot = g.encours.sum()
    for lo, hi, nom in TR:
        x = g[(g.retard > lo) & (g.retard <= hi)]
        gre.append({"arrete": d, "tranche": nom, "prets": int(len(x)), "encours": float(x.encours.sum()),
                    "part": float(x.encours.sum() / tot) if tot else None})
T["12_grenier_tranches"] = pd.DataFrame(gre)
R["grenier_tranches"] = gre
R["grenier_financement"] = {ex: solde(ex, r"^47111102$") for ex in EX}

# 4. Compte de tiers GNA ------------------------------------------------------------------------
gna = [{"exercice": ex, "solde_debiteur": solde(ex, r"^47111106$"), "mouvements_debit": mvt(ex, r"^47111106$", "d"),
        "mouvements_credit": mvt(ex, r"^47111106$", "c")} for ex in EX]
T["12_gna_47111106"] = pd.DataFrame(gna)
R["gna"] = gna

# 5. Provision pour indemnités de départ à la retraite du personnel ----------------------------------
pidr = [{"exercice": ex, "provision_196": -solde(ex, r"^196"), "dotation_6911": solde(ex, r"^6911"),
         "reprise_7911": -solde(ex, r"^7911")} for ex in EX]
T["12_pidr_196"] = pd.DataFrame(pidr)
R["pidr"] = pidr

# 6. Taux de sinistralité : formule du texte et variantes --------------------------------------------
sin = []
for ex in EX:
    num = solde(ex, r"^60") + solde(ex, r"^6901") + solde(ex, r"^7901")
    den = -solde(ex, r"^70") - solde(ex, r"^7906") + solde(ex, r"^6906")
    medoc = -solde(ex, r"^7022")
    bons = solde(ex, r"^(60800007|60800008)")
    autres60 = solde(ex, r"^60") - solde(ex, r"^(601|6013|6018|60800007)") - solde(ex, r"^(6021|6024)")
    cm, cd = -solde(ex, r"^7010"), -solde(ex, r"^7021")
    pm, pd_ = solde(ex, r"^(601|6013|6018|60800007)"), solde(ex, r"^(6021|6024)")
    sin.append({"exercice": ex, "charges_60": solde(ex, r"^60"), "dot_rep_6901_7901": solde(ex, r"^6901") + solde(ex, r"^7901"),
                "cotisations_70": -solde(ex, r"^70"), "dot_rep_7906_6906": -solde(ex, r"^7906") + solde(ex, r"^6906"),
                "numerateur": num, "denominateur": den, "taux_texte": num / den if den else None,
                "cot_medoc_7022": medoc, "taux_sans_medoc": num / (den - medoc) if den - medoc else None,
                "bons_60800007_60800008": bons, "taux_sans_bons": (num - bons) / den if den else None,
                "autres_charges_60": autres60,
                "cot_maladie_7010": cm, "prest_maladie": pm, "taux_maladie": pm / cm if cm else None,
                "cot_deces_7021": cd, "prest_deces": pd_, "taux_deces": pd_ / cd if cd else None})
T["12_sinistralite_variantes"] = pd.DataFrame(sin)
R["sinistralite"] = sin

json.dump(R, open(os.path.join(SORTIE, "provisions.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
with pd.ExcelWriter(os.path.join(SORTIE, "FPPN_Calculs_provisions.xlsx")) as xw:
    for k, v in T.items():
        v.to_excel(xw, sheet_name=k[:31], index=False)
print("provisions : ", {k: len(v) for k, v in T.items()}, "techniques absents :", R["comptes_techniques_absents"])
