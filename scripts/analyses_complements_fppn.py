# -*- coding: utf-8 -*-
"""
FPPN - Compléments d'analyses pour le rapport v2
================================================

Entrées : dossier de la base, resultats.json (analyses_fppn.py), série d'inflation Banque mondiale
          (FP.CPI.TOTL.ZG), tables de mortalité CIMA (CIMA H, CIMA F, TD, TV), dossier de l'extraction.
Sortie  : complements.json et onglets supplémentaires du classeur des calculs.

Contenu :
  - inflation 2020-2025 (Banque mondiale) et évolutions en termes réels ;
  - taux de croissance annuels moyens (TCAM) ;
  - tendances (régression linéaire, test de nullité de la pente, intervalle de confiance à 95 %) ;
  - intervalles de confiance (proportions, méthode de Wilson) ;
  - analyses de sensibilité (clé de répartition des cotisations, retraits d'épargne, pertes sur prêts) ;
  - indicateurs des tables de mortalité CIMA (espérance de vie, rente viagère, comparaison des tables) ;
  - extraits de données réelles (noms masqués) pour les diapositives de calcul pas à pas.
"""
import csv
import json
import math
import os
import sys

import numpy as np
import pandas as pd
import statsmodels.api as sm
from python_calamine import CalamineWorkbook

BASE, RES, INFL, TABLES, DONNEES, SORTIE = sys.argv[1:7]
R = json.load(open(RES, encoding="utf-8"))
C, T = {}, {}


def masque_nom(n):
    mots = str(n).split()
    return " ".join(m[0] + "." for m in mots[:3]) if mots else ""


def masque_mat(m):
    m = str(m)
    return m[:3] + "***" + m[-1] if len(m) >= 5 else m


# ---------------------------------------------------------------------------
# 1. Inflation (Banque mondiale, FP.CPI.TOTL.ZG, Côte d'Ivoire)
# ---------------------------------------------------------------------------
wb = CalamineWorkbook.from_path(INFL)
rows = wb.get_sheet_by_name("Data").to_python()
maj = rows[1][1]
hdr = [str(h).split(".")[0] for h in rows[3]]
civ = next(r for r in rows if len(r) > 1 and r[1] == "CIV")
infl = {int(h): float(civ[j]) for j, h in enumerate(hdr) if h.isdigit() and 2019 <= int(h) <= 2025 and civ[j] != ""}
C["inflation"] = {"source": "Banque mondiale, Indicateurs du développement dans le monde, FP.CPI.TOTL.ZG",
                  "mise_a_jour": str(maj), "valeurs": infl}
indice = {2020: 100.0}
for a in range(2021, 2026):
    indice[a] = indice[a - 1] * (1 + infl[a] / 100)
C["indice_prix_2020_100"] = indice

SIN = {s["exercice"]: s for s in R["sinistralite"]}
cot = {a: SIN[a]["cot_maladie"] + SIN[a]["cot_deces"] for a in range(2020, 2026)}
pm = {a: SIN[a]["prest_maladie"] for a in range(2020, 2026)}
reel = [{"exercice": a, "cotisations_nominales": cot[a], "cotisations_reelles_2020": cot[a] * 100 / indice[a],
         "prestations_maladie_nominales": pm[a], "prestations_maladie_reelles_2020": pm[a] * 100 / indice[a],
         "indice_prix": indice[a], "inflation": infl[a]} for a in range(2020, 2026)]
T["8_inflation_reel"] = pd.DataFrame(reel)
C["reel"] = reel
C["cot_croissance_nominale_2020_2025"] = cot[2025] / cot[2020] - 1
C["cot_croissance_reelle_2020_2025"] = (cot[2025] / indice[2025] * 100) / cot[2020] - 1
C["prix_croissance_2020_2025"] = indice[2025] / 100 - 1


# ---------------------------------------------------------------------------
# 2. Taux de croissance annuels moyens
# ---------------------------------------------------------------------------
def tcam(v0, v1, n):
    return (v1 / v0) ** (1 / n) - 1


BIL = {b["exercice"]: b for b in R["bilan"]}
GE = {g["arrete"]: g for g in R["grenier_encours"]}
gp = R["grenier_production"]
EFF = R["effectifs_mensuels"]
e0 = next(e for e in EFF if e["annee"] == 2024 and e["mois"] == 1)
e1 = next(e for e in EFF if e["annee"] == 2026 and e["mois"] == 8)
tc = [
    {"indicateur": "Cotisations maladie et décès", "debut": "2020", "fin": "2025", "v0": cot[2020], "v1": cot[2025], "annees": 5},
    {"indicateur": "Prestations maladie", "debut": "2021", "fin": "2025", "v0": pm[2021], "v1": pm[2025], "annees": 4},
    {"indicateur": "Épargne PEP et EFC (passif)", "debut": "2020", "fin": "2025", "v0": BIL[2020]["epargne_totale"], "v1": BIL[2025]["epargne_totale"], "annees": 5},
    {"indicateur": "Encours des prêts du Grenier", "debut": "31/12/2022", "fin": "31/12/2025", "v0": GE["31/12/2022"]["encours"], "v1": GE["31/12/2025"]["encours"], "annees": 3},
    {"indicateur": "Production de prêts du Grenier", "debut": "2022", "fin": "2025", "v0": gp["total"][1], "v1": gp["total"][4], "annees": 3},
    {"indicateur": "Cotisants actifs (maladie 936)", "debut": "janv. 2024", "fin": "août 2026", "v0": e0["cotisants"], "v1": e1["cotisants"], "annees": 31 / 12},
    {"indicateur": "Investissements au CMP", "debut": "2020", "fin": "2025", "v0": BIL[2020]["cmp_investissements"], "v1": BIL[2025]["cmp_investissements"], "annees": 5},
]
for t in tc:
    t["tcam"] = tcam(t["v0"], t["v1"], t["annees"])
    t["multiplicateur"] = t["v1"] / t["v0"]
T["8_tcam"] = pd.DataFrame(tc)
C["tcam"] = tc


# ---------------------------------------------------------------------------
# 3. Tendances : régression linéaire et test de la pente
# ---------------------------------------------------------------------------
def tendance(y, x=None, nom=""):
    y = np.asarray(y, dtype=float)
    x = np.arange(len(y), dtype=float) if x is None else np.asarray(x, dtype=float)
    X = sm.add_constant(x)
    m = sm.OLS(y, X).fit()
    lo, hi = m.conf_int(0.05)[1]
    return {"serie": nom, "n": int(len(y)), "pente": float(m.params[1]), "ic95_bas": float(lo), "ic95_haut": float(hi),
            "p_valeur": float(m.pvalues[1]), "r2": float(m.rsquared)}


pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
pre = pre[pre.retenu & (pre.annee >= 2020) & ((pre.annee < 2026) | (pre.mois <= 8))]
tend = []
eff = [e for e in EFF if e["cotisants"]]
x = [(e["annee"] - 2024) * 12 + e["mois"] - 1 for e in eff]
tend.append(tendance([e["cotisants"] for e in eff], x, "Cotisants actifs 936, par mois (janv. 2024 à août 2026)"))
sc = [SIN[a]["sc_maladie"] * 100 for a in range(2020, 2026)]
tend.append(tendance(sc, list(range(2020, 2026)), "Ratio S/C maladie, par an (2020 à 2025), en points de %"))
lp = [math.log(pm[a]) for a in range(2021, 2026)]
t_lp = tendance(lp, list(range(2021, 2026)), "Logarithme des prestations maladie (2021 à 2025)")
t_lp["croissance_annuelle"] = math.exp(t_lp["pente"]) - 1
t_lp["croissance_ic95_bas"] = math.exp(t_lp["ic95_bas"]) - 1
t_lp["croissance_ic95_haut"] = math.exp(t_lp["ic95_haut"]) - 1
tend.append(t_lp)
g9 = pre[(pre.poste == "9X4") & (pre.sens == "retour")].groupby(["annee", "mois"]).montant.sum().reset_index()
tend.append(tendance(g9.montant, (g9.annee - 2023) * 12 + g9.mois - 1, "Précomptes Grenier 9X4 retour, par mois (2023 à 2026), en FCFA"))
pv = pre[(pre.poste == "998") & (pre.sens == "retour")].groupby(["annee", "mois"]).montant.sum().reset_index()
tend.append(tendance(pv.montant, (pv.annee - 2020) * 12 + pv.mois - 1, "Versements PEP 998, par mois (2020 à 2026), en FCFA"))
T["8_tendances"] = pd.DataFrame(tend)
C["tendances"] = tend


# ---------------------------------------------------------------------------
# 4. Intervalles de confiance (Wilson)
# ---------------------------------------------------------------------------
def wilson(k, n, z=1.959964):
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return p, c - h, c + h


gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"))
g8 = gb[(gb.arrete == "2026-08-31") & (gb.portefeuille == "global")]
k30 = int((g8.retard > 30).sum())
p, lo, hi = wilson(k30, len(g8))
k180 = int((g8.retard > 180).sum())
p2, lo2, hi2 = wilson(k180, len(g8))
C["ic_grenier"] = {"prets": int(len(g8)), "retard30": k30, "p30": p, "p30_bas": lo, "p30_haut": hi,
                   "retard180": k180, "p180": p2, "p180_bas": lo2, "p180_haut": hi2}

# ---------------------------------------------------------------------------
# 5. Sensibilités
# ---------------------------------------------------------------------------
s25 = SIN[2025]
tot = s25["cot_maladie"] + s25["cot_deces"]
cle_compta = s25["cot_maladie"] / tot
cle_solde = 3.77 / 5
sens_cle = {"cle_comptable_maladie": cle_compta, "cle_solde_maladie": cle_solde,
            "sc_maladie_compta": s25["sc_maladie"], "sc_deces_compta": s25["sc_deces"],
            "sc_maladie_cle_solde": s25["prest_maladie"] / (tot * cle_solde),
            "sc_deces_cle_solde": s25["prest_deces"] / (tot * (1 - cle_solde)),
            "sc_maladie_cle_bareme": s25["prest_maladie"] / (tot * 0.7),
            "sc_deces_cle_bareme": s25["prest_deces"] / (tot * 0.3)}
C["sensibilite_cle"] = sens_cle

b26 = BIL[2026]
retraits = []
for taux in (0.05, 0.10, 0.20, 0.30):
    besoin = b26["epargne_totale"] * taux
    retraits.append({"taux_retrait": taux, "besoin": besoin, "actifs_dedies": b26["actifs_dedies_epargne"],
                     "tresorerie_totale": b26["tresorerie_totale"], "ecart_actifs_dedies": b26["actifs_dedies_epargne"] - besoin,
                     "ecart_tresorerie": b26["tresorerie_totale"] - besoin})
T["8_sensibilite_retraits"] = pd.DataFrame(retraits)
C["sensibilite_retraits"] = retraits
C["retrait_max_couvert"] = {"par_actifs_dedies": b26["actifs_dedies_epargne"] / b26["epargne_totale"],
                            "par_tresorerie": b26["tresorerie_totale"] / b26["epargne_totale"]}

enc180 = float(g8[g8.retard > 180].encours.sum())
enc90 = float(g8[g8.retard > 90].encours.sum())
pertes = [{"perte": t, "provision_180j": enc180 * t, "provision_90j": enc90 * t} for t in (0.25, 0.50, 1.00)]
T["8_sensibilite_pertes_grenier"] = pd.DataFrame(pertes)
C["sensibilite_pertes_grenier"] = {"encours_180j": enc180, "encours_90j": enc90, "scenarios": pertes}

mois_prov = []
for a in range(2020, 2027):
    s = SIN[a]
    nb_mois = 8 if a == 2026 else 12
    mensuel = (s["prest_maladie"] + s["prest_deces"]) / nb_mois
    mois_prov.append({"exercice": a, "provisions": BIL[a]["provisions_techniques"], "prestations_mensuelles": mensuel,
                      "mois_couverts": BIL[a]["provisions_techniques"] / mensuel if mensuel else None})
T["8_provisions_mois"] = pd.DataFrame(mois_prov)
C["provisions_mois"] = mois_prov

# ---------------------------------------------------------------------------
# 6. Tables de mortalité CIMA
# ---------------------------------------------------------------------------
tables = {}
for nom, f in (("CIMA H", "CIMA_H.xlsx"), ("CIMA F", "CIMA_F.xlsx"), ("TD CIMA", "TD_CIMA.xlsx"), ("TV CIMA", "TV_CIMA.xlsx")):
    w = CalamineWorkbook.from_path(os.path.join(TABLES, f))
    s = w.get_sheet_by_name(w.sheet_names[0]).to_python()
    lx = {int(r[0]): float(r[1]) for r in s[1:] if r[0] != ""}
    tables[nom] = {"feuille": w.sheet_names[0], "lx": lx}


def e_x(lx, x):
    """Espérance de vie résiduelle (abrégée + 1/2)."""
    return sum(lx.get(k, 0) for k in range(x + 1, 111)) / lx[x] + 0.5


def a_x(lx, x, i=0.035):
    """Valeur actuelle d'une rente viagère de 1 par an payable d'avance (ä_x)."""
    v = 1 / (1 + i)
    return sum(lx.get(x + k, 0) * v ** k for k in range(0, 111 - x)) / lx[x]


mort = []
for nom, t in tables.items():
    lx = t["lx"]
    mort.append({"table": nom, "feuille": t["feuille"], "e0": e_x(lx, 0), "e55": e_x(lx, 55), "e60": e_x(lx, 60),
                 "a55": a_x(lx, 55), "a60": a_x(lx, 60), "a65": a_x(lx, 65),
                 "q40": 1 - lx[41] / lx[40], "q55": 1 - lx[56] / lx[55],
                 "deces_20_60": (lx[20] - lx[60]) / lx[20], "survie_60": lx[60] / lx[0]})
T["8_tables_mortalite"] = pd.DataFrame(mort)
C["mortalite"] = mort
M_ = {m["table"]: m for m in mort}
C["mortalite_comparaison"] = {
    "deces_20_60_H_vs_TD": M_["CIMA H"]["deces_20_60"] / M_["TD CIMA"]["deces_20_60"] - 1,
    "deces_20_60_F_vs_TV": M_["CIMA F"]["deces_20_60"] / M_["TV CIMA"]["deces_20_60"] - 1}
ages = list(range(0, 111, 5))
C["courbes_survie"] = {"ages": ages, **{n: [t["lx"][a] / 1e6 * 100 for a in ages] for n, t in tables.items()}}
ages_q = list(range(20, 81, 5))
C["courbes_qx"] = {"ages": ages_q, **{n: [(1 - t["lx"][a + 1] / t["lx"][a]) * 1000 for a in ages_q] for n, t in tables.items()}}

# ---------------------------------------------------------------------------
# 7. Extraits pour les calculs pas à pas (noms et matricules masqués)
# ---------------------------------------------------------------------------
J = {l["id"]: l for l in csv.DictReader(open(os.path.join(DONNEES, "journal_fichiers.csv"), encoding="utf-8-sig"), delimiter=";")}
F = {}
for f in csv.DictReader(open(os.path.join(DONNEES, "journal_feuilles.csv"), encoding="utf-8-sig"), delimiter=";"):
    F.setdefault(f["id"], []).append(f)
EX = {}

# a) Ligne de total : fichier maladie 936 de février 2025
src = next(i for i, l in J.items() if l["fichier"] == "FEVRIER 2025_effectif_police_936.xlsx")
lignes = list(csv.reader(open(os.path.join(DONNEES, "tableaux", F[src][0]["csv"]), encoding="utf-8"), delimiter=";"))
dern = [l for l in lignes if any(c.strip() for c in l)][-4:]
EX["ligne_total"] = {"fichier": J[src]["dossier"] + "\\" + J[src]["fichier"], "id": src, "entete": lignes[0][:8],
                     "lignes": [[masque_mat(c) if j == 3 and any(ch.isalpha() for ch in c) else (masque_nom(c) if j == 4 else c)
                                 for j, c in enumerate(l[:8])] for l in dern],
                     "nb_lignes_fichier": len([l for l in lignes[1:] if any(c.strip() for c in l)])}
# b) Montants doublés (décès avril 2024)
p_all = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
cor = p_all[p_all.corrige].sort_values("montant_source", ascending=False).head(4)
EX["montants_doubles"] = {"fichier": "FPPN-SOLDE\\ALLOCATION DECES\\2024\\FPPN_ALLOC_DECES_8Z6_AVRIL_2024.xlsx",
                          "lignes": [[masque_mat(r.matricule), masque_nom(r.nom), f"{r.montant_source:.0f}", f"{r.montant:.0f}"]
                                     for r in cor.itertuples()], "total_corriges": int(p_all.corrige.sum())}
# c) Taux de cotisation : un agent, avril 2024
eff_d = pd.read_parquet(os.path.join(BASE, "effectifs_detail.parquet"))
a4 = eff_d[(eff_d.annee == 2024) & (eff_d.mois == 4) & eff_d.salaire_base.notna()]
z = pre[(pre.poste == "8Z6") & (pre.annee == 2024) & (pre.mois == 4)][["matricule", "montant"]].rename(columns={"montant": "m8z6"})
ex = a4.merge(z, on="matricule").head(3)
EX["taux_cotisation"] = {"lignes": [[masque_mat(r.matricule), r.grade, f"{r.salaire_base:.0f}", f"{r.cotisation_936:.0f}", f"{r.m8z6:.0f}",
                                     f"{100 * r.cotisation_936 / r.salaire_base:.2f}", f"{100 * r.m8z6 / r.salaire_base:.2f}",
                                     f"{100 * (r.cotisation_936 + r.m8z6) / r.salaire_base:.2f}"] for r in ex.itertuples()]}
# d) Comptes du ratio S/C maladie 2025, de la couverture de l'épargne et des fonds propres
bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal.nature == "comptes"].copy()
bal["net"] = bal.sol_d - bal.sol_c


def comptes(ex, motif):
    g = bal[(bal.exercice == ex) & bal.compte.str.match(motif)]
    return [[r.compte, r.intitule, f"{r.net:.0f}"] for r in g.itertuples() if abs(r.net) > 0]


EX["sc_2025"] = {"prestations": comptes(2025, r"^(601|6013|6018|60800007)"), "cotisations": comptes(2025, r"^7010")}
EX["epargne_2026"] = {"passif": comptes(2026, r"^41195"), "actifs": comptes(2026, r"^(5211020|5221020|5211030|5221030|526120|5330)")}
EX["fonds_propres_2025"] = {"comptes": comptes(2025, r"^1[0-2]")}
EX["provisions"] = {str(a): comptes(a, r"^(15801|15802|19011|19012)") for a in range(2020, 2027)}
# e) Balance âgée : prêts en retard (exemples) au 31/08/2026
ret = g8[g8.retard > 180].sort_values("retard", ascending=False).head(4)
EX["prets_retard"] = {"lignes": [[masque_mat(r.mecano), masque_nom(r.nom), f"{r.montant_debloque:.0f}", f"{r.encours:.0f}", f"{r.retard:.0f}"]
                                 for r in ret.itertuples()]}
# f) Mois le plus bas du rapprochement aller / retour (PEP)
inv = pd.read_parquet(os.path.join(BASE, "sources_mois.parquet"))
iv = inv[inv.retenu & (inv.poste == "998")]
a_ = iv[iv.sens == "aller"].set_index(["annee", "mois"])
r_ = iv[iv.sens == "retour"].set_index(["annee", "mois"])
jj = a_[["montant", "lignes", "fichier"]].join(r_[["montant", "lignes", "fichier"]], lsuffix="_a", rsuffix="_r", how="inner")
jj["taux"] = jj.montant_r / jj.montant_a
lo_ = jj.taux.idxmin()
hi_ = jj.taux.idxmax()
EX["aller_retour_pep"] = {"min": {"mois": list(lo_), **{k: (float(v) if isinstance(v, (int, float, np.floating)) else v) for k, v in jj.loc[lo_].items()}},
                          "max": {"mois": list(hi_), **{k: (float(v) if isinstance(v, (int, float, np.floating)) else v) for k, v in jj.loc[hi_].items()}}}
# g) Entrées de juillet (936)
fl = {(f["annee"], f["m"]): f for f in R["flux_mensuels"]}
EX["entrees_juillet"] = {"2024": fl.get((2024, 7)), "2025": fl.get((2025, 7))}
C["extraits"] = EX


def conv(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return None if np.isnan(o) else float(o)
    return str(o)


os.makedirs(SORTIE, exist_ok=True)
json.dump(C, open(os.path.join(SORTIE, "complements.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=conv)
with pd.ExcelWriter(os.path.join(SORTIE, "FPPN_Calculs_complements.xlsx")) as w:
    for k, v in T.items():
        v.to_excel(w, sheet_name=k[:31], index=False)
print("complements :", len(C), "blocs ;", len(T), "tableaux")
