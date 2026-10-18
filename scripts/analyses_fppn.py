# -*- coding: utf-8 -*-
"""
FPPN - Analyses actuarielles et statistiques sur les données reçues le 15/10/2026
=================================================================================

Entrée : le dossier de la base (base_fppn.py, grenier_fppn.py, compta_fppn.py).
Sortie : resultats.json (chiffres et séries des graphiques du rapport) et
         FPPN_Calculs_rapport.xlsx (un onglet par tableau, avec la source de chaque chiffre).

Période d'analyse : janvier 2020 à août 2026. Montants en FCFA.
"""
import json
import os
import sys

import numpy as np
import pandas as pd

BASE = sys.argv[1] if len(sys.argv) > 1 else "base"
SORTIE = sys.argv[2] if len(sys.argv) > 2 else "resultats"
os.makedirs(SORTIE, exist_ok=True)
MOIS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]
R, T = {}, {}           # R : chiffres et séries pour le rapport ; T : tableaux pour le classeur


def dans_periode(df, a="annee", m="mois"):
    return df[(df[a] >= 2020) & ((df[a] < 2026) | (df[m] <= 8))]


def lib_mois(a, m):
    return f"{MOIS[m - 1]} {str(a)[2:]}"


# ---------------------------------------------------------------------------
# Chargement
# ---------------------------------------------------------------------------
pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
pre = dans_periode(pre[pre["retenu"]])
inv = pd.read_parquet(os.path.join(BASE, "sources_mois.parquet"))
eff = pd.read_parquet(os.path.join(BASE, "effectifs_detail.parquet"))
rej = pd.read_parquet(os.path.join(BASE, "rejets.parquet"))
gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"))
gp = pd.read_parquet(os.path.join(BASE, "grenier_production.parquet"))
bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal["nature"] == "comptes"].copy()
bal["net"] = bal["sol_d"] - bal["sol_c"]
EX = list(range(2020, 2027))


def solde(ex, motif):
    g = bal[(bal["exercice"] == ex) & bal["compte"].str.match(motif)]
    return float(g["net"].sum())


# ---------------------------------------------------------------------------
# 1. Données : couverture et rapprochement aller / retour
# ---------------------------------------------------------------------------
inv_p = dans_periode(inv[inv["retenu"]])
cov = inv_p.groupby(["poste", "sens"]).agg(mois_recus=("mois", "count"), lignes=("lignes", "sum"),
                                           montant=("montant", "sum")).reset_index()
T["1_couverture"] = cov
R["couverture"] = cov.to_dict("records")

rap = []
for poste in ["870", "998", "9X4", "9W3", "CGRAE", "8Z7"]:
    a = inv_p[(inv_p.poste == poste) & (inv_p.sens == "aller")].set_index(["annee", "mois"])
    r = inv_p[(inv_p.poste == poste) & (inv_p.sens == "retour")].set_index(["annee", "mois"])
    j = a[["lignes", "montant"]].join(r[["lignes", "montant"]], lsuffix="_aller", rsuffix="_retour", how="inner")
    if len(j):
        rap.append({"poste": poste, "mois_communs": len(j),
                    "montant_aller": j.montant_aller.sum(), "montant_retour": j.montant_retour.sum(),
                    "taux_recouvrement": j.montant_retour.sum() / j.montant_aller.sum(),
                    "taux_min_mensuel": (j.montant_retour / j.montant_aller).min(),
                    "taux_max_mensuel": (j.montant_retour / j.montant_aller).max()})
T["1_aller_retour"] = pd.DataFrame(rap)
R["aller_retour"] = rap

anom = rej.copy()
anom["motif_court"] = anom["motif"].str.replace(r" :.*", "", regex=True)
T["1_rejets"] = anom.groupby("motif_court").agg(entrees=("n", "count"), lignes=("n", "sum")).reset_index()
tot = anom[anom.motif_court == "ligne de total ou sans nom"]
R["lignes_total_ecartees"] = {"n": int(len(tot)), "montant": float(tot["montant"].fillna(0).sum())}
R["montants_corriges_8Z6"] = int(pd.read_parquet(os.path.join(BASE, "precomptes.parquet"),
                                                 columns=["corrige"])["corrige"].sum())

# ---------------------------------------------------------------------------
# 2. Effectifs et cotisations
# ---------------------------------------------------------------------------
m936 = pre[pre.poste == "936"]
e = m936.groupby(["annee", "mois"]).agg(cotisants=("matricule", "nunique"), montant=("montant", "sum")).reset_index()
cg = pre[(pre.poste == "CGRAE") & (pre.sens == "retour")].groupby(["annee", "mois"]).agg(
    retraites=("matricule", "nunique"), montant_cgrae=("montant", "sum")).reset_index()
e = e.merge(cg, on=["annee", "mois"], how="outer").sort_values(["annee", "mois"])
e["libelle"] = [lib_mois(a, m) for a, m in zip(e.annee, e.mois)]
T["2_effectifs_mensuels"] = e
R["effectifs_mensuels"] = e.fillna(np.nan).replace({np.nan: None}).to_dict("records")

# Entrées et sorties entre deux mois consécutifs reçus
mois_ok = sorted(set(zip(m936.annee, m936.mois)))
ens = {k: set(g["matricule"]) for k, g in m936.groupby(["annee", "mois"])}
flux = []
for (a0, m0), (a1, m1) in zip(mois_ok, mois_ok[1:]):
    if (a1 * 12 + m1) - (a0 * 12 + m0) != 1:
        continue
    s0, s1 = ens[(a0, m0)], ens[(a1, m1)]
    flux.append({"mois": lib_mois(a1, m1), "annee": a1, "m": m1, "entrees": len(s1 - s0), "sorties": len(s0 - s1),
                 "effectif": len(s1)})
flux = pd.DataFrame(flux)
T["2_entrees_sorties"] = flux
R["flux_annuels"] = flux.groupby("annee")[["entrees", "sorties"]].sum().reset_index().to_dict("records")
R["flux_mensuels"] = flux.to_dict("records")

# Taux de cotisation maladie (936) rapporté au salaire de base, mois détaillés
d = eff.dropna(subset=["salaire_base", "cotisation_936"])
d = d[d.salaire_base > 0]
d["taux"] = d.cotisation_936 / d.salaire_base
R["taux_936"] = {"mois_detailles": sorted({lib_mois(a, m) for a, m in zip(d.annee, d.mois)}),
                 "lignes": int(len(d)), "mediane": float(d.taux.median()),
                 "part_3_5_pct": float(((d.taux - 0.035).abs() <= 0.0005).mean()),
                 "q05": float(d.taux.quantile(0.05)), "q95": float(d.taux.quantile(0.95))}
hist = pd.cut(d.taux * 100, bins=[0, 3, 3.4, 3.45, 3.55, 3.6, 4, 100],
              labels=["< 3 %", "3 à 3,4 %", "3,4 à 3,45 %", "3,45 à 3,55 %", "3,55 à 3,6 %", "3,6 à 4 %", "> 4 %"]
              ).value_counts().sort_index()
T["2_taux_936"] = hist.rename("lignes").reset_index()
R["taux_936_hist"] = {"categories": [str(x) for x in hist.index], "valeurs": [int(x) for x in hist.values]}

# Rapport décès (8Z6) / maladie (936) par mutualiste, même mois
dd = pre[pre.poste.isin(["8Z6", "936"])].pivot_table(index=["annee", "mois", "matricule"], columns="poste",
                                                     values="montant", aggfunc="sum").dropna()
dd = dd[dd["936"] > 0]
dd["r"] = dd["8Z6"] / dd["936"]
R["rapport_8Z6_936"] = {"mediane": float(dd.r.median()), "theorique": 1.5 / 3.5, "lignes": int(len(dd)),
                        "part_theorique": float(((dd.r - 1.5 / 3.5).abs() <= 0.01).mean())}
rq = dd.reset_index().groupby(["annee"]).r.median()
T["2_rapport_8Z6_936"] = rq.reset_index()

# Grades et ancienneté (dernier mois détaillé)
dern = eff.sort_values(["annee", "mois"]).groupby(["annee", "mois"]).size().index[-1]
g_last = eff[(eff.annee == dern[0]) & (eff.mois == dern[1])].copy()
ref = pd.Timestamp(year=dern[0], month=dern[1], day=1)
g_last["anciennete"] = (ref - g_last.date_recrutement).dt.days / 365.25
anc = pd.cut(g_last.anciennete, bins=[-1, 5, 10, 15, 20, 25, 30, 35, 60],
             labels=["0-5 ans", "5-10", "10-15", "15-20", "20-25", "25-30", "30-35", "35 ans et +"]).value_counts().sort_index()
R["anciennete"] = {"mois": lib_mois(*dern), "categories": [str(x) for x in anc.index],
                   "valeurs": [int(x) for x in anc.values], "mediane": float(g_last.anciennete.median()),
                   "salaire_base_median": float(g_last.salaire_base.median())}
T["2_anciennete"] = anc.rename("agents").reset_index()

# Cotisations comptabilisées et précomptées (année complète 2025)
c = []
for ex in EX:
    c.append({"exercice": ex, "cot_maladie_compta": -solde(ex, r"^7010"), "cot_deces_compta": -solde(ex, r"^7021")})
cot = pd.DataFrame(c)
p_an = pre[pre.sens == "retour"].groupby(["poste", "annee"]).agg(montant=("montant", "sum"),
                                                                 mois=("mois", "nunique")).reset_index()
for poste, col in [("936", "prec_936"), ("8Z6", "prec_8Z6"), ("CGRAE", "prec_cgrae")]:
    x = p_an[p_an.poste == poste].set_index("annee")
    cot[col] = cot.exercice.map(x["montant"])
    cot[col + "_mois"] = cot.exercice.map(x["mois"])
T["2_cotisations"] = cot
R["cotisations"] = cot.replace({np.nan: None}).to_dict("records")

# ---------------------------------------------------------------------------
# 3. Sinistralité, frais et résultat (comptabilité)
# ---------------------------------------------------------------------------
s = []
for ex in EX:
    cm, cd = -solde(ex, r"^7010"), -solde(ex, r"^7021")
    pm = solde(ex, r"^(601|6013|6018|60800007)")
    pdc = solde(ex, r"^(6021|6024)")
    prime_ass = solde(ex, r"^6518")
    frais = solde(ex, r"^(61|62|63|64|65|66)") - prime_ass
    produits = -solde(ex, r"^7")
    charges = solde(ex, r"^6")
    hao = -solde(ex, r"^8")
    s.append({"exercice": ex, "cot_maladie": cm, "prest_maladie": pm, "sc_maladie": pm / cm if cm else None,
              "cot_deces": cd, "prest_deces": pdc, "prime_assurance_deces": prime_ass,
              "sc_deces": pdc / cd if cd else None, "sc_deces_avec_prime": (pdc + prime_ass) / cd if cd else None,
              "frais_gestion": frais, "taux_frais": frais / (cm + cd) if (cm + cd) else None,
              "produits": produits, "charges": charges, "resultat_hao": hao,
              "resultat": produits - charges + hao})
sin = pd.DataFrame(s)
T["3_sinistralite_resultat"] = sin
R["sinistralite"] = sin.replace({np.nan: None}).to_dict("records")

natures = {"Pharmacie": r"^60111", "Consultations": r"^60112", "Hospitalisations": r"^60113",
           "Analyses médicales": r"^60114", "Radio et imagerie": r"^60115",
           "Autres prestations maladie": r"^6012", "Autres (transport, rééducation, bons)": r"^(60116|60118|60119|6013|6018|60800007)"}
nat = pd.DataFrame([{"nature": k, **{str(ex): solde(ex, v) for ex in EX}} for k, v in natures.items()])
T["3_prestations_maladie_nature"] = nat
R["prestations_maladie_nature"] = nat.to_dict("records")

# ---------------------------------------------------------------------------
# 4. Bilan : provisions, épargne, placements, créances, fonds propres
# ---------------------------------------------------------------------------
bl = []
for ex in EX:
    bl.append({
        "exercice": ex,
        "provision_maladie": -solde(ex, r"^(15801|19011)"), "provision_obseques": -solde(ex, r"^(15802|19012)"),
        "epargne_pep": -solde(ex, r"^411951"), "epargne_efc": -solde(ex, r"^411952"),
        "banque_efc": solde(ex, r"^(5211020|5221020)"), "banque_pep": solde(ex, r"^(5211030|5221030|526120)"),
        "fcp": solde(ex, r"^5330"), "tresorerie_totale": solde(ex, r"^5[1-7]"),
        "creance_hopital_police": solde(ex, r"^47111100"), "cmp_investissements": solde(ex, r"^(47111101|47111104|47111105)"),
        "grenier_financement": solde(ex, r"^47111102"), "cotisations_a_recevoir_etat": solde(ex, r"^4111100"),
        "terrains": solde(ex, r"^222"), "immobilisations_nettes": solde(ex, r"^2"),
        "fournisseurs_cliniques": -solde(ex, r"^40110"),
        "reserves": -solde(ex, r"^1[01]"), "report_a_nouveau": -solde(ex, r"^12[01]"),
        "pertes_cumulees": -solde(ex, r"^129"),
        "prets_pep": solde(ex, r"^2711[0-9]|^27120"),
        "total_actif": float(bal[(bal.exercice == ex) & bal.compte.str.match(r"^[2-5]")].net.clip(lower=0).sum()),
    })
bil = pd.DataFrame(bl)
bil["epargne_totale"] = bil.epargne_pep + bil.epargne_efc
bil["actifs_dedies_epargne"] = bil.banque_efc + bil.banque_pep + bil.fcp
bil["couverture_epargne"] = bil.actifs_dedies_epargne / bil.epargne_totale
bil["provisions_techniques"] = bil.provision_maladie + bil.provision_obseques
bil["fonds_propres_hors_resultat"] = bil.reserves + bil.report_a_nouveau + bil.pertes_cumulees
bil["resultat_exercice"] = bil.exercice.map(sin.set_index("exercice").resultat)
bil["fonds_propres"] = bil.fonds_propres_hors_resultat + bil.resultat_exercice
T["4_bilan"] = bil
R["bilan"] = bil.replace({np.nan: None}).to_dict("records")

# Épargne PEP : versements cumulés connus (fichiers 2001 à 2026) comparés au compte 41195100
pep_all = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
pep_all = pep_all[(pep_all.poste == "998") & (pep_all.sens == "retour") & pep_all.retenu
                  & ((pep_all.annee < 2026) | (pep_all.mois <= 8))]
pa = pep_all.groupby("annee").agg(versements=("montant", "sum"), epargnants=("matricule", "nunique")).reset_index()
pa["cumul"] = pa.versements.cumsum()
pa["compte_41195100"] = pa.annee.map(bil.set_index("exercice").epargne_pep)
T["4_pep_versements"] = pa
R["pep_versements"] = pa.replace({np.nan: None}).to_dict("records")
efc = pre[(pre.poste == "870") & (pre.sens == "retour")].groupby("annee").agg(
    versements=("montant", "sum"), epargnants=("matricule", "nunique"), mois=("mois", "nunique")).reset_index()
efc["compte_41195200"] = efc.annee.map(bil.set_index("exercice").epargne_efc)
efc["variation_compte"] = efc.compte_41195200.diff()
T["4_efc_versements"] = efc
R["efc_versements"] = efc.replace({np.nan: None}).to_dict("records")

# Références des produits sous le poste 998 (PEP, PTI, PTQ, PTS…)
refs = pep_all.assign(produit=pep_all.reference.str.upper().str.extract(r"^([A-Z\-]+)")[0].fillna("sans référence"))
rf = refs[refs.annee >= 2020].groupby("produit").agg(lignes=("montant", "size"), montant=("montant", "sum"),
                                                     mutualistes=("matricule", "nunique")).sort_values(
    "montant", ascending=False).reset_index()
T["4_poste_998_references"] = rf
R["poste_998_references"] = rf.head(8).to_dict("records")

# ---------------------------------------------------------------------------
# 5. Grenier du Policier
# ---------------------------------------------------------------------------
gg = gb[gb.portefeuille == "global"].copy()
evo = gg.groupby("arrete").agg(prets=("compte", "count"), encours=("encours", "sum"),
                               debloque=("montant_debloque", "sum")).reset_index()
evo["financement_fppn_47111102"] = [solde(d.year, r"^47111102") if d.month == 12 or d == pd.Timestamp("2026-08-31")
                                    else None for d in evo.arrete]
T["5_grenier_encours"] = evo
R["grenier_encours"] = [{**r, "arrete": r["arrete"].strftime("%d/%m/%Y")} for r in evo.replace({np.nan: None}).to_dict("records")]

g8 = gg[gg.arrete == "2026-08-31"].copy()
tranches = pd.cut(g8.retard, bins=[-1, 0, 30, 90, 180, 365, 10000],
                  labels=["À jour", "1 à 30 j", "31 à 90 j", "91 à 180 j", "181 à 365 j", "Plus de 365 j"])
par = g8.groupby(tranches, observed=False).agg(prets=("compte", "count"), encours=("encours", "sum")).reset_index()
par["part_encours"] = par.encours / par.encours.sum()
T["5_grenier_retards"] = par.rename(columns={"retard": "tranche"})
R["grenier_retards"] = [{"tranche": str(r["retard"]), "prets": int(r["prets"]), "encours": float(r["encours"]),
                         "part": float(r["part_encours"])} for r in par.to_dict("records")]
R["grenier_par"] = {"par30": float(g8[g8.retard > 30].encours.sum() / g8.encours.sum()),
                    "par90": float(g8[g8.retard > 90].encours.sum() / g8.encours.sum()),
                    "par0": float(g8[g8.retard > 0].encours.sum() / g8.encours.sum()),
                    "encours": float(g8.encours.sum()), "prets": int(len(g8)),
                    "top10": float(g8.nlargest(10, "encours").encours.sum() / g8.encours.sum()),
                    "taux_median": float(g8.taux.median()), "duree_mediane": float(g8.duree_mois.median()),
                    "emprunteurs": int(g8.mecano.nunique())}
# Retards dans le temps (fins d'année et mois 2026)
pt = []
for d, g in gg.groupby("arrete"):
    pt.append({"arrete": d.strftime("%d/%m/%Y"), "par30": float(g[g.retard > 30].encours.sum() / g.encours.sum()),
               "par90": float(g[g.retard > 90].encours.sum() / g.encours.sum())})
T["5_grenier_par_historique"] = pd.DataFrame(pt)
R["grenier_par_historique"] = pt

# Production annuelle par type de crédit (fichiers annuels 2021-2025 et mensuels 2026 par portefeuille)
prod = gp.copy()
prod["annee"] = prod.periode_au.dt.year
annuel = prod[prod.portefeuille == "global"]
mens26 = prod[(prod.annee == 2026) & (prod.portefeuille != "global") & (prod.periode_au <= "2026-08-31")]
pr = pd.concat([annuel, mens26])
pa_type = pr.groupby(["annee", "type_credit"]).montant_debloque.sum().unstack(fill_value=0)
T["5_grenier_production"] = pa_type.reset_index()
R["grenier_production"] = {"annees": [int(x) for x in pa_type.index], "types": list(pa_type.columns),
                           "valeurs": {c: [float(v) for v in pa_type[c]] for c in pa_type.columns},
                           "total": [float(v) for v in pa_type.sum(axis=1)],
                           "nombre": [int(v) for v in pr.groupby("annee").size()]}

# Échéances dues (balance âgée) comparées aux précomptes 9X4 retour, mois 2026
ech = []
for d, g in gg[gg.arrete.dt.year == 2026].groupby("arrete"):
    p9 = pre[(pre.poste == "9X4") & (pre.sens == "retour") & (pre.annee == d.year) & (pre.mois == d.month)]
    actifs = g[g.encours > 0]
    ech.append({"mois": lib_mois(d.year, d.month), "echeances_balance": float(actifs.montant_echeance.sum()),
                "precompte_9x4": float(p9.montant.sum()) if len(p9) else None,
                "emprunteurs_balance": int(actifs.mecano.nunique()), "precomptes_9x4": int(p9.mat6.nunique())})
T["5_grenier_echeances_precomptes"] = pd.DataFrame(ech)
R["grenier_echeances"] = ech

# ---------------------------------------------------------------------------
# 6. Retraités CGRAE et produits annexes (9W3, 8Z7)
# ---------------------------------------------------------------------------
cgr = pre[(pre.poste == "CGRAE") & (pre.sens == "retour")]
cga = cgr.groupby(["annee", "mois"]).agg(retraites=("matricule", "nunique"), montant=("montant", "sum")).reset_index()
cga["moyen"] = cga.montant / cga.retraites
cga["libelle"] = [lib_mois(a, m) for a, m in zip(cga.annee, cga.mois)]
T["6_cgrae"] = cga
R["cgrae"] = cga.to_dict("records")
ann = pre[pre.poste.isin(["9W3", "8Z7"]) & (pre.sens == "retour")].groupby(["poste", "annee"]).agg(
    adherents=("matricule", "nunique"), montant=("montant", "sum"), mois=("mois", "nunique")).reset_index()
T["6_produits_annexes"] = ann
R["produits_annexes"] = ann.to_dict("records")

# ---------------------------------------------------------------------------
# 7. Données externes validées
# ---------------------------------------------------------------------------
R["externe"] = {
    "taux_directeur": [["01/01/2020", 2.50], ["24/06/2020", 2.00], ["16/06/2022", 2.25], ["16/09/2022", 2.50],
                       ["16/12/2022", 2.75], ["16/03/2023", 3.00], ["16/09/2023", 3.25], ["16/12/2023", 3.50],
                       ["16/06/2025", 3.25]],
    "inflation": {"2020": 2.4, "2021": 4.2, "2022": 5.2, "2023": 4.4, "2024": 3.5},
    "taux_technique_cima": 3.5}
# Rendement apparent des DAT et du FCP (produits financiers / placements moyens)
rend = []
for ex in EX[1:-1]:
    prod_fin = -solde(ex, r"^77")
    plac = (solde(ex, r"^52[2-6]|^533") + solde(ex - 1, r"^52[2-6]|^533")) / 2
    rend.append({"exercice": ex, "produits_financiers": prod_fin, "placements_moyens": plac,
                 "rendement": prod_fin / plac if plac else None})
T["7_rendement_placements"] = pd.DataFrame(rend)
R["rendement_placements"] = pd.DataFrame(rend).replace({np.nan: None}).to_dict("records")

# ---------------------------------------------------------------------------
# Écriture
# ---------------------------------------------------------------------------
def conv(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return None if np.isnan(o) else float(o)
    if isinstance(o, (pd.Timestamp,)):
        return o.strftime("%d/%m/%Y")
    return str(o)


json.dump(R, open(os.path.join(SORTIE, "resultats.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1,
          default=conv)
with pd.ExcelWriter(os.path.join(SORTIE, "FPPN_Calculs_rapport.xlsx")) as w:
    for k, v in T.items():
        v.to_excel(w, sheet_name=k[:31], index=False)
print("Tableaux :", len(T))
