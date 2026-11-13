# -*- coding: utf-8 -*-
"""
FPPN - Vérification, sur les données, des constats de la revue indépendante du rapport du 11/11/2026 (13/11/2026)
================================================================================================================
Chaque fonction recalcule un chiffre cité par un réviseur (constats_A.md à constats_E.md) à partir :
- des balances (base/compta_balances.parquet, nature « comptes », solde = sol_d - sol_c) ;
- des grands livres des comptes 2023 au 31/08/2026 (lus ici, journal « RAN » exclu) ;
- du classeur du 11/11/2026 recalculé par Excel ;
- du texte du rapport du 11/11/2026 (une entrée par diapositive).
Utilisation : python verif_revue_fppn.py <dossier_base> <classeur_recalcule.xlsx> <rapport_texte.json> <dossier grands livres> <sortie.json>
"""
import json, math, os, re, sys
import numpy as np
import pandas as pd
from python_calamine import CalamineWorkbook

BASE, CLASSEUR, TEXTE, GLDIR, SORTIE = sys.argv[1:6]
bal = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
bal = bal[bal.nature == "comptes"].copy()
bal["solde"] = bal.sol_d - bal.sol_c
TX = {d["diapo"]: d["texte"] for d in json.load(open(TEXTE, encoding="utf-8"))}
R = {}


def S(ex, *pref, excl=(), col="solde"):
    b = bal[bal.exercice == ex]
    m = b.compte.str.startswith(tuple(pref)) & ~b.compte.str.startswith(tuple(excl)) if excl else b.compte.str.startswith(tuple(pref))
    return float(b.loc[m, col].sum())


def dans(diapos, motif):
    """Diapositives (parmi celles citées) dont le texte contient le motif (regex)."""
    return [d for d in diapos if re.search(motif, TX.get(d, ""), re.I)]


def feuille(nom):
    w = CalamineWorkbook.from_path(CLASSEUR)
    return w.get_sheet_by_name(nom).to_python()


def pct(a, b):
    return round(100 * a / b, 2)


# ---------------------------------------------------------------- grands livres (classes 6 et 7, et 4011)
FGL = {2023: "2023/GRAND LIVRE DES COMPTES FPPN 2023.xlsx", 2024: "Nouveaux fichiers/GRAND LIVRE FPPN 2024.xlsx",
       2025: "2025/GRAND LIVRE DES COMPTES FPPN 2025.xlsx", 2026: "2026/GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx"}
lignes = []
for an, f in FGL.items():
    w = CalamineWorkbook.from_path(os.path.join(GLDIR, f))
    rows = w.get_sheet_by_name(w.sheet_names[0]).to_python()
    hd = hc = hs = None
    cpt = None
    for r in rows:
        t = [str(c) for c in r]
        if hd is None and any("Mouvement" in x and "bit" in x for x in t):
            hd = next(j for j, x in enumerate(t) if "Mouvement" in x and "bit" in x)
            hc = next(j for j, x in enumerate(t) if "Mouvement" in x and "dit" in x and j != hd)
            hs = next(j for j, x in enumerate(t) if "Solde" in x)
            continue
        c0 = t[0].strip()
        if re.fullmatch(r"\d{8}(\.0)?", c0):
            cpt = c0[:8]
            continue
        if any("Total compte" in x for x in t):
            cpt = None
            continue
        if not cpt or not re.match(r"\d{4}-\d{2}-\d{2}", c0) or t[1].strip() == "RAN" or not cpt.startswith(("6", "7", "4011")):
            continue
        txt = [c for c in r[3:hd] if isinstance(c, str) and c.strip()]
        num = lambda a, b: sum(float(c) for c in r[a:b] if isinstance(c, (int, float)))
        lignes.append({"exercice": an, "compte": cpt, "date": c0[:10], "journal": t[1].strip(), "piece": t[2].strip(),
                       "libelle": txt[0].strip() if txt else "", "debit": num(hd, hc), "credit": num(hc, hs)})
GL = pd.DataFrame(lignes)

# ================================================================ G1 Adossement, provisions dans le passif à un an
c4 = bal[(bal.exercice == 2026) & bal.compte.str.startswith("4")]
dettes4 = float((c4.sol_c - c4.sol_d).clip(lower=0).sum())
epargne26 = -S(2026, "41195100") - S(2026, "41195200")
prov26 = -S(2026, "15801", "15802")
R["G1_passif_un_an"] = {"dettes_classe4_crediteurs": dettes4, "epargne": epargne26, "dettes_hors_epargne": dettes4 - epargne26,
                        "provisions_15801_15802": prov26, "passif_un_an": dettes4 - epargne26 + prov26}
a1, sorties1 = 3457866951, 3028376687.4883485
besoin = 19581392
R["G1_couverture"] = {"rapport": pct(a1, dettes4 - epargne26 + prov26 + sorties1),
                      "provisions_au_besoin": pct(a1, dettes4 - epargne26 + besoin + sorties1),
                      "scenario4_sans_epargne": pct(a1, dettes4 - epargne26 + besoin),
                      "manque_5_ans_corrige": -12596052822.432278 + (prov26 - besoin),
                      "ecart_au_dela_5_ans_corrige": -835832022 + (prov26 - besoin)}
eng = 17931954598 + prov26
R["G1_engagements_D_B2"] = {"engagements_rapport": eng, "couverture_rapport": pct(21925420866, eng),
                            "avec_besoin_et_interets": pct(21925420866, eng - (prov26 - besoin) + 3278913238),
                            "avec_besoin_sans_interets": pct(21925420866, eng - (prov26 - besoin))}
R["G1_pont_fonds_propres"] = {"fonds_propres_2026": S(2026, "1", excl=("15", "16", "17", "18", "19")) * -1,
                              "A": -879498700 + 4809716898 - 3278913238 - 14016392 + 26482228,
                              "D": -879498700 + 4809716898 - 3278913238 - 14016392}
R["G1_diapo6_deux_excedents"] = {"4327397229_dans": dans(range(3, 28), r"4 327 397 229"), "4809716898_dans": dans(range(3, 28), r"4 809 716 898")}
R["G1_borne_favorable"] = dans(range(1, 301), r"borne favorable")

# ================================================================ G2 Intérêts dus
ep = {}
for ex in range(2020, 2027):
    ep[ex] = {"pep_d": -float(bal[(bal.exercice == ex) & (bal.compte == "41195100")].eval("ouv_d-ouv_c").sum()),
              "pep_f": -S(ex, "41195100"),
              "efc_d": -float(bal[(bal.exercice == ex) & (bal.compte == "41195200")].eval("ouv_d-ouv_c").sum()),
              "efc_f": -S(ex, "41195200")}


def interets(t_pep, t_efc, prorata2020=1.0, capitalise=False):
    tot = 0.0
    for ex in range(2020, 2027):
        e = ep[ex]
        f = 8 / 12 if ex == 2026 else (prorata2020 if ex == 2020 else 1.0)
        base_p = (e["pep_d"] + e["pep_f"]) / 2 + (tot if capitalise else 0) * 0
        i = base_p * t_pep(ex) * f + (e["efc_d"] + e["efc_f"]) / 2 * t_efc(ex) * f
        if capitalise:
            i += tot * ((t_pep(ex) + t_efc(ex)) / 2) * f
        tot += i
    return round(tot)


R["G2_interets"] = {"rapport_3_5_et_5_5": interets(lambda a: .035, lambda a: .055),
                    "efc_3_5_avant_2024": interets(lambda a: .035, lambda a: .055 if a >= 2024 else .035),
                    "efc_0_avant_2024": interets(lambda a: .035, lambda a: .055 if a >= 2024 else 0),
                    "tout_3_5": interets(lambda a: .035, lambda a: .035),
                    "rapport_2020_sur_4_mois": interets(lambda a: .035, lambda a: .055, prorata2020=4 / 12),
                    "soldes": ep}
R["G2_interets_2025_base_fin_vs_moyenne"] = {"fin": round(ep[2025]["efc_f"] * .055), "moyenne": round((ep[2025]["efc_d"] + ep[2025]["efc_f"]) / 2 * .055)}
pf = {ex: -S(ex, "77") for ex in range(2020, 2027)}
R["G2_produits_financiers_77"] = pf
R["G2_16_1_fois"] = round(2804987325 / sum(pf[a] for a in range(2021, 2026)), 2)
tp = json.load(open(os.path.join(os.path.dirname(BASE), "test_interets", "regle_interets_pep.json"), encoding="utf-8"))
R["G2_test_pep"] = {"remboursements": tp["A_total"], "dossiers": tp["B_regle_1_035"]}

# ================================================================ G3 Ratio de 16,28 %
res = {c: -S(2026, c) * -1 for c in ("52110200", "52110300", "52210200", "53300001")}
R["G3_couverture"] = {"actifs_reserves": 2562885661, "epargne": epargne26, "ratio": pct(2562885661, epargne26),
                      "fcp_53300001": S(2026, "53300001"), "sans_fcp": pct(2562885661 - S(2026, "53300001"), epargne26)}

# ================================================================ G4 Double comptage des factures de la clinique en 2026
g25 = GL[(GL.exercice == 2025) & GL.compte.str.startswith("6012") & GL.libelle.str.upper().str.contains("CMP PROV")]
g26 = GL[(GL.exercice == 2026) & GL.compte.str.startswith("601") & GL.piece.isin(["462", "552", "619"])]
ext = GL[(GL.exercice == 2026) & GL.compte.str.startswith("60") & (GL.credit > 0)]
mal26 = S(2026, "601") + S(2026, "60800007")
cot_mal26 = -S(2026, "7010")
r26 = S(2026, "60")
c70_26 = -S(2026, "70")
dbl = float(g26.debit.sum())
R["G4_cmp"] = {"provision_2025_OD": float(g25.debit.sum()), "provision_2025_pieces": g25.piece.tolist(),
               "factures_2026_SIN": dbl, "factures_2026_detail": g26[["piece", "debit"]].to_dict("records"),
               "credits_comptes_60_en_2026": float(ext.credit.sum()),
               "sc_maladie_rapport": pct(mal26, cot_mal26), "sc_maladie_corrige": pct(mal26 - dbl, cot_mal26),
               "reglementaire_rapport": pct(r26, c70_26), "reglementaire_corrige": pct(r26 - dbl, c70_26)}

# ================================================================ G5 Ratio combiné, frais, hors soins
R["G5_ratio_combine"] = {"B_B2_reclasse": pct(1284143358 + 875420236 + 0.705 * 487492003, 2409281030),
                         "A_M1_resultat_reclasse": -237774567 + 487492003 - 343609286,
                         "dotations_68_2025": S(2025, "6812", "6813"), "dotations_68_2026": S(2026, "68"),
                         "frais_formule_rapport_2025": pct(S(2025, "61", "62", "63", "64", "65", "66", excl=("6518",)), -S(2025, "7010") - S(2025, "7021"))}
hs = {}
mot = re.compile(r"FUTURE VIS|GIS PLUS|SACS? KABA|HADJ|CEREMONIE|C[EÉ]R[EÉ]M|LOBBING|MISSION|ARBRE DE NOEL|ECRAN GEANT")
for ex in (2023, 2024, 2025, 2026):
    g = GL[(GL.exercice == ex) & (GL.compte.str.startswith("601") | (GL.compte == "60800007")) & (GL.journal != "SIN")]
    hs[ex] = float(g[g.libelle.str.upper().str.contains(mot)].eval("debit-credit").sum())
R["G5_faux_negatifs_libelles"] = hs
# reproduction du classement de sinistralite_gl_fppn.py (fonction cat), puis règles élargies proposées par le réviseur B (M2)
def cat(l):
    if re.search(r"\bDON\b", l) and re.search(r"SOIN|DIALYSE|DON MED|ANTI-AN|KIT D.ACCOUCH|DON DE SANG|COUVERTURE FRAIS|PHCIE", l): return "soins"
    if re.search(r"\bDON\b|K'DO|CADEAU|PANIER", l): return "hors soins"
    if re.search(r"FUTURE VISI|GADGET|SAC KABA|IMPRIMERIE|IMPRIME|NEW PRINT|CARTE", l): return "hors soins"
    if re.search(r"00444\d{3}", l): return "hors soins"
    if re.search(r"RECLASSEMENT FRAIS DE RECOUVREMENT|PROV CPTE", l): return "hors soins"
    return "soins"


ELARGI = re.compile(r"FUTURE VIS|GIS PLUS|SACS? KABA|HADJ|CEREMONIE|CEREM|LOBBING|ARBRE DE NOEL|ECRAN GEANT|LOCAT")
gp = GL[GL.compte.str[:4].isin(["6011", "6012", "6013", "6018", "6080"])].copy()
gp["m"] = gp.debit - gp.credit
gp["cat"] = gp.libelle.str.upper().map(cat)
fn = gp[(gp.cat == "soins") & (gp.journal != "SIN") & gp.libelle.str.upper().str.contains(ELARGI)]
tot = gp.groupby("exercice").m.sum()
hs0 = gp[gp.cat == "hors soins"].groupby("exercice").m.sum()
R["G5_reproduction_classement"] = {int(a): {"prestations": float(tot[a]), "hors_soins_script": float(hs0.get(a, 0)),
                                           "part_hors_soins_script": pct(hs0.get(a, 0), tot[a]),
                                           "faux_negatifs": float(fn[fn.exercice == a].m.sum()),
                                           "part_hors_soins_elargie": pct(hs0.get(a, 0) + fn[fn.exercice == a].m.sum(), tot[a])}
                                   for a in tot.index}
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sinistralite_gl_fppn.py"), encoding="utf-8").read()
R["G5_regles_script"] = {k: (k in src) for k in ("FUTURE VISI", "FUTURE VIS", "GIS PLUS", "SAC KABA", "SACS", "HADJ", "LOBBING")}
R["G5_dons_6583"] = {ex: S(ex, "6583") for ex in range(2021, 2027)}
R["G5_60800008"] = {ex: S(ex, "60800008") for ex in (2023, 2024)}

# ================================================================ G6 Sinistralité, cotisations, tendances
mal = {ex: S(ex, "601") + S(ex, "60800007") for ex in range(2020, 2027)}
cm = {ex: -S(ex, "7010") for ex in range(2020, 2027)}
cd = {ex: -S(ex, "7021") for ex in range(2020, 2027)}
R["G6_part_maladie"] = {ex: pct(cm[ex], cm[ex] + cd[ex]) for ex in range(2020, 2027)}
R["G6_taux_reglementaire"] = {ex: pct(S(ex, "60"), -S(ex, "70")) for ex in range(2020, 2027)}
R["G6_sc_maladie"] = {ex: pct(mal[ex], cm[ex]) for ex in range(2020, 2027)}
R["G6_tcam_maladie"] = {"2020_2025": round(100 * ((mal[2025] / mal[2020]) ** (1 / 5) - 1), 2),
                        "2021_2025": round(100 * ((mal[2025] / mal[2021]) ** (1 / 4) - 1), 2)}
y = np.array([R["G6_sc_maladie"][a] for a in range(2020, 2026)])
x = np.arange(6)
pente, ordo = np.polyfit(x, y, 1)
resid = y - (pente * x + ordo)
se = math.sqrt((resid ** 2).sum() / 4 / ((x - x.mean()) ** 2).sum())
R["G6_regression_sc_maladie"] = {"pente": round(pente, 2), "t": round(pente / se, 2), "residus": resid.round(2).tolist(),
                                 "pente_2023_2025": round(np.polyfit(x[3:], y[3:], 1)[0], 2)}
R["G6_comptes_2025_2026"] = {c: (S(2025, c), S(2026, c)) for c in ("65214001", "63421000")}
R["G6_6012"] = {c: S(2025, c) for c in ("60120000", "60121000", "60122000")}
p = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"), columns=["poste", "sens", "annee", "montant", "retenu"])
R["G6_precomptes_936_2025"] = float(p[(p.poste == "936") & (p.sens == "retour") & p.retenu & (p.annee == 2025)].montant.sum())

# ================================================================ G7 Provisions maladie, décès, personnel
R["G7_maladie"] = {"un_mois_2025": round(mal[2025] / 12), "besoin_rapport": 12481392}
d19 = feuille("19c_deces_declares_tard")
R["G7_deces_tardifs_etats"] = sorted({str(r[0]) for r in d19 if isinstance(r[0], float)})
cr = bal[(bal.exercice == 2020) & (bal.compte == "40110000")]
R["G7_delai_2020"] = {"credits_40110000_2020": float(cr.mvt_c.sum()), "mois_sur_4": round(1394774547 / (float(cr.mvt_c.sum()) / 4), 2)}
pidr = [r for r in feuille("14b_pidr_individuel") if isinstance(r[0], str) and re.match(r"A2\d-\d+", r[0])]
R["G7_pidr_agents"] = len(pidr)

# ================================================================ G8 Liquidité, marge, fonds propres
R["G8_comptes"] = {c: {ex: S(ex, c) for ex in (2025, 2026)} for c in ("11100000", "12100000", "10620000", "47111103", "31200000", "27110000", "27120000", "48500000", "22900000", "23100000", "196")}
R["G8_montant_3932124730"] = bool(((bal.exercice == 2026) & ((bal.sol_d - 3932124730).abs() < 1 | ((bal.sol_c - 3932124730).abs() < 1))).any())
R["G8_reserve_mois"] = {"annualise": round(820232819 * 12 / (1798855783 * 12 / 8), 2), "non_annualise": round(820232819 * 12 / 1798855783, 2)}
R["G8_hao_2025"] = {c: S(2025, c) for c in ("82200000", "84100000")}
R["G8_70220000_2025"] = S(2025, "70220000")

# ================================================================ G9 Épargne : sorties, versements
R["G9_epargne_2025"] = {"credits": float(bal[(bal.exercice == 2025) & bal.compte.isin(["41195100", "41195200"])].mvt_c.sum()),
                        "debits": float(bal[(bal.exercice == 2025) & bal.compte.isin(["41195100", "41195200"])].mvt_d.sum())}
R["G9_pep_2026"] = {"sorties_8_mois": float(bal[(bal.exercice == 2026) & (bal.compte == "41195100")].mvt_d.sum()), "ouverture": ep[2026]["pep_d"]}
R["G9_annee1_par_produit"] = round(ep[2026]["pep_f"] * 0.3029928321911165 + ep[2026]["efc_f"] * 0.08290359437634029)

# ================================================================ G10 Projection
fr = lambda ex: S(ex, "61", "62", "63", "64", "65", "66", excl=("6518",))
c26 = -S(2026, "7010") - S(2026, "7021")
p26 = mal[2026] + S(2026, "6021", "6024")
R["G10_2026"] = {"cotisations": c26, "prestations": p26, "frais": fr(2026), "resultat": c26 - p26 - fr(2026),
                 "resultat_annualise": round((c26 - p26 - fr(2026)) * 1.5), "frais_annualises": round(fr(2026) * 1.5),
                 "hausse_maladie_2026_annualisee": pct(mal[2026] * 1.5, mal[2025]) - 100}
R["G10_resultat_2025_deux_frais"] = {"61_66_hors_6518": round(-S(2025, "7010") - S(2025, "7021") - S(2025, "70220000") - (mal[2025] + S(2025, "6021", "6024")) - fr(2025)),
                                     "612_613_618_62_68": round(S(2025, "612", "613", "618", "62", "63", "64", "65", "66", "67", "68"))}
R["G10_rendement_77"] = {ex: pf[ex] for ex in range(2021, 2026)}

json.dump(R, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
for k, v in R.items():
    if k not in ("G2_interets",):
        print(k, ":", json.dumps(v, ensure_ascii=False, default=str)[:600])
print("G2_interets :", {k: v for k, v in R["G2_interets"].items() if k != "soldes"})
