# -*- coding: utf-8 -*-
"""
FPPN - Vérification indépendante des chiffres du rapport à partir des fichiers du client
========================================================================================
Relit chaque balance directement dans sa conversion brute (une feuille = un CSV, sans la base de travail),
recalcule les agrégats avec les définitions du classeur, et compare :
  A. balance brute et base de travail, compte par compte (soldes débit et crédit) ;
  B. équilibre de chaque balance (total des soldes débiteurs = total des soldes créditeurs) ;
  C. agrégats recalculés et agrégats publiés (ratios.json, utilisés par le rapport et le classeur) ;
  D. balances renvoyées le 22/10/2026 et balances utilisées.
Utilisation : python verif_sources_fppn.py <dossier_tableaux> <base> <ratios.json> <dossier_lot1_balances> <sortie.json>
"""
import csv
import glob
import json
import os
import re
import sys

import pandas as pd

TAB, BASE, RATIOS, LOT1, SORTIE = sys.argv[1:6]
FICH = {2020: "F00002", 2021: "F00010", 2022: "F00018", 2023: "F00026", 2024: "F00034", 2025: "F00042", 2026: "F00050"}
OUT = {"A_comptes": [], "B_equilibre": [], "C_agregats": [], "D_renvois": []}


def num(x):
    x = (x or "").replace(" ", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return None


def lire_csv(path):
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        return list(csv.reader(f, delimiter=";"))


def balance_brute(rows):
    """Comptes à 8 chiffres -> (solde débit, solde crédit), d'après les deux dernières colonnes Débit / Crédit de l'en-tête."""
    pos = None
    for r in rows[:40]:
        lab = [i for i, c in enumerate(r) if c.strip() in ("Débit", "Crédit")]
        if len(lab) >= 2:
            pos = lab
    sd, sc = pos[-2], pos[-1]
    res = {}
    for r in rows:
        if r and re.fullmatch(r"\d{8}", r[0].strip()):
            vals = [(i, num(c)) for i, c in enumerate(r) if num(c) is not None and i > 0]
            d = sum(v for i, v in vals if 0 <= i - sd <= 1)
            c = sum(v for i, v in vals if 0 <= i - sc <= 1)
            res[r[0].strip()] = (d, c)
    return res


def balance_xlsx(path):
    df = pd.read_excel(path, header=None, dtype=str).fillna("")
    return balance_brute(df.values.tolist())


base = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
base = base[base.nature == "comptes"]
BR = {}
for an, f in FICH.items():
    path = glob.glob(os.path.join(TAB, f + "_*.csv"))[0]
    bb = balance_brute(lire_csv(path))
    BR[an] = bb
    g = base[base.exercice == an].set_index("compte")
    ecarts = []
    for k in sorted(set(bb) | set(g.index)):
        d1, c1 = bb.get(k, (0.0, 0.0))
        d2 = float(g.sol_d.get(k, 0) or 0) if k in g.index else 0.0
        c2 = float(g.sol_c.get(k, 0) or 0) if k in g.index else 0.0
        if abs((d1 - c1) - (d2 - c2)) > 0.5:
            ecarts.append({"compte": k, "brut": d1 - c1, "base": d2 - c2})
    OUT["A_comptes"].append({"exercice": an, "fichier": os.path.basename(path), "comptes_bruts": len(bb), "comptes_base": int(len(g)), "ecarts": ecarts})
    td, tc = sum(v[0] for v in bb.values()), sum(v[1] for v in bb.values())
    OUT["B_equilibre"].append({"exercice": an, "soldes_debiteurs": round(td), "soldes_crediteurs": round(tc), "ecart": round(td - tc)})


def solde(an, prefixes, mode="S"):
    t = 0.0
    for k, (d, c) in BR[an].items():
        for p in prefixes:
            ok = (k == p[1:]) if p.startswith("=") else k.startswith(p)
            if ok:
                s = d - c
                t += s if mode == "S" else (max(s, 0) if mode == "D" else max(-s, 0))
                break
    return t


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from classeur_formules_fppn import AGG  # noqa: E402  (définitions des agrégats, mêmes que le classeur)

R = json.load(open(RATIOS, encoding="utf-8"))["agregats"]
for an in FICH:
    val = {}
    for cle, lib, defi in AGG:
        if isinstance(defi, list) and all(m in ("S", "D", "C") for _, m, _ in defi):
            val[cle] = sum(sg * solde(an, pf, m) for sg, m, pf in defi)
    for cle in ("cotisations_acquises", "charges_prestations", "cot_maladie", "cot_deces", "prest_maladie", "prest_deces", "epargne_pep", "epargne_efc",
                "provisions_techniques", "fonds_mutualistes_10", "tresorerie_actif", "disponibilites", "placements_tresorerie", "creance_hopital", "cmp", "grenier",
                "produits_7", "charges_6", "charges_fonctionnement", "charges_personnel", "provision_196", "ran_debiteur_129", "reserve_disponible_106"):
        k2 = {"cot_maladie": "cot_maladie_7010", "cot_deces": "cot_deces_7021", "provision_196": "provision_196"}.get(cle, cle)
        if k2 in val and cle in R[str(an)]:
            pub = R[str(an)][cle]
            OUT["C_agregats"].append({"exercice": an, "agregat": cle, "recalcul_brut": round(val[k2]), "publie": round(pub), "ecart": round(val[k2] - pub)})

for nom, an in (("BALANCE 2021.xlsx", 2021), ("BALANCE FPPN 2025.xlsx", 2025)):
    p = os.path.join(LOT1, nom)
    if os.path.exists(p):
        bb = balance_xlsx(p)
        e = [k for k in set(bb) | set(BR[an]) if abs((bb.get(k, (0, 0))[0] - bb.get(k, (0, 0))[1]) - (BR[an].get(k, (0, 0))[0] - BR[an].get(k, (0, 0))[1])) > 0.5]
        OUT["D_renvois"].append({"fichier": nom, "exercice": an, "comptes": len(bb), "comptes_en_ecart": sorted(e)[:30], "nb_ecarts": len(e)})
json.dump(OUT, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for x in OUT["A_comptes"]:
    print("A", x["exercice"], x["comptes_bruts"], x["comptes_base"], "écarts", len(x["ecarts"]), x["ecarts"][:3])
for x in OUT["B_equilibre"]:
    print("B", x)
print("C agrégats", len(OUT["C_agregats"]), "en écart", sum(1 for x in OUT["C_agregats"] if abs(x["ecart"]) > 1))
for x in OUT["C_agregats"]:
    if abs(x["ecart"]) > 1:
        print("  C", x)
for x in OUT["D_renvois"]:
    print("D", x)
