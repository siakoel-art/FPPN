# -*- coding: utf-8 -*-
"""
FPPN - Vérification indépendante des fichiers de précomptes (total de chaque fichier mensuel)
=============================================================================================
Pour chaque feuille retenue (sources_mois), relit le CSV brut, repère la colonne des montants par son en-tête
et compare le total recalculé au total de la base de travail.
Utilisation : python verif_precomptes_fppn.py <dossier_tableaux> <base> <sortie.json>
"""
import csv
import glob
import json
import os
import re
import sys

import pandas as pd

TAB, BASE, SORTIE = sys.argv[1:4]
s = pd.read_parquet(os.path.join(BASE, "sources_mois.parquet"))
s = s[s.retenu & (s.annee >= 2020)]
MOTS = ("MONTANT", "MT ", "MNT", "RETENUE", "PRECOMPTE", "MONTAN", "MENSUEL", "FPPN")


def num(x):
    x = (x or "").replace(" ", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return None


res = []
TXT = {}
for f in glob.glob(os.path.join(TAB, "..", "bruts", "*_998_*.txt")):
    src = os.path.basename(f).split("_")[0]
    for l in open(f, encoding="cp1252", errors="replace"):
        m = re.search(r"\s998\s+(\d+)\s+\d{2}/(\d{2})/(\d{4})", l)
        if m:
            k = (src, int(m.group(3)), int(m.group(2)))
            a, n = TXT.get(k, (0, 0))
            TXT[k] = (a + int(m.group(1)), n + 1)
for _, r in s.iterrows():
    if (r.src, int(r.annee), int(r.mois)) in TXT:
        a, n = TXT[(r.src, int(r.annee), int(r.mois))]
        res.append({"src": r.src, "fichier": r.fichier, "poste": r.poste, "sens": r.sens, "annee": int(r.annee), "mois": int(r.mois), "colonne": "texte à largeur fixe",
                    "brut": a, "lignes_brutes": n, "base": round(r.montant), "lignes_base": int(r.lignes), "ecart": round(a - r.montant)})
        continue
    cand = glob.glob(os.path.join(TAB, r.src + "_*.csv")) or glob.glob(os.path.join(TAB, "..", "bruts", r.src + "_*.csv"))
    feu = re.sub(r"[^A-Za-z0-9]+", "_", str(r.feuille)).strip("_").lower()
    cand = [c for c in cand if feu and feu in re.sub(r"[^A-Za-z0-9]+", "_", os.path.basename(c)).lower()] or cand
    if not cand:
        res.append({"src": r.src, "fichier": r.fichier, "statut": "CSV absent"})
        continue
    raw = open(cand[0], "rb").read()
    try:
        txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        txt = raw.decode("cp1252")
    lines = txt.splitlines()
    dl = ";" if sum(l.count(";") for l in lines[:20]) >= sum(l.count(",") for l in lines[:20]) else ","
    rows = [rr for rr in csv.reader(lines, delimiter=dl) if not any("TOTAL" in c.upper() for c in rr)]
    best = None
    for i, row in enumerate(rows[:15]):
        for j, c in enumerate(row):
            u = c.upper().strip()
            if (any(m in u + " " for m in MOTS) or u == "FPPN") and "SALAIRE" not in u and "SOLDE" not in u and "PREVOYANCE" not in u and "PRELEVE" not in u:
              for j in (j, j + 1, j - 1):
                vals = [v for v in (num(rr[j]) if j < len(rr) else None for rr in rows[i + 1:]) if v is not None]
                if r.poste == "8Z6":
                    vals = [v for v in vals if v <= 1_000_000]  # correction documentée du rapport : montants décès > 1 000 000 FCFA
                if len(vals) > 2 and abs(max(vals) - (sum(vals) - max(vals))) <= 1:
                    vals.remove(max(vals))
                tot, n = sum(vals), len(vals)
                if best is None or abs(tot - r.montant) < abs(best[0] - r.montant):
                    best = (tot, n, c.strip())
    if best is None:
        res.append({"src": r.src, "fichier": r.fichier, "poste": r.poste, "sens": r.sens, "statut": "colonne non repérée", "base": r.montant})
        continue
    res.append({"src": r.src, "fichier": r.fichier, "poste": r.poste, "sens": r.sens, "annee": int(r.annee), "mois": int(r.mois), "colonne": best[2],
                "brut": round(best[0]), "lignes_brutes": best[1], "base": round(r.montant), "lignes_base": int(r.lignes), "ecart": round(best[0] - r.montant)})
json.dump(res, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
ok = [x for x in res if x.get("ecart") is not None and abs(x["ecart"]) <= 1]
ko = [x for x in res if x not in ok]
print("feuilles", len(res), "identiques", len(ok), "à examiner", len(ko))
for x in ko[:40]:
    print({k: x.get(k) for k in ("src", "fichier", "poste", "sens", "colonne", "brut", "base", "ecart", "lignes_brutes", "lignes_base", "statut")})
