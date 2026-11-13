# -*- coding: utf-8 -*-
"""
FPPN - Écritures des comptes d'épargne (41195100 PEP, 41195200 EFC) lues dans les grands livres des comptes, 2020 au 31/08/2026.
Une ligne par écriture : exercice, compte, date, journal, pièce, libellé, référence, débit, crédit.
Les colonnes débit et crédit sont repérées par l'en-tête « Mouvement débit / crédit » de chaque fichier.
Utilisation : python gl_epargne_fppn.py <dossier « Balance générale et Grand livre »> <sortie.parquet>
"""
import os, re, sys
import pandas as pd
from python_calamine import CalamineWorkbook

G, SORTIE = sys.argv[1:3]
FICHIERS = {2020: "2020/GRAND LIVRE DU 01-09-2020 AU 31-12-2020.xlsx",
            2021: "Nouveaux fichiers/GRAND LIVRE 2021.xlsx",
            2022: "2022/Nouveau/GRAND LIVRE DES COMPTES 2022.xlsx",
            2023: "2023/GRAND LIVRE DES COMPTES FPPN 2023.xlsx",
            2024: "Nouveaux fichiers/GRAND LIVRE FPPN 2024.xlsx",
            2025: "2025/GRAND LIVRE DES COMPTES FPPN 2025.xlsx",
            2026: "2026/GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx"}
COMPTES = ("41195100", "41195200")
out = []
for an, f in FICHIERS.items():
    w = CalamineWorkbook.from_path(os.path.join(G, f))
    rows = w.get_sheet_by_name(w.sheet_names[0]).to_python()
    hd = hc = hs = None
    compte = None
    for r in rows:
        txt = [str(c) for c in r]
        if hd is None and any("bit" in t and "Mouvement" in t for t in txt):
            hd = next(j for j, t in enumerate(txt) if "Mouvement" in t and "bit" in t)
            hc = next(j for j, t in enumerate(txt) if "Mouvement" in t and "dit" in t and j != hd)
            hs = next(j for j, t in enumerate(txt) if "Solde" in t)
            continue
        c0 = txt[0].strip() if r else ""
        if re.fullmatch(r"\d{8}(\.0)?", c0) and not re.match(r"\d{4}-\d{2}-\d{2}", c0):
            compte = c0[:8]
            continue
        if any("Total compte" in t for t in txt):
            compte = None
            continue
        if compte not in COMPTES or not re.match(r"\d{4}-\d{2}-\d{2}", c0):
            continue
        num = lambda a, b: sum(float(c) for c in r[a:b] if isinstance(c, (int, float)))
        textes = [c for c in r[3:hd] if isinstance(c, str) and c.strip()]
        out.append({"exercice": an, "compte": compte, "date": c0[:10], "journal": txt[1].strip(), "piece": txt[2].strip(),
                    "libelle": textes[0].strip() if textes else "", "reference": textes[1].strip() if len(textes) > 1 else "",
                    "debit": num(hd, hc), "credit": num(hc, hs), "fichier": os.path.basename(f)})
df = pd.DataFrame(out)
df.to_parquet(SORTIE, index=False)
print(df.groupby(["exercice", "compte", df.journal.eq("RAN")]).agg(lignes=("debit", "size"), debit=("debit", "sum"), credit=("credit", "sum")).to_string())
