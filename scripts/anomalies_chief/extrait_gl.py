# -*- coding: utf-8 -*-
"""Extrait des grands livres (2021 au 31/08/2026) toutes les écritures des comptes de prestations (601..., 602..., 608...)
et du compte 40110000 (factures des cliniques à payer). Une ligne par écriture avec le fichier source.
Utilisation : python extrait_gl.py <dossier 04-Documents reçus du client> <sortie.parquet>"""
import datetime as dt
import os
import re
import sys

import openpyxl
import pandas as pd

R = sys.argv[1]
D = os.path.join(R, "01-download", "AUDIT ADENKA", "DAFC")
N = os.path.join(R, "02-Nouveaux_GL et Balances", "Nouveaux fichiers")
F = {2021: os.path.join(N, "GRAND LIVRE 2021.xlsx"),
     2022: os.path.join(N, "GRAND LIVRE DES COMPTES 2022.xlsx"),
     2023: os.path.join(D, "2023", "GRAND LIVRE DES COMPTES FPPN 2023.xlsx"),
     2024: os.path.join(N, "GRAND LIVRE FPPN 2024.xlsx"),
     2025: os.path.join(D, "2025", "GRAND LIVRE DES COMPTES FPPN 2025.xlsx"),
     2026: os.path.join(D, "2026", "GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx")}
GARDE = re.compile(r"^(60[128]\d{5}|40110000)$")


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


out = []
for an, f in F.items():
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
    cur, entete = None, {}
    for i, l in enumerate(ws.iter_rows(values_only=True), 1):
        if not l:
            continue
        if l[0] == "Date":  # ligne d'en-tête : repère les colonnes
            for j, x in enumerate(l):
                if isinstance(x, str):
                    entete[x.split("\n")[0].strip().lower()] = j
            continue
        c0 = str(l[0]).split(".")[0] if l[0] is not None else ""
        if re.fullmatch(r"\d{8}", c0) and isinstance(l[2] if len(l) > 2 else None, str) and not isinstance(l[1], str):
            cur = c0
            continue
        if re.fullmatch(r"\d{8}", c0) and (len(l) < 2 or l[1] in (None, "")):
            cur = c0
            continue
        if cur and GARDE.match(cur) and isinstance(l[0], (dt.datetime, dt.date)):
            # débit en colonne 20 et crédit en colonne 22 (cellules fusionnées décalées d'une colonne sur l'en-tête), contrôlé avec les balances
            out.append(dict(an=an, fichier=os.path.basename(f), ligne=i, compte=cur, date=l[0].date() if isinstance(l[0], dt.datetime) else l[0],
                            journal=l[1], piece=str(l[2] or "").strip(), libelle=str(l[5] or "").strip(),
                            facture=str(l[12] or "").strip() if len(l) > 12 else "",
                            debit=num(l[19]) if len(l) > 19 else 0.0, credit=num(l[21]) if len(l) > 21 else 0.0))
    print(an, os.path.basename(f), sum(1 for o in out if o["an"] == an))
d = pd.DataFrame(out)
d["montant"] = d.debit - d.credit
d.to_parquet(sys.argv[2], index=False)
print(d.groupby(["an"]).size())
