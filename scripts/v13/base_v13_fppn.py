# -*- coding: utf-8 -*-
"""
FPPN v13 (13/11/2026) : base retraitée de l'extourne manquante de la provision de la clinique.
Constat : grand livre 2025, OD 805, 806 et 807 du 31/12/2025, « CMP PROV FACT 10/2025 » à « 12/2025 », 3 × 85 000 000 FCFA
au débit du 60120000 et au crédit du 40110000. Grand livre 2026 : aucune extourne (aucun crédit en classe 60, aucune écriture
« PROV » au 40110000), et les factures réelles d'octobre à décembre 2025 (SIN 462, 552, 619 du 31/01/2026, 240 471 119 FCFA)
sont passées en charges 2026.
Retraitement : l'écriture d'extourne du 01/01/2026 est ajoutée à la balance 2026 : crédit 60120000 et débit 40110000 de
255 000 000 FCFA. Les autres exercices et les autres comptes ne changent pas.
Utilisation : python base_v13_fppn.py <dossier_base> <dossier_base_v13>
"""
import os
import shutil
import sys

import pandas as pd

SRC, DST = sys.argv[1:3]
EXTOURNE = 255000000.0
os.makedirs(DST, exist_ok=True)
for f in os.listdir(SRC):
    if f != "compta_balances.parquet":
        shutil.copy2(os.path.join(SRC, f), os.path.join(DST, f))
b = pd.read_parquet(os.path.join(SRC, "compta_balances.parquet"))
m26 = (b.nature == "comptes") & (b.exercice == 2026)
for compte, col in (("60120000", "mvt_c"), ("40110000", "mvt_d")):
    k = m26 & (b.compte == compte)
    assert k.sum() == 1, compte
    b.loc[k, col] += EXTOURNE
    net = b.loc[k, "ouv_d"] + b.loc[k, "mvt_d"] - b.loc[k, "ouv_c"] - b.loc[k, "mvt_c"]
    b.loc[k, "sol_d"] = net.clip(lower=0)
    b.loc[k, "sol_c"] = (-net).clip(lower=0)
    print(compte, b.loc[k, ["mvt_d", "mvt_c", "sol_d", "sol_c"]].to_dict("records"))
b.to_parquet(os.path.join(DST, "compta_balances.parquet"), index=False)
