# -*- coding: utf-8 -*-
"""
FPPN - Qualité des bases reçues, variable par variable (complétude, cohérence des dates et des montants)
========================================================================================================
Bases contrôlées : lignes retenues des fichiers de précomptes (Solde, CGRAE) et prêts du Grenier au 31/08/2026.
Utilisation : python qualite_bases_fppn.py <dossier_base> <sortie.json>
"""
import json
import os
import sys

import pandas as pd

BASE, SORTIE = sys.argv[1:3]
OUT = {"completude": [], "coherence": []}
pre = pd.read_parquet(os.path.join(BASE, "precomptes.parquet"))
pre = pre[pre.retenu & (pre.annee >= 2020) & ((pre.annee < 2026) | (pre.mois <= 8))]
LIB = {("936", "retour"): "Précomptes maladie (936)", ("8Z6", "retour"): "Précomptes décès (8Z6)", ("998", "retour"): "Versements PEP (998)",
       ("870", "retour"): "Versements EFC (870)", ("9X4", "retour"): "Remboursements Grenier (9X4)", ("CGRAE", "retour"): "Retraités (CGRAE)"}


def vide(s):
    return s.isna() | (s.astype(str).str.strip().isin(["", "nan", "None", "NaT"]))


for (po, se), lib in LIB.items():
    g = pre[(pre.poste == po) & (pre.sens == se)]
    n = len(g)
    for var, mask in (("Matricule", vide(g.matricule)), ("Nom", vide(g.nom)), ("Montant", g.montant.isna()), ("Grade", vide(g.grade)), ("Date de naissance", g.date_naissance.isna())):
        OUT["completude"].append({"base": lib, "variable": var, "lignes": int(n), "manquantes": int(mask.sum())})
    OUT["coherence"].append({"base": lib, "controle": "Montant nul ou négatif", "lignes": int(n), "anomalies": int((g.montant <= 0).sum())})
    d = g[~vide(g.matricule)].duplicated(subset=["matricule", "annee", "mois"], keep=False).sum()
    OUT["coherence"].append({"base": lib, "controle": "Même matricule plusieurs fois dans le même mois", "lignes": int(n), "anomalies": int(d)})

gb = pd.read_parquet(os.path.join(BASE, "grenier_balances.parquet"))
g = gb[(gb.portefeuille == "global") & (gb.arrete == "2026-08-31")]
n = len(g)
lib = "Prêts du Grenier au 31/08/2026"
for var, mask in (("Date de déblocage", g.date_deblocage.isna()), ("Date de première échéance", g.date_premiere_echeance.isna()), ("Date de dernière échéance", g.date_derniere_echeance.isna()),
                  ("Taux", g.taux.isna() | (g.taux <= 0)), ("Durée", g.duree_mois.isna() | (g.duree_mois <= 0)), ("Mécano de l'emprunteur", vide(g.mecano))):
    OUT["completude"].append({"base": lib, "variable": var, "lignes": int(n), "manquantes": int(mask.sum())})
for ctrl, mask in (("Première échéance avant le déblocage", g.date_premiere_echeance < g.date_deblocage),
                   ("Dernière échéance avant la première", g.date_derniere_echeance < g.date_premiere_echeance),
                   ("Encours supérieur au montant débloqué", g.encours > g.montant_debloque + 0.5),
                   ("Encours négatif", g.encours < 0), ("Retard négatif", g.retard < 0),
                   ("Déblocage après la date de l'arrêté", g.date_deblocage > g.arrete)):
    OUT["coherence"].append({"base": lib, "controle": ctrl, "lignes": int(n), "anomalies": int(mask.fillna(False).sum())})
json.dump(OUT, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k in OUT:
    for x in OUT[k]:
        print(k, x)
