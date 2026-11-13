# -*- coding: utf-8 -*-
"""
FPPN - Règle réellement appliquée aux intérêts du PEP à la sortie (13/11/2026)
=============================================================================
Suite de test_interets_epargne_fppn.py, après contrôle d'un échantillon de 20 dossiers : le montant remboursé est
égal au total versé multiplié par 1,035, quelle que soit la durée d'épargne. Ce script le vérifie de deux façons.

A. Sans rapprochement par le nom (toutes les écritures) : un remboursement PEP (compte 41195100, libellé REMB, REGLT,
   SOLDE ou RESIL, hors acomptes, rachats partiels et trop-perçus) « porte 3,5 % » si montant / 1,035 est un multiple
   de 5 000 FCFA à 1 FCFA près (97 % des précomptes PEP sont des multiples de 5 000 FCFA). Intérêt payé = montant - montant / 1,035.
   Contrôle : la même règle est testée avec 1 (aucun intérêt) pour les acomptes et rachats partiels.
B. Sur les policiers sortis avec un historique complet (test_interets_epargne_fppn.py) : base retenue par le FPPN
   = reçu / 1,035, comparée au total versé (précomptes 998 et dépôts en espèces).
   Intérêts promis : 3,5 % « tous les ans » (« MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx », partie PEP),
   intérêt simple de la fin du mois de chaque précompte à la date de sortie.

Utilisation : python regle_interets_pep_fppn.py <gl_epargne.parquet> <detail_nominatif.parquet> <dossier_sortie>
Sorties : regle_interets_pep.json et regle_interets_pep.xlsx (agrégés, sans nom ni matricule).
"""
import json, os, sys
import numpy as np
import pandas as pd

GLP, DET, SORTIE = sys.argv[1:4]


def porte(montant, k, pas=5000):
    b = montant / k
    return (b - (b / pas).round() * pas).abs() <= 1


g = pd.read_parquet(GLP)
g = g[(g.journal != "RAN") & (g.debit > 0) & (g.compte == "41195100")].copy()
L = g.libelle.str.upper()
g["type"] = np.select([L.str.contains("TROP"), L.str.contains("RACHAT"), L.str.contains("ACP|ACOMPTE|AC "),
                       L.str.contains("REMB|REGLT|RGLT|SOLDE|RESIL")],
                      ["trop-perçu", "rachat partiel", "acompte", "remboursement"], default="autre")
g["x_1_035"] = porte(g.debit, 1.035)
g["x_1"] = porte(g.debit, 1.0)
g["interet_paye"] = np.where((g.type == "remboursement") & g.x_1_035, g.debit - g.debit / 1.035, 0.0)
A_type = g.groupby("type").agg(ecritures=("debit", "size"), montant=("debit", "sum"),
                               part_x_1_035=("x_1_035", "mean"), part_sans_interet=("x_1", "mean")).reset_index()
r = g[g.type == "remboursement"]
A_an = r.groupby("exercice").agg(remboursements=("debit", "size"), montant=("debit", "sum"), avec_3_5=("x_1_035", "sum"),
                                 interet_paye=("interet_paye", "sum")).reset_index()

d = pd.read_parquet(DET)
c = d[(d.produit == "PEP") & (d.statut == "historique complet")].copy()
c["base_fppn"] = c.recu / 1.035
c["diff"] = c.verse - c.base_fppn
mensuel = c.verse / c.mois_precomptes
c["cas"] = np.select([c["diff"].abs() <= 5, (c["diff"] > 0) & (c["diff"] <= 3.5 * mensuel), c["diff"] > 3.5 * mensuel],
                     ["reçu = versé x 1,035", "versé non repris : 1 à 3 mois", "versé non repris : plus de 3 mois"],
                     default="reçu > versé x 1,035")
c["annees"] = (pd.to_datetime(c.date_sortie) - pd.to_datetime(c.premier_precompte.str[-4:] + "-" + c.premier_precompte.str[:2] + "-01")).dt.days / 365.25
B = c.groupby("cas").agg(policiers=("recu", "size"), recu=("recu", "sum"), verse=("verse", "sum"),
                         interets_promis=("interets_promis", "sum"), duree_mediane_ans=("annees", "median")).reset_index()
B["paye_en_plus"] = B.recu - B.verse
flat = c[c.cas == "reçu = versé x 1,035"]
OUT = {"A_par_type": A_type.to_dict("records"), "A_remboursements_par_exercice": A_an.to_dict("records"),
       "A_total": {"remboursements": int(len(r)), "montant": float(r.debit.sum()), "avec_3_5": int(r.x_1_035.sum()),
                   "interet_paye": float(r.interet_paye.sum())},
       "B_par_cas": B.to_dict("records"),
       "B_regle_1_035": {"policiers": int(len(flat)), "verse": float(flat.verse.sum()), "paye_en_plus": float(flat.recu.sum() - flat.verse.sum()),
                         "interets_promis": float(flat.interets_promis.sum()),
                         "manque": float(flat.interets_promis.sum() - (flat.recu.sum() - flat.verse.sum())),
                         "duree_ans_quartiles": flat.annees.quantile([0.1, 0.25, 0.5, 0.75, 0.9]).round(2).to_dict()},
       "B_non_repris_1_3_mois": float(c.loc[c.cas == "versé non repris : 1 à 3 mois", "diff"].sum())}
json.dump(OUT, open(os.path.join(SORTIE, "regle_interets_pep.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=float)
with pd.ExcelWriter(os.path.join(SORTIE, "regle_interets_pep.xlsx")) as xw:
    A_type.to_excel(xw, sheet_name="A_par_type", index=False)
    A_an.to_excel(xw, sheet_name="A_par_exercice", index=False)
    B.to_excel(xw, sheet_name="B_dossiers_complets", index=False)
pd.set_option("display.width", 250)
print(A_type.round(3).to_string()); print(A_an.round(0).to_string()); print(B.round(2).to_string())
print({k: v for k, v in OUT.items() if k in ("A_total", "B_regle_1_035", "B_non_repris_1_3_mois")})
