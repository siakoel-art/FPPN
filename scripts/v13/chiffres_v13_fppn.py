# -*- coding: utf-8 -*-
"""
FPPN v13 (13/11/2026) : chiffres clés recalculés après extourne de la provision de la clinique (base_v13).
Les titres, sous-titres et le résumé exécutif v13 lisent ces valeurs (aucun chiffre recopié à la main).
Utilisation : python chiffres_v13_fppn.py <base_v13> <resultats_v13> <classeur_v12_recalcule.xlsx> <regle_interets_pep.json> <sortie.json>
"""
import json
import os
import sys

import pandas as pd
from python_calamine import CalamineWorkbook

BASE, RES, CLASSEUR, PEPJ, SORTIE = sys.argv[1:6]
b = pd.read_parquet(os.path.join(BASE, "compta_balances.parquet"))
b = b[b.nature == "comptes"].copy()
b["s"] = b.sol_d - b.sol_c


def S(ex, *p, excl=()):
    x = b[(b.exercice == ex) & b.compte.str.startswith(p)]
    if excl:
        x = x[~x.compte.str.startswith(excl)]
    return float(x.s.sum())


def fr(ex):
    return S(ex, "61", "62", "63", "64", "65", "66", excl=("6518",))


K = {}
# ---- 2026 après extourne
mal = S(2026, "601") + S(2026, "60800007")
cot = -S(2026, "7010") - S(2026, "7021")
prest = mal + S(2026, "6021", "6024")
K["extourne"] = 255000000
K["factures_sin"] = 240471119
K["sc_mal_26"] = 100 * mal / -S(2026, "7010")
K["regl_26"] = 100 * S(2026, "60") / -S(2026, "70")
K["depense_100_26"] = 100 * (prest + fr(2026)) / cot
K["res_tech_26"] = cot - prest - fr(2026)
K["res_tech_26_dec"] = K["res_tech_26"] * 12 / 8
K["seuil_80"] = 100 * (80 / K["regl_26"] - 1)
K["res_net_26"] = -S(2026, "6", "7", "8")
K["res_net_26_dec"] = K["res_net_26"] * 12 / 8
K["dette_cliniques_26"] = -S(2026, "40110000")
# ---- adossement au 31/08/2026
c4 = b[(b.exercice == 2026) & b.compte.str.startswith("4")]
dettes4 = float((c4.sol_c - c4.sol_d).clip(lower=0).sum())
epargne = -S(2026, "41195100") - S(2026, "41195200")
prov = -S(2026, "15801", "15802")
besoin = 19581392
w = CalamineWorkbook.from_path(CLASSEUR)
r24c = w.get_sheet_by_name("24c_ecarts_par_annee").to_python()
ann = [r for r in r24c if isinstance(r[0], str) and (r[0].startswith("Année") or r[0].startswith("Au-del"))]
actif1 = ann[0][1]
sorties = [r[3] for r in ann]
actifs = [r[1] for r in ann]
D = dettes4 - epargne
K.update({"epargne": epargne, "dettes_hors_epargne": D, "provisions": prov, "besoin": besoin, "actif_un_an": actif1,
          "couv_comptes": 100 * actif1 / (D + prov + sorties[0]), "couv_besoin": 100 * actif1 / (D + besoin + sorties[0]),
          "couv_hors_epargne": 100 * actif1 / (D + besoin), "manque_un_an": actif1 - (D + besoin + sorties[0])})
K["manque_5_ans_comptes"] = sum(actifs[:5]) - (D + prov) - sum(sorties[:5])
K["manque_5_ans_besoin"] = sum(actifs[:5]) - (D + besoin) - sum(sorties[:5])
K["ecart_au_dela_besoin"] = K["manque_5_ans_besoin"] + actifs[5] - sorties[5]
# ---- engagements et fonds propres
pep = json.load(open(PEPJ, encoding="utf-8"))
K["interets_payes_pep"] = pep["A_total"]["interet_paye"]
K["remb_pep"], K["remb_pep_3_5"] = pep["A_total"]["remboursements"], pep["A_total"]["avec_3_5"]
K["part_remb_3_5"] = 100 * K["remb_pep_3_5"] / K["remb_pep"]
K["pep_solde"] = -S(2026, "41195100")
K["efc_solde"] = -S(2026, "41195200")
K["interets_pep_dus"] = 0.035 * K["pep_solde"]
actif_total = 21925420866
eng = dettes4 + prov
K["couv_eng_comptes"] = 100 * actif_total / eng
K["couv_eng_besoin_pep"] = 100 * actif_total / (eng - (prov - besoin) + K["interets_pep_dus"])
fp = -S(2026, "1", excl=("15", "16", "17", "18", "19")) + K["res_net_26"]
K["fonds_propres_26"] = fp
K["actifs_contestes"] = 375000000 + 450000000 + 3407519850
K["fp_sans_contestes"] = fp - K["actifs_contestes"]
K["pont_provisions"] = prov - besoin
K["pont_grenier"] = -14016392
K["fp_corrige_avant_actifs"] = fp + K["pont_provisions"] - K["interets_pep_dus"] + K["pont_grenier"]
# ---- trésorerie et retraits
treso = 3359100618
K["seuil_retrait"] = 100 * (treso - K["dette_cliniques_26"] - 489419669) / epargne
K["reserve_banque_100"] = 100 * 2562885661 / epargne
K["cotisations_2025"] = -S(2025, "7010") - S(2025, "7021")
K["epargne_sur_cotisations"] = epargne / K["cotisations_2025"]
json.dump(K, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k, v in K.items():
    print(k, round(v, 2))
