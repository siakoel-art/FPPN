# -*- coding: utf-8 -*-
"""Calculs des deux anomalies pour la note à la responsable de mission (17/11/2026).
1. Dépenses qui ne sont pas des soins dans les prestations maladie : prestations et sinistralité avant et après, 2021 au 31/08/2026.
2. Provision des factures de la clinique (CMP) d'octobre à décembre 2025, factures reçues en 2026, et autres factures de 2025 passées en 2026.
Entrées : gl_prestations.parquet (extrait_gl.py, écritures des grands livres) et compta_balances.parquet (balances du client lues
telles quelles, seuls les comptes 70 de cotisations sont pris ici).
Sortie : resultats.json et tables .parquet dans le dossier de travail.
Utilisation : python calcul.py <dossier de travail> <compta_balances.parquet>"""
import json
import os
import re
import sys

import pandas as pd

W, BALP = sys.argv[1], sys.argv[2]
gl = pd.read_parquet(os.path.join(W, "gl_prestations.parquet"))
gl = gl[gl.journal != "RAN"].copy()
ANS = list(range(2021, 2027))

# ---------------------------------------------------------------- cotisations (balances du client)
b = pd.read_parquet(BALP)
b = b[(b.nature == "comptes") & b.exercice.isin(ANS)]
FBAL = b.groupby("exercice").fichier.first().to_dict()
cot = lambda an, pre: float((b[(b.exercice == an) & b.compte.str.startswith(pre)].eval("mvt_c - mvt_d")).sum())
COT = {an: {"maladie": cot(an, "7010"), "deces": cot(an, "7021"), "toutes": cot(an, "70")} for an in ANS}

# ---------------------------------------------------------------- prestations (grands livres)
p = gl[gl.compte.str.startswith("60")].copy()
MAL = p.compte.str.startswith("601") | (p.compte == "60800007")       # formule maladie du rapport
CLS = p.compte.str[:4].isin(["6011", "6012", "6013", "6018", "6080"])  # comptes lus pour le classement


def cat(l):
    """Règles du rapport (sinistralite_gl_fppn.py) : classement par le libellé de l'écriture."""
    if re.search(r"\bDON\b", l) and re.search(r"SOIN|DIALYSE|DON MED|ANTI-AN|KIT D.ACCOUCH|DON DE SANG|COUVERTURE FRAIS|PHCIE", l):
        return ("Soins", "don à caractère médical, gardé dans les soins")
    if re.search(r"\bDON\b|K'DO|CADEAU|PANIER", l):
        return ("Dons (autorités, élus, délégués)", "mot DON, K'DO, CADEAU ou PANIER")
    if re.search(r"FUTURE VISI|GADGET|SAC KABA|IMPRIMERIE|IMPRIME|NEW PRINT|CARTE", l):
        return ("Gadgets, imprimés et cartes", "mot GADGET, IMPRIME, CARTE, SAC KABA ou nom d'un imprimeur")
    if re.search(r"00444\d{3}", l):
        return ("Honoraires d'un cabinet", "facture n° 00444..., tous les trois mois")
    if re.search(r"RECLASSEMENT FRAIS DE RECOUVREMENT|PROV CPTE", l):
        return ("Régularisations comptables", "reclassement ou provision de compte")
    return ("Soins", "")


ELARGI = re.compile(r"FUTURE VIS|GIS PLUS|SACS? KABA|HADJ|CEREMONIE|CEREM|LOBBING|ARBRE DE NOEL|ECRAN GEANT|LOCAT")
p["libU"] = p.libelle.str.upper()
cc = p.libU.map(cat)
p["categorie"], p["regle"] = cc.str[0], cc.str[1]
p.loc[~CLS, ["categorie", "regle"]] = ["Soins", ""]
# variante élargie : écritures restées « soins », hors journal des factures (SIN), dont le libellé désigne un imprimeur, une cérémonie, le hadj...
el = CLS & (p.categorie == "Soins") & (p.journal != "SIN") & p.libU.str.contains(ELARGI)
p["variante"] = ""
p.loc[el, "variante"] = "Autres dépenses repérées (imprimeurs, cérémonies, hadj, lobbying, locations)"
p["regle_var"] = ""
p.loc[el, "regle_var"] = p.loc[el, "libU"].str.extract("(" + ELARGI.pattern + ")", expand=False).fillna("")

R = {"sources": {"balances": {int(k): v for k, v in FBAL.items()}, "grands_livres": gl.groupby("an").fichier.first().to_dict()}}
R["sources"]["grands_livres"] = {int(k): v for k, v in R["sources"]["grands_livres"].items()}
lignes = []
for an in ANS:
    x = p[p.an == an]
    prest_mal = float(x[MAL.loc[x.index]].montant.sum())
    prest_tot = float(x.montant.sum())
    hs = float(x[CLS.loc[x.index] & (x.categorie != "Soins")].montant.sum())
    var = float(x[x.variante != ""].montant.sum())
    hs_mal = float(x[MAL.loc[x.index] & (x.categorie != "Soins")].montant.sum())
    var_mal = float(x[MAL.loc[x.index] & (x.variante != "")].montant.sum())
    lignes.append(dict(an=an, cot_maladie=COT[an]["maladie"], cot_toutes=COT[an]["toutes"], prest_maladie=prest_mal, prest_toutes=prest_tot,
                       prest_classees=float(x[CLS.loc[x.index]].montant.sum()), hors_soins=hs, hors_soins_dans_maladie=hs_mal,
                       variante=var, variante_dans_maladie=var_mal))
T = pd.DataFrame(lignes).set_index("an")
T["sc_mal_avant"] = T.prest_maladie / T.cot_maladie
T["sc_mal_apres"] = (T.prest_maladie - T.hors_soins_dans_maladie) / T.cot_maladie
T["sc_mal_apres_var"] = (T.prest_maladie - T.hors_soins_dans_maladie - T.variante_dans_maladie) / T.cot_maladie
T["sc_glob_avant"] = T.prest_toutes / T.cot_toutes
T["sc_glob_apres"] = (T.prest_toutes - T.hors_soins) / T.cot_toutes
T["sc_glob_apres_var"] = (T.prest_toutes - T.hors_soins - T.variante) / T.cot_toutes
T["part_hs"] = T.hors_soins / T.prest_classees
print(T.round(4).to_string())
T.to_parquet(os.path.join(W, "t1_avant_apres.parquet"))
cats = p[p.categorie != "Soins"].pivot_table(index="categorie", columns="an", values="montant", aggfunc="sum").fillna(0)
print(cats.round(0).to_string())
det = p[(p.categorie != "Soins") | (p.variante != "")][["an", "date", "journal", "piece", "compte", "libelle", "facture", "montant", "categorie", "regle", "variante", "regle_var", "fichier", "ligne"]]
det.sort_values(["an", "categorie", "date"]).to_parquet(os.path.join(W, "t1_detail.parquet"))
print("variante par année", p[p.variante != ""].groupby("an").montant.sum().round(0).to_dict())

# ---------------------------------------------------------------- 2. provision de la clinique fin 2025
od = gl[(gl.an == 2025) & (gl.journal == "OD") & gl.piece.isin(["805", "806", "807"])]
sin = gl[(gl.an == 2026) & (gl.journal == "SIN") & gl.piece.isin(["462", "552", "619"])]
prov = od[od.compte == "60120000"].groupby("piece").montant.sum()
fact_charge = sin[sin.compte.str.startswith("60")].groupby("piece").montant.sum()
fact_dette = sin[sin.compte == "40110000"].groupby("piece").credit.sum()
mois = [("Octobre 2025", "805", "462"), ("Novembre 2025", "806", "552"), ("Décembre 2025", "807", "619")]
M = [dict(mois=m, od=o, provision=float(prov[o]), sin=s, facture=float(fact_charge[s]), facture_dette=float(fact_dette[s])) for m, o, s in mois]
for r in M:
    assert abs(r["facture"] - r["facture_dette"]) < 1
R["clinique_mois"] = M
R["provision"] = float(prov.sum())
R["factures"] = float(fact_charge.sum())
# annulation de la provision en 2026 : crédit d'un compte 60 ou débit du 40110000 qui reprend les 255 000 000 FCFA
an26 = gl[gl.an == 2026]
R["credits_60_2026"] = float(an26[an26.compte.str.startswith("60")].credit.sum())
R["annulation_trouvee"] = bool(((an26.compte == "40110000") & an26.libelle.str.upper().str.contains("PROV")).any())
dette = gl[gl.compte == "40110000"]
R["dette_40110000_fin_08_2026"] = float(-(gl[(gl.compte == "40110000")].montant.sum()))  # sans les à-nouveaux : mouvements seuls
od.to_parquet(os.path.join(W, "t2_od.parquet"))
sin.to_parquet(os.path.join(W, "t2_sin.parquet"))
# factures de la clinique en 2026 par mois (pour repérer un mois manquant)
cmp26 = an26[(an26.compte == "40110000") & (an26.journal == "SIN") & an26.libelle.str.upper().str.contains(r"CMP|CLINIQUE MODERNE") &
             ~an26.libelle.str.upper().str.contains("ELEVE")][["date", "piece", "libelle", "facture", "credit", "fichier", "ligne"]]
cmp26.to_parquet(os.path.join(W, "t2_cmp_2026.parquet"))
print(cmp26.to_string())

# autres factures de 2025 passées en 2026 (libellé ou numéro de facture portant l'année 2025)
x = an26[an26.compte.str.startswith("60") & (an26.journal != "RAN") & ~an26.piece.isin(["462", "552", "619"])]
m25 = x.libelle.str.upper().str.contains(r"2025|/25\b|-25\b|\b25/|DEC 25|JANV-DEC 25") | x.facture.str.contains(r"2025|/25\b|-25\b|/12/25")
m25 &= ~x.facture.str.startswith("CH")  # numéro de chèque (CH0002025), pas une année
mixte = (x.libelle.str.upper() + " " + x.facture.str.upper()).str.contains(r"/26\b|26$|2026|JANV26")
aut = x[m25 & ~mixte][["date", "journal", "piece", "compte", "libelle", "facture", "montant", "fichier", "ligne"]]
aut_mixte = x[m25 & mixte][["date", "journal", "piece", "compte", "libelle", "facture", "montant", "fichier", "ligne"]]
aut.to_parquet(os.path.join(W, "t2_autres.parquet"))
aut_mixte.to_parquet(os.path.join(W, "t2_autres_mixtes.parquet"))
print(aut.to_string()); print("mixtes"); print(aut_mixte.to_string())
R["autres_2025"] = float(aut.montant.sum())
R["autres_2025_mixtes"] = float(aut_mixte.montant.sum())

# effets en 2026 : on retire les 255 000 000 FCFA (annulation de la provision), puis on combine avec l'anomalie 1
t = T.loc[2026]
corr = R["provision"]
R["effets_2026"] = {
    "sc_mal": [t.prest_maladie / t.cot_maladie, (t.prest_maladie - corr) / t.cot_maladie,
               (t.prest_maladie - corr - t.hors_soins_dans_maladie) / t.cot_maladie],
    "sc_glob": [t.prest_toutes / t.cot_toutes, (t.prest_toutes - corr) / t.cot_toutes, (t.prest_toutes - corr - t.hors_soins) / t.cot_toutes]}
R["boni_2025"] = R["provision"] - R["factures"]
print(json.dumps({k: v for k, v in R.items() if k != "sources"}, indent=1, ensure_ascii=False, default=float))
json.dump(R, open(os.path.join(W, "resultats.json"), "w", encoding="utf-8"), indent=1, ensure_ascii=False, default=float)
