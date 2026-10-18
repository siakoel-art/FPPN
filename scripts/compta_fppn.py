# -*- coding: utf-8 -*-
"""
FPPN - Lecture des balances comptables (exports Sage 100) de l'extraction
=========================================================================

Sortie : compta_balances.parquet, une ligne par compte et par exercice :
  ouverture (débit, crédit), mouvements de la période (débit, crédit), soldes cumulés (débit, crédit).
Contrôle : compta_controle.parquet compare, pour chaque fichier, la somme des comptes lus aux lignes
« Totaux de la balance » imprimées par Sage.
"""
import csv
import os
import re
import sys
import unicodedata

import pandas as pd

csv.field_size_limit(2**31 - 1)
CHAMPS = ["ouv_d", "ouv_c", "mvt_d", "mvt_c", "sol_d", "sol_c"]


def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().strip()


def nombre(v):
    s = str(v).strip().replace(" ", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


def lire_balance(chemin):
    lignes = list(csv.reader(open(chemin, encoding="utf-8"), delimiter=";"))
    entete = " ".join(" ".join(l) for l in lignes[:6])
    m = re.search(r"P.riode du;*\s*(\d{4}-\d\d-\d\d)", ";".join(";".join(l) for l in lignes[:4]))
    du = m.group(1) if m else None
    m = re.search(r"au;*\s*(\d{4}-\d\d-\d\d)", ";".join(";".join(l) for l in lignes[:4]))
    au = m.group(1) if m else None
    provisoire = "PROVISOIRE" in norm(entete)
    # positions des colonnes Débit / Crédit
    pos = None
    for l in lignes[:20]:
        n = [i for i, c in enumerate(l) if norm(c) in ("DEBIT", "CREDIT")]
        if len(n) == 6:
            pos = n
            break
    comptes, totaux = [], {}
    for l in lignes:
        if not l:
            continue
        valeurs = [(i, nombre(c)) for i, c in enumerate(l) if i > 3 and nombre(c) is not None]
        slot = {}
        for i, v in valeurs:
            k = min(range(6), key=lambda k: abs(i - (pos[k] + 1)))
            slot[CHAMPS[k]] = v
        compte = l[0].strip()
        libelle = next((c.strip() for c in l[1:9] if c.strip()), "")
        if re.match(r"^\d{6,10}$", compte):
            comptes.append({"compte": compte, "intitule": libelle, **{c: slot.get(c, 0.0) for c in CHAMPS}})
        elif "TOTAUX DE LA BALANCE" in norm(" ".join(l)):
            totaux = {c: slot.get(c, 0.0) for c in CHAMPS}
    return comptes, totaux, du, au, provisoire


def construire(dossier_donnees, dossier_sortie):
    J = list(csv.DictReader(open(os.path.join(dossier_donnees, "journal_fichiers.csv"), encoding="utf-8-sig"),
                            delimiter=";"))
    F = {}
    for f in csv.DictReader(open(os.path.join(dossier_donnees, "journal_feuilles.csv"), encoding="utf-8-sig"),
                            delimiter=";"):
        F.setdefault(f["id"], []).append(f)
    lignes, controle = [], []
    for j in J:
        if "DAFC" not in j["dossier"] or not j["action"].startswith("Converti"):
            continue
        for f in F.get(j["id"], []):
            if norm(f["feuille"]) not in ("BALANCE DES COMPTES", "BALANCE DES TIERS"):
                continue
            nature = "comptes" if "COMPTES" in norm(f["feuille"]) else "tiers"
            comptes, totaux, du, au, prov = lire_balance(os.path.join(dossier_donnees, "tableaux", f["csv"]))
            for c in comptes:
                lignes.append({"nature": nature, "periode_du": du, "periode_au": au, "provisoire": prov,
                               "src": j["id"], "fichier": j["fichier"], **c})
            somme = {c: sum(x[c] for x in comptes) for c in CHAMPS}
            controle.append({"nature": nature, "periode_du": du, "periode_au": au, "provisoire": prov,
                             "fichier": j["fichier"], "comptes": len(comptes),
                             **{f"ecart_{c}": round(somme[c] - totaux.get(c, float("nan")), 2) for c in CHAMPS}})
    bal = pd.DataFrame(lignes)
    for c in ("periode_du", "periode_au"):
        bal[c] = pd.to_datetime(bal[c])
    bal["exercice"] = bal["periode_au"].dt.year
    bal.to_parquet(os.path.join(dossier_sortie, "compta_balances.parquet"), index=False)
    ctl = pd.DataFrame(controle)
    ctl.to_parquet(os.path.join(dossier_sortie, "compta_controle.parquet"), index=False)
    return ctl


if __name__ == "__main__":
    ctl = construire(sys.argv[1] if len(sys.argv) > 1 else "donnees", sys.argv[2] if len(sys.argv) > 2 else "base")
    print(ctl.to_string())
