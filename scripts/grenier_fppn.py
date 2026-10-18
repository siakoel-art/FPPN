# -*- coding: utf-8 -*-
"""
FPPN - Lecture des exports du logiciel de prêts du Grenier du Policier (Zenith Technology)
=========================================================================================

Les fichiers « BalanceAgeeDetaille » et « production_de_la_periode » portent l'extension .xls mais
sont des tableaux au format page web. Ce script les lit sans perte et produit :
  - grenier_balances.parquet   : une ligne par prêt et par date d'arrêté (encours, retard, échéances) ;
  - grenier_production.parquet : une ligne par prêt débloqué sur la période du fichier ;
  - grenier_controle.parquet   : pour chaque fichier, nombre de prêts et totaux lus, comparés aux
                                 totaux imprimés par le logiciel en bas du fichier.
Entrée : le dossier « donnees » de l'extraction. Sortie : le dossier de la base.
"""
import csv
import io
import os
import re
import sys
import unicodedata

import pandas as pd


def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().strip()


def texte(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    m = re.match(r'^="(.*)"$', s)
    return m.group(1) if m else s


def montant(v):
    s = re.sub(r"[\s\u00a0\u202f]", "", texte(v).replace("XOF", ""))
    if s in ("", "nan"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def date_fr(s):
    return pd.to_datetime(s.replace("-", "/"), format="%d/%m/%Y")


DATE = r"(\d\d[/-]\d\d[/-]\d{4})"


def portefeuille(fichier):
    m = re.search(r"\b(CRF|CRO|CRS)\b", norm(fichier))
    return m.group(1) if m else "global"


def lire_tables(chemin):
    return pd.read_html(io.StringIO(open(chemin, encoding="utf-8-sig").read()))


def construire(dossier_donnees, dossier_sortie):
    J = [l for l in csv.DictReader(open(os.path.join(dossier_donnees, "journal_fichiers.csv"), encoding="utf-8-sig"),
                                   delimiter=";") if "page web" in l["action"] and "GRENIER" in norm(l["dossier"])]
    balances, production, controle = [], [], []
    for l in J:
        chemin = os.path.join(dossier_donnees, l["sortie"].replace("\\", os.sep))
        tables = lire_tables(chemin)
        titre = " ".join(texte(x) for x in tables[0].iloc[0].tolist())
        pf = portefeuille(l["fichier"])
        detail = max(tables, key=lambda t: t.shape[0])
        cols = [norm(c) for c in detail.columns]
        detail.columns = cols
        lignes = detail[~detail.iloc[:, 0].astype(str).str.upper().str.startswith(("TOTAL", "NAN"))]
        lignes = lignes[lignes.iloc[:, 1].astype(str).str.upper() != "TOTAL"]
        if "BALANCE AGEE" in norm(titre):
            m = re.search(r"AU " + DATE, titre)
            arrete = date_fr(m.group(1)) if m else pd.NaT
            for _, r in lignes.iterrows():
                balances.append({
                    "arrete": arrete, "portefeuille": pf, "compte": texte(r.get("N COMPTE")),
                    "mecano": texte(r.get("MECANO")), "nom": texte(r.get("NOM ET PRENOMS BENEF.")),
                    "montant_debloque": montant(r.get("MONTANT DEBLOQUE")),
                    "date_deblocage": pd.to_datetime(texte(r.get("DATE DEBLOCAGE")), errors="coerce"),
                    "date_premiere_echeance": pd.to_datetime(texte(r.get("DATE 1ERE ECHEANCE")), errors="coerce"),
                    "date_derniere_echeance": pd.to_datetime(texte(r.get("DATE DERNIERE ECHEANCE")), errors="coerce"),
                    "montant_echeance": montant(r.get("MONTANT ECHEANCE")),
                    "prochaine_echeance": pd.to_datetime(texte(r.get("PROCHAINE ECHEANCE")), errors="coerce"),
                    "duree_mois": montant(r.get("DUREE")), "taux": montant(r.get("TAUX INTERET")),
                    "depot_garantie": montant(r.get("DEPOT DE GARANTIE")),
                    "encours": montant(r.get("ENCOURS CREDIT")), "retard": montant(r.get("DUREE RETARD")),
                    "src": l["id"], "fichier": l["fichier"]})
            # totaux imprimés : table « TOTAL DES MONTANTS DEBLOQUES … NOMBRE »
            imprime = {}
            for t in tables:
                c = [norm(x) for x in t.columns]
                if "NOMBRE" in c:
                    t.columns = c
                    tot = t[t.iloc[:, 0].astype(str).str.upper().str.startswith("TOTAL")]
                    if len(tot):
                        imprime = {"debloque_imprime": montant(tot.iloc[0]["TOTAL DES MONTANTS DEBLOQUES"]),
                                   "encours_imprime": montant(tot.iloc[0]["TOTAL DES ENCOURS DE CREDIT"]),
                                   "nombre_imprime": montant(tot.iloc[0]["NOMBRE"])}
            lu = [b for b in balances if b["src"] == l["id"]]
            controle.append({"src": l["id"], "fichier": l["fichier"], "type": "balance âgée", "arrete": arrete,
                             "portefeuille": pf, "nombre_lu": len(lu),
                             "debloque_lu": sum(b["montant_debloque"] or 0 for b in lu),
                             "encours_lu": sum(b["encours"] or 0 for b in lu), **imprime})
        elif "PRODUCTION" in norm(titre):
            m = re.search(r"DU " + DATE + " AU " + DATE, titre)
            du = date_fr(m.group(1)) if m else pd.NaT
            au = date_fr(m.group(2)) if m else pd.NaT
            for _, r in lignes.iterrows():
                if not texte(r.get("MECANO")) or norm(texte(r.get("MECANO"))).startswith("TOTAL"):
                    continue
                production.append({
                    "periode_du": du, "periode_au": au, "portefeuille": pf,
                    "dossier": texte(r.get("N DOSSIER")), "mecano": texte(r.get("MECANO")),
                    "nom": texte(r.get("BENEFICIAIRE")), "client": texte(r.get("N CLIENT")),
                    "compte": texte(r.get("N COMPTE")), "type_credit": texte(r.get("TYPE CREDIT")),
                    "montant_debloque": montant(r.get("MONTANT DEBLOQUE")), "encours": montant(r.get("ENCOURS CREDIT")),
                    "date_deblocage": pd.to_datetime(texte(r.get("DATE DEBLOCAGE")), errors="coerce"),
                    "premiere_echeance": pd.to_datetime(texte(r.get("1")), errors="coerce"),
                    "derniere_echeance": pd.to_datetime(texte(r.get("DERNIERE ECHEANCE")), errors="coerce"),
                    "montant_echeance": montant(r.get("MONTANT ECHEANCE")),
                    "duree_mois": montant(r.get("DUREE (MOIS)")), "taux": montant(r.get("TAUX (%)")),
                    "type_garantie": texte(r.get("TYPE GARANTIE")),
                    "valeur_garantie": montant(r.get("VALEUR GARANTIE")),
                    "decision": texte(r.get("DECISION")), "statut": texte(r.get("STATUT")),
                    "src": l["id"], "fichier": l["fichier"]})
            lu = [p for p in production if p["src"] == l["id"]]
            controle.append({"src": l["id"], "fichier": l["fichier"], "type": "production", "arrete": au,
                             "portefeuille": pf, "nombre_lu": len(lu),
                             "debloque_lu": sum(p["montant_debloque"] or 0 for p in lu)})
    pd.DataFrame(balances).to_parquet(os.path.join(dossier_sortie, "grenier_balances.parquet"), index=False)
    pd.DataFrame(production).to_parquet(os.path.join(dossier_sortie, "grenier_production.parquet"), index=False)
    ctl = pd.DataFrame(controle)
    ctl.to_parquet(os.path.join(dossier_sortie, "grenier_controle.parquet"), index=False)
    return ctl


if __name__ == "__main__":
    ctl = construire(sys.argv[1] if len(sys.argv) > 1 else "donnees", sys.argv[2] if len(sys.argv) > 2 else "base")
    print(ctl.to_string())
