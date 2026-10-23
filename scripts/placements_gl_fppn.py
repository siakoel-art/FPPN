# -*- coding: utf-8 -*-
"""
FPPN - Placements lus dans les grands livres : encours moyen pondéré par les jours et intérêts reçus
===================================================================================================
Les balances ne donnent que le solde de clôture. Le DAT EFC (compte 52210200) est un dépôt renouvelé chaque mois
et soldé avant la clôture : seul le grand livre permet de mesurer l'encours réellement placé pendant l'année.
Comptes suivis : DAT (5221…), intérêts courus (5261…), FCP (5330…), produits financiers (77…).
Utilisation : python placements_gl_fppn.py <sortie.json> <an>=<grand livre des comptes (.xlsx ou .csv)> ...
"""
import csv
import datetime as dt
import json
import re
import sys

import openpyxl

SORTIE = sys.argv[1]
SUIVIS = ("5221", "5261", "5330", "77")


def lignes(path):
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            for r in csv.reader(f, delimiter=";"):
                yield r
    else:
        for r in openpyxl.load_workbook(path, read_only=True).worksheets[0].iter_rows(values_only=True):
            yield ["" if c is None else c for c in r]


def nombre(x):
    if isinstance(x, (int, float)):
        return float(x)
    x = str(x).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return None


def date(x):
    if isinstance(x, dt.datetime):
        return x.date()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(x))
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"(\d{2})/(\d{2})/(\d{2,4})", str(x))
    if m:
        a = int(m.group(3))
        return dt.date(a + 2000 if a < 100 else a, int(m.group(2)), int(m.group(1)))
    return None


def lire(path, an):
    cd = cc = None
    cpt, ecr = None, []
    for r in lignes(path):
        txt = [str(c).replace("\n", " ") for c in r]
        if cd is None:
            for j, c in enumerate(txt):
                if "Mouvement" in c and "débit" in c.lower():
                    cd = j
                if "Mouvement" in c and "crédit" in c.lower():
                    cc = j
            continue
        if txt and re.fullmatch(r"\d{8}", txt[0].strip()):
            cpt = txt[0].strip()
            continue
        d = date(r[0]) if r else None
        if d is None or cpt is None or not cpt.startswith(SUIVIS):
            continue
        # le montant peut être décalé d'une colonne par rapport à l'en-tête
        deb = next((nombre(r[k]) for k in (cd, cd + 1) if k < len(r) and nombre(r[k]) is not None), None)
        cre = next((nombre(r[k]) for k in (cc, cc + 1) if k < len(r) and nombre(r[k]) is not None), None)
        if deb is not None and cre is not None and cd + 1 >= cc:
            cre = None
        ecr.append({"compte": cpt, "date": d.isoformat(), "libelle": txt[3].strip() if len(txt) > 3 else "",
                    "debit": deb or 0.0, "credit": cre or 0.0})
    return ecr


def moyenne_ponderee(ecr, comptes, an, ouverture=0.0):
    """Solde moyen sur l'année, pondéré par le nombre de jours où chaque solde est resté en place."""
    deb, fin = dt.date(an, 1, 1), (dt.date(2026, 8, 31) if an == 2026 else dt.date(an, 12, 31))  # 2026 : arrêté provisoire du 31/08
    mvts = sorted((dt.date.fromisoformat(e["date"]), e["debit"] - e["credit"]) for e in ecr
                  if e["compte"].startswith(comptes) and not e["libelle"].startswith(("A.N", "AN ", "A N")))
    solde, jour, cumul, pic = ouverture, deb, 0.0, ouverture
    for d, m in mvts:
        d = max(d, deb)
        cumul += solde * (d - jour).days
        solde += m
        pic = max(pic, solde)
        jour = d
    cumul += solde * ((fin - jour).days + 1)
    return cumul / ((fin - deb).days + 1), solde, pic


OUT = {}
for arg in sys.argv[2:]:
    an, path = arg.split("=", 1)
    an = int(an)
    e = lire(path, an)
    ouv = {}
    for x in e:
        if x["libelle"].startswith(("A.N", "AN ", "A N")):
            ouv[x["compte"]] = ouv.get(x["compte"], 0.0) + x["debit"] - x["credit"]
    res = {"fichier": path.split("/")[-1], "ecritures": len(e)}
    for cle, cpts in (("dat_efc", ("52210200", "52210201")), ("dat_pep", ("52210300",)), ("dat_orabank", ("52210400",)),
                      ("fcp", ("5330",)), ("dat_tous", ("5221",)), ("placements", ("5221", "5330"))):
        o = sum(v for k, v in ouv.items() if k.startswith(cpts))
        m, s, p = moyenne_ponderee(e, cpts, an, o)
        res[cle] = {"ouverture": round(o), "moyenne": round(m), "cloture": round(s), "pic": round(p)}
    fm = []
    for mois in range(1, 13 if an < 2026 else 9):
        lim = dt.date(an, mois, 28) + dt.timedelta(days=4)
        lim = lim - dt.timedelta(days=lim.day)          # dernier jour du mois
        fm.append(round(sum(x["debit"] - x["credit"] for x in e if x["compte"] in ("52210200", "52210201")
                            and dt.date.fromisoformat(x["date"]) <= lim and not x["libelle"].startswith(("A.N", "AN ", "A N")))))
    res["dat_efc_fin_de_mois"] = fm
    res["detail"] = [{"date": x["date"], "compte": x["compte"], "debit": x["debit"], "credit": x["credit"]} for x in sorted(e, key=lambda z: z["date"])]
    res["produits_77"] = round(sum(x["credit"] - x["debit"] for x in e if x["compte"].startswith("77")))
    res["interets_dat_recus"] = round(sum(x["credit"] - x["debit"] for x in e if x["compte"].startswith("77")
                                          and not re.search(r"COURUS", x["libelle"], re.I)))
    res["renouvellements_efc"] = sum(1 for x in e if x["compte"] in ("52210200", "52210201") and "REN" in x["libelle"].upper())
    OUT[an] = res
    print(an, json.dumps({k: v for k, v in res.items() if k != "detail"}, ensure_ascii=False))
json.dump(OUT, open(SORTIE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
