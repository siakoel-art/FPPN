# -*- coding: utf-8 -*-
"""Met à jour, dans les scripts des livrables, les numéros de diapositive du rapport actuariel et sa date de version.

Remplace « diapositive(s) 95 à 98, 112 et 126 », « Diap. 5, 95 à 98 » et les listes de numéros seules (colonnes
« Détail ») selon la correspondance ancienne → nouvelle version, et « rapport actuariel et statistique du 23/10/2026 »
par la date de la nouvelle version.

Usage : python3 maj_renvois_diapos_fppn.py <corr.json> <date> <fichier.py> [--listes-seules] …
"""
import json
import re
import sys

CORR = {int(k): v for k, v in json.load(open(sys.argv[1], encoding="utf-8"))["correspondance"].items()}
DATE = sys.argv[2]
LISTE = r"\d{1,3}(?:(?:, | à | et )\d{1,3})*"


def remap_liste(txt):
    parts = re.split(r"(, | à | et )", txt)
    out = []
    for p in parts:
        if p.isdigit():
            out.append(str(CORR.get(int(p), int(p))))
        else:
            out.append(p)
    # une plage dont les bornes se suivent ou s'inversent devient une énumération simple
    res = "".join(out)
    res = re.sub(r"(\d+) à (\d+)", lambda m: m.group(1) if m.group(1) == m.group(2) else (f"{m.group(2)} à {m.group(1)}" if int(m.group(1)) > int(m.group(2)) else m.group(0)), res)
    return res


def traiter(texte, listes_seules):
    n = 0

    def sub_pref(m):
        nonlocal n
        n += 1
        return m.group(1) + remap_liste(m.group(2))
    texte = re.sub(r"(\b[Dd]iapositives? |\bDiap\. )(" + LISTE + ")", sub_pref, texte)
    if listes_seules:
        def sub_lit(m):
            nonlocal n
            n += 1
            return '"' + remap_liste(m.group(1)) + '"'
        texte = re.sub(r'"(' + LISTE + r')"', sub_lit, texte)
    texte = re.sub(r"(rapport actuariel et statistique du )23/10/2026", r"\g<1>" + DATE, texte, flags=re.I)
    return texte, n


args = sys.argv[3:]
listes = "--listes-seules" in args
for f in [a for a in args if not a.startswith("--")]:
    t = open(f, encoding="utf-8").read()
    t2, n = traiter(t, listes)
    open(f, "w", encoding="utf-8").write(t2)
    print(f, "renvois traités :", n)
