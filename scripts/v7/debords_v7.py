# -*- coding: utf-8 -*-
"""Compare les mesures PowerPoint de v6 et v7 : liste les chevauchements nouveaux en v7.
Utilisation : python debords_v7.py <mes_v6.json> <mes_v7.json>"""
import json
import sys
from collections import defaultdict


def charge(f):
    out = defaultdict(list)
    for r in json.load(open(f, encoding="utf-8-sig")):
        out[r["n"]].append(r)
    return out


IGNORE = ("ZoneTexte 25",)


def conflits(formes):
    res = set()
    for a in formes:
        for b in formes:
            if a is b or a["nom"] in IGNORE or b["nom"] in IGNORE:
                continue
            if b["haut"] <= a["haut"] + 1 or min(a["d"], b["d"]) - max(a["g"], b["g"]) < 5:
                continue
            if a["bas"] > b["haut"] + 2:
                res.add((a["nom"], b["nom"], round(a["bas"] - b["haut"], 1)))
        if a["bas"] > 540 and a["nom"] not in IGNORE:
            res.add((a["nom"], "bas de page", round(a["bas"] - 540, 1)))
    return res


v6, v7 = charge(sys.argv[1]), charge(sys.argv[2])
tot = 0
for n in sorted(v7):
    anc = {(x, y) for x, y, _ in conflits(v6[n])}
    neufs = [c for c in conflits(v7[n]) if (c[0], c[1]) not in anc]
    if neufs:
        tot += 1
        print(f"D{n}", "; ".join(f"{a} déborde sur {b} de {e} pt" for a, b, e in sorted(neufs)))
print(tot, "diapositives avec un débordement nouveau")
