# -*- coding: utf-8 -*-
"""Correspondance des numéros de diapositive entre deux versions du rapport (alignement des titres).

Les diapositives « Méthode : … » supprimées renvoient au tableau des méthodes retenues de la partie 3.
Usage : python3 correspondance_diapos_fppn.py <ancien.pptx> <nouveau.pptx> <sortie.json>
"""
import difflib
import json
import re
import sys

from pptx import Presentation


def titres(chemin):
    out = []
    for sl in Presentation(chemin).slides:
        t = next((sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.name in ("Adenka_Titre", "Rectangle 9", "TextBox 26") and sh.text_frame.text.strip()), "")
        out.append(re.sub(r"\s+", " ", t))
    return out


A, B = titres(sys.argv[1]), titres(sys.argv[2])
corr = {}
sm = difflib.SequenceMatcher(None, A, B, autojunk=False)
for op, i1, i2, j1, j2 in sm.get_opcodes():
    if op == "equal" or (op == "replace" and i2 - i1 == j2 - j1):
        for k in range(i2 - i1):
            corr[i1 + k + 1] = j1 + k + 1
# diapositives déplacées (même titre ailleurs) : rattachées par le titre
pos_b = {}
for j, t in enumerate(B, 1):
    pos_b.setdefault(t, []).append(j)
meth = [j for j, t in enumerate(B, 1) if t == "Méthodologie & Limites"][1:]  # sans le séparateur
for i, t in enumerate(A, 1):
    if i in corr:
        continue
    if t.replace(" et ", " & ").lower() in {u.lower(): u for u in B} and t.startswith("Outils"):
        corr[i] = B.index(next(u for u in B if u.lower() == t.replace(" et ", " & ").lower())) + 1
    elif t in pos_b and len(pos_b[t]) == 1:
        corr[i] = pos_b[t][0]
    elif t.startswith("Méthode : ") and meth:
        corr[i] = meth[0]
manque = [(i, t) for i, t in enumerate(A, 1) if i not in corr]
json.dump({"correspondance": corr, "sans_correspondance": manque, "ancien": len(A), "nouveau": len(B)}, open(sys.argv[3], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("correspondances :", len(corr), "sans correspondance :", manque[:20])
