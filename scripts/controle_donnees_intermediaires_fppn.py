# -*- coding: utf-8 -*-
"""
FPPN - Contrôle des données intermédiaires reconstituées (13/11/2026)
Compare, onglet par onglet, les nombres des classeurs régénérés (dossier des résultats) avec ceux d'un classeur
de référence recalculé par Excel (les formules du classeur livré n'ont pas de valeur en cache).
Comparaison des multiensembles de nombres arrondis à l'unité (valeurs absolues >= 1).
Utilisation : python controle_donnees_intermediaires_fppn.py <classeur_reference_recalcule.xlsx> <dossier_resultats>
"""
import sys, glob, os, collections
from python_calamine import CalamineWorkbook as W
ref = W.from_path(sys.argv[1]); R = sys.argv[2]
def nums(rows):
    c = collections.Counter()
    for r in rows:
        for v in r:
            if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) >= 1:
                c[round(float(v))] += 1
    return c
refs = {s: nums(ref.get_sheet_by_name(s).to_python()) for s in ref.sheet_names}
out = []
for f in sorted(glob.glob(os.path.join(R, "*.xlsx"))):
    w = W.from_path(f)
    for s in w.sheet_names:
        n = nums(w.get_sheet_by_name(s).to_python())
        if not n: continue
        if s in refs:
            tot = sum(n.values()); ok = sum(min(k, refs[s][v]) for v, k in n.items())
            manq = sorted((v for v in n if v not in refs[s]), key=abs, reverse=True)[:5]
            out.append((os.path.basename(f), s, tot, ok, manq))
        else:
            out.append((os.path.basename(f), s, sum(n.values()), None, []))
for f, s, t, ok, m in out:
    print(f"{f[:28]:28} {s[:28]:28} {t:6} " + ("absent du classeur" if ok is None else f"{ok:6} ({100*ok/t:6.2f} %) " + str(m)))
