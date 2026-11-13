# -*- coding: utf-8 -*-
"""Construit scripts/rapport_v13_fppn.py à partir de rapport_v12_fppn.py (version du 13/11/2026).

Version 13 :
- revue indépendante vérifiée (registre du 13/11/2026) : corrections du texte, extourne de la provision de la clinique
  (base retraitée par base_v13_fppn.py, chiffres relus dans chiffres_v13.json) ;
- règles de rédaction du 12 et du 13/11/2026 : titres-messages et sous-titres-conclusions (titres_v13.py), taux en
  « X FCFA pour 100 FCFA, soit X % », résumé exécutif en 8 diapositives (resume_v13.py), tableaux en 12 points au moins
  découpés en 1/2, 2/2, validité au 11/11/2026 ;
- seules les nouveautés du 13/11/2026 sont en bleu : le bleu des versions précédentes repasse en noir ;
- diapositives nouvelles (ajouts_v13.py) placées après la diapositive qu'elles complètent.
Utilisation : python patch_v13.py   (écrit scripts/rapport_v13_fppn.py)
Le rapport se génère ensuite avec les arguments de la v12 suivis de <chiffres_v13.json>.
"""
import os
import re

ICI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(ICI)
s = open(os.path.join(SCRIPTS, "rapport_v12_fppn.py"), encoding="utf-8").read()


def rep(a, b, n=1):
    global s
    k = s.count(a)
    assert (k == n) if n else k >= 1, (k, a[:120])
    s = s.replace(a, b)


# ---- 1. arguments : chiffres v13 et dossier v13
rep("RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J, V11J, PBCJ = sys.argv[1:14]",
    "RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J, V11J, PBCJ, K13J = sys.argv[1:15]\n"
    "V13 = os.path.join(os.path.dirname(os.path.abspath(__file__)), \"v13\")")
rep("V10 = json.load(open(V10J, encoding=\"utf-8\"))\n",
    "V10 = json.load(open(V10J, encoding=\"utf-8\"))\nK13 = json.load(open(K13J, encoding=\"utf-8\"))\n"
    "exec(open(os.path.join(V13, \"outils_v13.py\"), encoding=\"utf-8\").read())\n")

# ---- 2. bleu : seules les nouveautés du 13/11/2026
lignes = s.split("\n")
i199 = next(i for i, l in enumerate(lignes) if l.startswith('RT = {x["code"]: x for x in Q["ratios"]}; exec('))
lignes.insert(i199 + 1, "exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), \"v13\", \"bleu_v13.py\"), encoding=\"utf-8\").read())")
s = "\n".join(lignes)

# ---- 3. résumé exécutif remplacé (les calculs du bloc sont conservés, ses diapositives retirées)
deb = s.index('separateur("R", "Résumé exécutif"')
fin_sep = s.index("\n", s.index("Cette partie peut être présentée seule", deb)) + 1
s = s[:fin_sep] + "_N0_R13, _SL0_R13 = _snap(), len(SL)\n" + s[fin_sep:]
p1 = s.index('separateur(1, "Présentation du FPPN"')
s = (s[:p1] + "del SL[_SL0_R13:]\n_omis(_N0_R13)\n"
     "exec(open(os.path.join(V13, \"resume_v13.py\"), encoding=\"utf-8\").read())\n" + s[p1:])

# ---- 4. corrections du texte (registre de la revue, 13/11/2026)
exec(open(os.path.join(ICI, "corrections_v13.py"), encoding="utf-8").read())

# ---- 5. diapositives nouvelles (ajoutées avant les listes, placées ensuite par la passe v13)
rep('SL.append({"type": "liste_figures", "notes": ""})',
    'exec(open(os.path.join(V13, "ajouts_v13.py"), encoding="utf-8").read())\nSL.append({"type": "liste_figures", "notes": ""})')

# ---- 6. passe v13 avant la numérotation : titres, renvois, placement, découpage des tableaux
rep("_POS = {}\nfor i, x in enumerate(SL, 1):",
    "exec(open(os.path.join(V13, \"passe_v13.py\"), encoding=\"utf-8\").read())\n_POS = {}\nfor i, x in enumerate(SL, 1):")

# ---- 7. diapositives sur les dettes, placées sur leurs emplacements après le rendu
rep('if "BLEU_P" in globals():  # v14',
    'exec(open(os.path.join(V13, "dettes_v13.py"), encoding="utf-8").read())\nif "BLEU_P" in globals():  # v14')

# ---- 7 bis. annexes ajoutées après la passe : découpage en 12 points juste avant le rendu (après les contenus, numéros inchangés)
rep("selection(TIMESAVER, tmp, [x[\"gab\"] if x[\"type\"] == \"gab\" else GAB[x[\"type\"]] for x in SL])",
    "exec(open(os.path.join(V13, \"annexes_v13.py\"), encoding=\"utf-8\").read())\nselection(TIMESAVER, tmp, [x[\"gab\"] if x[\"type\"] == \"gab\" else GAB[x[\"type\"]] for x in SL])")

# ---- 8. contrôle des chiffres : un chiffre d'un tableau découpé peut être sur l'une des diapositives du découpage
rep('    t_["retrouve"] = t_["texte"] in textes[t_["diapositive"] - 1] if t_["diapositive"] <= len(textes) else False',
    '    t_["retrouve"] = any(t_["texte"] in textes[t_["diapositive"] - 1 + j_] for j_ in range(t_.get("_n13", 1)) if t_["diapositive"] + j_ <= len(textes))')

# ---- 9. tableaux : police d'au moins 12 points
rep("    taille = taille or spec.get(\"taille\") or (11 if len(lignes) <= 8 else 10)",
    "    taille = max(12, taille or spec.get(\"taille\") or 12)")
rep("    while sum(hs) > hmax and taille > 7.5:", "    while sum(hs) > hmax and taille > 12:")

open(os.path.join(SCRIPTS, "rapport_v13_fppn.py"), "w", encoding="utf-8").write(s)
print("rapport_v13_fppn.py écrit,", s.count("\n"), "lignes")
