# -*- coding: utf-8 -*-
"""Construit scripts/rapport_v12_fppn.py à partir de rapport_v11_fppn.py (remplacements contrôlés et insertions).

Version 12 (remarques du 23/10/2026) :
- codes couleur par boules (🔴 🟠 🟡 🟢 ⚪), sans remplissage, avec légende sous chaque tableau ;
- matrices « valeur + boule » : indicateurs clés et ratios prudentiels de 2020 au 31/08/2026, ratios par garantie, tableau 96 ;
- « retraité » réservé aux policiers à la retraite ; la lecture corrigée des comptes devient « corrigée selon le PCMS » ;
- constats au format fait, norme, écart, effet, avec la définition exacte du besoin mesuré ;
- encadré « Définition et lecture » sur les diapositives de ratios et diapositives de définitions avant chaque tableau de bord ;
- ratios par garantie avant les ratios globaux ; sinistralité réglementaire globale (Décision n° 05/2022, § 1.2.1.2.5) ;
- bandeau « Comment lire » sur la carte des risques ;
- barème tarifaire du FPPN, barème de la CGRAE, logigrammes des cotisations et circulaire n° 00230 (Code CIMA) intégrés.
"""
import os
import re

REPO = "/home/user/FPPN/scripts"
ICI = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(REPO, "rapport_v11_fppn.py"), encoding="utf-8").read()


def rep(a, b, n=1):
    global s
    k = s.count(a)
    assert k == n, (k, a[:120])
    s = s.replace(a, b)


def rep_re(motif, b, n=None):
    global s
    s2, k = re.subn(motif, b, s, flags=re.S)
    assert k >= 1 and (n is None or k == n), (k, motif[:120])
    s = s2


def tranche(debut, fin, nouveau):
    global s
    assert s.count(debut) == 1 and s.count(fin) == 1, (s.count(debut), s.count(fin), debut[:80], fin[:80])
    i, j = s.index(debut), s.index(fin)
    assert i < j
    s = s[:i] + nouveau + s[j:]


def bloc(nom):
    return open(os.path.join(ICI, nom), encoding="utf-8").read()


# ================================================================== en-tête
rep("FPPN - Rapport actuariel et statistique, version 11 (PowerPoint, charte Adenka, Timesaver)",
    "FPPN - Rapport actuariel et statistique, version 12 (PowerPoint, charte Adenka, Timesaver)")
rep("Utilisation : python rapport_v11_fppn.py", "Utilisation : python rapport_v12_fppn.py")

# ================================================================== rendement des placements des grands livres dans le ratio RL6
rep('AG = {int(k): v for k, v in Q["agregats"].items()}\n', '''AG = {int(k): v for k, v in Q["agregats"].items()}
# v12 : le taux brut de rentabilité des placements (RL6) suit l'encours lu dans les grands livres, comme la partie 9
RT["RL6"]["valeurs"] = {x["exercice"]: x["rendement"] for x in V11["rendement_gl"]}
RT["RL6"]["formule"] = "Produits financiers / encours moyen des DAT et du FCP, pondéré par les jours (grands livres)"
''')

# ================================================================== boules de couleur
rep('VERT_C, JAUNE_C = "70AD47", "FFD966"\n', bloc("moteur_boules.py"))
rep('''            if i > 0 and j in cols_c:
                cc = code_couleur(spec["entetes"][j], v)
                if cc:
                    fond, coul, gras_c = cc[0], cc[1], True
            for li_, ligne_ in enumerate(str(v).split("\\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                for morceau, gras, surl in morceaux_balises(ligne_):
                    r = p.add_run()
                    r.text = morceau
                    r.font.name = "Helvetica"
                    r.font.size = Pt(taille)
                    r.font.bold = gras_c or gras
                    r.font.color.rgb = RGBColor.from_string(coul)
                    if surl:
                        surligner(r)''', '''            if i > 0 and j in cols_c:
                cc = code_couleur(spec["entetes"][j], v)
                if cc:
                    b_, t_ = BOULE.get(cc[0], ""), str(v).strip()
                    v = b_ if (len(t_) <= 32 and ":" not in t_ and "\\n" not in t_) else b_ + " " + t_
            if str(v) in BOULES:
                p.alignment = PP_ALIGN.CENTER
            for li_, ligne_ in enumerate(str(v).split("\\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                for morceau, gras, surl in morceaux_balises(ligne_):
                    runs_texte(p, morceau, taille, gras_c or gras, coul, surl)''')
rep('''    cols_c = colonnes_codees(spec)
    if cols_c:
        hmax -= 0.62''', '''    cols_c = colonnes_codees(spec)
    if cols_c or spec.get("legende"):
        hmax -= 0.62''')
rep('''    if cols_c:
        return legende_codes(slide, spec, cols_c, x0, y + sum(hs) * 1.07 + 0.12, sum(larg))''',
    '''    if cols_c or spec.get("legende"):
        return legende_codes(slide, spec, cols_c, x0, y + sum(hs) * 1.07 + 0.12, sum(larg))''')
tranche("def legende_codes(slide, spec, cols, x0, y, largeur):", "def hauteurs_lignes(spec, taille):", bloc("moteur_legende.py"))

# ================================================================== encadré « Définition et lecture »
rep('''    if not paras:
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - 0.25 - _hv(x, sum(spec["largeurs"])))
        valable(slide, x, bas + 0.3, sum(spec["largeurs"]))
        return''', '''    enc = x.get("encadre")
    h_enc = (hauteur_texte(enc, 30.2, 10, 3) + 0.95) if enc else 0.0
    if not paras:
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - 0.25 - _hv(x, sum(spec["largeurs"])) - (h_enc + 0.25 if enc else 0))
        bas = valable(slide, x, bas + 0.3, sum(spec["largeurs"]))
        if enc:
            encadre(slide, 1.4, bas + 0.25, 31.2, h_enc, enc, x.get("encadre_titre", "Définition et lecture"))
        return''')
rep('''        xt = 1.4 + sum(spec["largeurs"]) + 0.6
        corps(slide, xt, 3.85, 32.6 - xt, BAS_CONTENU - 3.85, paras)''', '''        xt = 1.4 + sum(spec["largeurs"]) + 0.6
        y_d = 3.85
        if enc:
            h_d = hauteur_texte(enc, 32.6 - xt - 0.8, 10, 3) + 0.95
            encadre(slide, xt, y_d, 32.6 - xt, h_d, enc, x.get("encadre_titre", "Définition et lecture"))
            y_d += h_d + 0.3
        corps(slide, xt, y_d, 32.6 - xt, BAS_CONTENU - y_d, paras)''')
rep('''        besoin = max(1.6, min(4.2, hauteur_texte(paras, 31.2, 11) + 0.2))
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.55 - _hv(x))
        bas = valable(slide, x, bas + 0.3, 31.2)
        corps(slide, 1.4, bas + 0.3, 31.2, BAS_CONTENU - bas - 0.3, paras, 13)''', '''        besoin = max(1.6, min(4.2, hauteur_texte(paras, 31.2, 11) + 0.2))
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.55 - _hv(x) - (h_enc + 0.25 if enc else 0))
        bas = valable(slide, x, bas + 0.3, 31.2)
        if enc:
            encadre(slide, 1.4, bas + 0.25, 31.2, h_enc, enc, x.get("encadre_titre", "Définition et lecture"))
            bas += h_enc + 0.25
        corps(slide, 1.4, bas + 0.3, 31.2, BAS_CONTENU - bas - 0.3, paras, 13)''')

# ================================================================== carte des risques : bandeau « Comment lire » au-dessus du schéma
rep('''        for (haut, bas), c_, e_ in zip(paires, d["causes"], d["effets"]):
            ecrire(haut, [c_])
            ecrire(bas, [e_])''', '''        for (haut, bas), c_, e_ in zip(paires, d["causes"], d["effets"]):
            ecrire(haut, [c_])
            ecrire(bas, [e_])
        if d.get("lecture"):
            LECT83[id(slide)] = d["lecture"]''')
rep('''print("retouches de mise en page réappliquées :", _appl)''', '''print("retouches de mise en page réappliquées :", _appl)
# v12 : bandeau « Comment lire » de la carte des risques, posé après les retouches de l'utilisateur
for sl in prs.slides:
    if id(sl) in LECT83 or any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text == "Carte des risques" for sh in sl.shapes):
        lect_ = LECT83.get(id(sl)) or next(iter(LECT83.values()), None)
        if lect_:
            compresser(sl, 3.5, 16.3, 6.1, 16.3)
            encadre(sl, 1.4, 3.85, 31.2, 2.0, [lect_], "Comment lire")''')

exec(open(os.path.join(ICI, "etape2.py"), encoding="utf-8").read())
open(os.path.join(REPO, "rapport_v12_fppn.py"), "w", encoding="utf-8").write(s)
print("rapport_v12_fppn.py écrit")
