# -*- coding: utf-8 -*-
"""Construit scripts/rapport_v11_fppn.py à partir de rapport_v10_fppn.py (remplacements contrôlés et insertions).

Version 11 (vérification de cohérence du 23/10/2026) :
- surlignage jaune des affirmations qui reposent sur une déclaration ou attendent une pièce du FPPN ([[…]] dans le texte) ;
- codes couleur dans toutes les colonnes de statut, priorité, criticité, horizon, avec un tableau de signification sous le tableau ;
- placements lus dans les grands livres (encours moyen pondéré par les jours) ; provision du personnel avec le plafond SMIG ;
- corrections de la vérification de cohérence ; formulations établies sur pièces ; numérotation de la PBC du 23/10/2026 ;
- synthèse des constats, recommandations avec horizon, ratios par garantie, qualité des bases, dispositif actuariel.
"""
import os
import re

REPO = "/home/user/FPPN/scripts"
ICI = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(REPO, "rapport_v10_fppn.py"), encoding="utf-8").read()


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


# ================================================================== en-tête, arguments, données v11
rep("FPPN - Rapport actuariel et statistique, version 10 (PowerPoint, charte Adenka, Timesaver)",
    "FPPN - Rapport actuariel et statistique, version 11 (PowerPoint, charte Adenka, Timesaver)")
rep("Utilisation : python rapport_v10_fppn.py", "Utilisation : python rapport_v11_fppn.py")
rep("<carte_sensibilites.json> <donnees_v9.json> <donnees_v10.json>\n", "<carte_sensibilites.json> <donnees_v9.json> <donnees_v10.json> <donnees_v11.json> <correspondance_pbc.json>\n")
rep("RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J = sys.argv[1:12]",
    "RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J, V11J, PBCJ = sys.argv[1:14]")
rep('V9 = json.load(open(V9J, encoding="utf-8"))\n', '''V9 = json.load(open(V9J, encoding="utf-8"))
V11 = json.load(open(V11J, encoding="utf-8"))
PBCM = {int(k): v for k, v in json.load(open(PBCJ, encoding="utf-8"))["mapping"].items()}
# v11 : encours des placements lu dans les grands livres (moyenne pondérée par les jours), à la place de la moyenne des soldes de clôture
R["rendement_placements"] = V11["rendement_gl"]
''')
rep('SRC_ENT14 = "Source : entretien du 19/10/2026 (14h) avec M. Koffi, chef comptable (déclarations à confirmer par document)"',
    'SRC_ENT14 = "Source : entretien du 19/10/2026 (14h) avec M. Koffi, chef comptable"')
rep('SRC_ENT16 = "Source : entretien du 19/10/2026 (16h) avec M. Koffi, chef comptable (déclarations à confirmer par document)"',
    'SRC_ENT16 = "Source : entretien du 19/10/2026 (16h) avec M. Koffi, chef comptable"')
rep('("CMP", r"\\bCMP\\b", "Centre Médical du Policier (rapport de gestion 2023)"),', '("CMP", r"\\bCMP\\b", "Clinique médicale du policier"),')

# ================================================================== surlignage jaune : [[texte]]
rep('''def longueur(paras):
    return sum(len(p.replace("**", "")) for p in paras)''', '''def _brut(t):
    return str(t).replace("**", "").replace("[[", "").replace("]]", "")


def longueur(paras):
    return sum(len(_brut(p)) for p in paras)


JAUNE_SURL = "FFFF00"


def surligner(r):
    """Surlignage jaune d'un morceau de texte (affirmation à confirmer par une pièce ou une réponse du FPPN)."""
    rPr = r._r.get_or_add_rPr()
    for h in rPr.findall(qn("a:highlight")):
        rPr.remove(h)
    hl = rPr.makeelement(qn("a:highlight"), {})
    clr = hl.makeelement(qn("a:srgbClr"), {"val": JAUNE_SURL})
    hl.append(clr)
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addprevious(hl)
    else:
        rPr.append(hl)


def morceaux_balises(texte):
    """Découpe un texte en (morceau, gras, surligné) selon les balises **…** et [[…]]."""
    out = []
    for m1 in re.split(r"(\\[\\[.+?\\]\\])", str(texte)):
        if not m1:
            continue
        surl = m1.startswith("[[") and m1.endswith("]]")
        m1 = m1[2:-2] if surl else m1
        for m2 in re.split(r"(\\*\\*.+?\\*\\*)", m1):
            if not m2:
                continue
            gras = m2.startswith("**") and m2.endswith("**")
            out.append((m2[2:-2] if gras else m2, gras, surl))
    return out''')
rep('''    lignes = sum(math.ceil(len(p.replace("**", "")) / cpl) for p in paras)''', '''    lignes = sum(math.ceil(len(_brut(p)) / cpl) for p in paras)''')
rep('''        for morceau in re.split(r"(\\*\\*.+?\\*\\*)", bloc):
            if not morceau:
                continue
            r = p.add_run()
            gras = morceau.startswith("**") and morceau.endswith("**")
            r.text = morceau[2:-2] if gras else morceau
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = gras
            r.font.color.rgb = RGBColor.from_string(NOIR)''', '''        for morceau, gras, surl in morceaux_balises(bloc):
            r = p.add_run()
            r.text = morceau
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = gras
            r.font.color.rgb = RGBColor.from_string(NOIR)
            if surl:
                surligner(r)''')

# ================================================================== tableaux : codes couleur et tableau de signification
rep('''def hauteurs_lignes(spec, taille):''', '''VERT_C, JAUNE_C = "70AD47", "FFD966"
MOTS_COL = ("Statut", "Priorité", "Criticité", "Appréciation", "Verdict", "Horizon", "Conforme", "Reçu", "Disponible", "Utilisable", "Résolu", "État")
CODES = [  # (début de valeur en minuscules, fond, police, libellé de légende, sens)
    ("non conforme", ROUGE, "FFFFFF", "Non conforme", "norme non respectée"),
    ("non remplie", ROUGE, "FFFFFF", "Non remplie", "condition non remplie"),
    ("non démontr", ORANGE, NOIR, "Non démontré", "non démontré par les documents reçus"),
    ("non reçu", ROUGE, "FFFFFF", "Non reçu", "document non reçu"),
    ("non communiqu", ROUGE, "FFFFFF", "Non communiquée", "méthode du FPPN non transmise"),
    ("non mesur", ORANGE, NOIR, "Non mesuré", "donnée absente"),
    ("insuffisant", ROUGE, "FFFFFF", "Insuffisant", "exigence non satisfaite"),
    ("critique", ROUGE, "FFFFFF", "Critique", "non-conformité réglementaire, impact stratégique majeur, risque élevé"),
    ("urgent", ROUGE, "FFFFFF", "Urgent", "à traiter en premier"),
    ("haute", ROUGE, "FFFFFF", "Haute", "à traiter en premier"),
    ("court terme", ROUGE, "FFFFFF", "Court terme", "moins de 6 mois"),
    ("non", ROUGE, "FFFFFF", "Non", "réponse négative à la question de la colonne"),
    ("partiel", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("élevé", ORANGE, NOIR, "Élevé", "impact opérationnel significatif, risque important"),
    ("moyen terme", ORANGE, NOIR, "Moyen terme", "6 à 18 mois"),
    ("divergent", ORANGE, NOIR, "Divergent", "formule différente du texte"),
    ("au-dessus", ORANGE, NOIR, "Au-dessus", "au-dessus de la norme"),
    ("grade seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("passées seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("table seulement", ORANGE, NOIR, "Partiel", "en partie seulement"),
    ("en attente", JAUNE_C, NOIR, "En attente", "document du FPPN attendu"),
    ("à apprécier", JAUNE_C, NOIR, "À apprécier", "norme non chiffrée"),
    ("long terme", VERT_C, NOIR, "Long terme", "plus de 18 mois"),
    ("faible", VERT_C, NOIR, "Faible", "opportunité d'optimisation ou de bonne pratique"),
    ("sans texte", GRIS_C, NOIR, "Sans texte applicable", "aucun texte ne fixe la règle"),
    ("conforme", VERT_C, NOIR, "Conforme", "norme respectée"),
    ("oui", VERT_C, NOIR, "Oui", "réponse positive à la question de la colonne"),
    ("reçu", VERT_C, NOIR, "Reçu", "document reçu"),
    ("remplie", VERT_C, NOIR, "Remplie", "condition remplie"),
    ("démontré", VERT_C, NOIR, "Démontré", "démontré par les documents"),
    ("résolu", VERT_C, NOIR, "Résolu", "réponse établie par les documents"),
]


def code_couleur(entete, v):
    """Couleur d'une cellule de statut : renvoie (fond, police, libellé, sens) ou None."""
    e, t = str(entete), str(v).strip().lower()
    if not t or t in ("-", "sans objet"):
        return None
    if "criticité" in e.lower() and t.startswith("moyen"):
        return JAUNE_C, NOIR, "Moyen", "amélioration souhaitable, impact ponctuel"
    if "priorité" in e.lower() and t.startswith("faible"):
        return VERT_C, NOIR, "Faible", "à traiter en dernier"
    if t.startswith(("moyenne", "moyen")) and not t.startswith("moyen terme"):
        return ORANGE, NOIR, "Moyenne", "à traiter ensuite"
    for debut, f_, c_, lib, sens in CODES:
        if t.startswith(debut):
            return f_, c_, lib, sens
    return None


def colonnes_codees(spec):
    """Colonnes de statut d'un tableau : en-tête reconnu et au moins une valeur codée."""
    cols = []
    for j, e in enumerate(spec["entetes"]):
        if spec.get("sans_code"):
            break
        if any(m.lower() in str(e).lower() for m in MOTS_COL) or str(e).strip().endswith("?") or spec.get("statut_col") == j:
            vals = [l[j] for l in spec["lignes"] if j < len(l)]
            if vals and sum(1 for v in vals if code_couleur(e, v)) >= max(1, len(vals) // 2):
                cols.append(j)
    return cols


def legende_codes(slide, spec, cols, x0, y, largeur):
    """Tableau de signification des couleurs, sous le tableau."""
    vus = []
    for j in cols:
        for l in spec["lignes"]:
            cc = code_couleur(spec["entetes"][j], l[j]) if j < len(l) else None
            if cc and cc[2] not in [v[2] for v in vus]:
                vus.append(cc)
    if not vus:
        return y
    par_ligne = max(1, min(len(vus), int(largeur // 7.8)))
    lignes_ = [vus[k:k + par_ligne] for k in range(0, len(vus), par_ligne)]
    h = 0.42
    shp = slide.shapes.add_table(len(lignes_), 1 + 2 * par_ligne, Cm(x0), Cm(y), Cm(largeur), Cm(h * len(lignes_)))
    t = shp.table
    t._tbl.tblPr.set("firstRow", "0")
    t._tbl.tblPr.set("bandRow", "0")
    lc = (largeur - 2.2) / par_ligne
    t.columns[0].width = Cm(2.2)
    for k in range(par_ligne):
        t.columns[1 + 2 * k].width = Cm(min(2.6, lc * 0.4))
        t.columns[2 + 2 * k].width = Cm(lc - min(2.6, lc * 0.4))
    for i, lg in enumerate(lignes_):
        t.rows[i].height = Cm(h)
        cells = [("Code couleur" if i == 0 else "", "FFFFFF", NOIR, True)]
        for f_, c_, lib, sens in lg:
            cells += [(lib, f_, c_, True), (sens, "FFFFFF", NOIR, False)]
        cells += [("", "FFFFFF", NOIR, False)] * (1 + 2 * par_ligne - len(cells))
        for j, (texte, f_, c_, g_) in enumerate(cells):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.08)
            c.margin_top = c.margin_bottom = Cm(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.text = ""
            r = p.add_run()
            r.text = texte
            r.font.name = "Helvetica"
            r.font.size = Pt(7)
            r.font.bold = g_
            r.font.color.rgb = RGBColor.from_string(c_)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(f_)
    return y + h * len(lignes_)


def hauteurs_lignes(spec, taille):''')
rep('''            n = max(1, math.ceil(len(str(v)) / cpl))''', '''            n = max(1, math.ceil(len(_brut(v)) / cpl))''')
rep('''    taille = taille or spec.get("taille") or (11 if len(lignes) <= 8 else 10)
    hmax = (hauteur_max or (BAS_CONTENU - y))''', '''    taille = taille or spec.get("taille") or (11 if len(lignes) <= 8 else 10)
    hmax = (hauteur_max or (BAS_CONTENU - y))
    cols_c = colonnes_codees(spec)
    if cols_c:
        hmax -= 0.62''')
rep('''            r = p.add_run()
            r.text = str(v)
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = (i == 0) or (j == 0)
            fond, coul = (NOIR, "FFFFFF") if i == 0 else ((GRIS if i % 2 == 0 else "FFFFFF"), NOIR)
            if i > 0 and spec.get("statut_col") == j:
                for debut, f_, c_ in COUL_STATUT:
                    if str(v).startswith(debut):
                        fond, coul = f_, c_
                        r.font.bold = True
                        break
            if i > 0 and str(v).startswith("Non communiquée"):
                coul = ROUGE
                r.font.bold = True
            r.font.color.rgb = RGBColor.from_string(coul)''', '''            fond, coul = (NOIR, "FFFFFF") if i == 0 else ((GRIS if i % 2 == 0 else "FFFFFF"), NOIR)
            gras_c = (i == 0) or (j == 0)
            if i > 0 and j in cols_c:
                cc = code_couleur(spec["entetes"][j], v)
                if cc:
                    fond, coul, gras_c = cc[0], cc[1], True
            for morceau, gras, surl in morceaux_balises(v):
                r = p.add_run()
                r.text = morceau
                r.font.name = "Helvetica"
                r.font.size = Pt(taille)
                r.font.bold = gras_c or gras
                r.font.color.rgb = RGBColor.from_string(coul)
                if surl:
                    surligner(r)''')
rep('''            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(fond)
    return y + sum(hs)''', '''            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(fond)
    if cols_c:
        return legende_codes(slide, spec, cols_c, x0, y + sum(hs) * 1.07 + 0.12, sum(larg))
    return y + sum(hs)''')

# ================================================================== post-traitement : numéros de la PBC, légende du surlignage
rep('''    # Sigles de la diapositive, définis en bas de page (au-dessus de la source)''', '''    # v11 : numéros d'items de la PBC du 23/10/2026
    def _items(m):
        nums = []
        for a_, b_ in re.findall(r"(\\d+)(?: à (\\d+))?", m.group(2)):
            for k_ in range(int(a_), int(b_ or a_) + 1):
                if PBCM.get(k_) and PBCM[k_] not in nums:
                    nums.append(PBCM[k_])
        nums.sort()
        if not nums:
            return m.group(0)
        txt_ = ", ".join(str(n_) for n_ in nums[:-1]) + (" et " if len(nums) > 1 else "") + str(nums[-1])
        pre_ = m.group(1)
        if "item" in pre_:
            pre_ = "PBC, items " if len(nums) > 1 else "PBC, item "
        return pre_ + txt_
    surl_ = False
    for sh in sl.shapes:
        tfs = [sh.text_frame] if sh.has_text_frame else []
        if getattr(sh, "has_table", False) and sh.has_table:
            tfs += [c.text_frame for row in sh.table.rows for c in row.cells]
        for tf_ in tfs:
            for p_ in tf_.paragraphs:
                for r in p_.runs:
                    r.text = re.sub(r"(PBC,? items? |PBC )((?:\\d+(?: à \\d+)?(?:, | et )?)+)", _items, r.text)
                    if r._r.find(qn("a:rPr")) is not None and r._r.find(qn("a:rPr")).find(qn("a:highlight")) is not None:
                        surl_ = True
    # Sigles de la diapositive, définis en bas de page (au-dessus de la source)''')
rep('''        sig_ = sigles_du_texte(texte)
        if sig_:
            ligne_s = "Sigles : " + " | ".join(sig_)''', '''        sig_ = sigles_du_texte(texte)
        if sig_ or surl_:
            ligne_s = ("Surligné en jaune : affirmation à confirmer par une pièce ou une réponse écrite du FPPN. " if surl_ else "") + ("Sigles : " + " | ".join(sig_) if sig_ else "")''')

exec(open(os.path.join(ICI, "etape2.py"), encoding="utf-8").read())
open(os.path.join(REPO, "rapport_v11_fppn.py"), "w", encoding="utf-8").write(s)
print("rapport_v11_fppn.py écrit (étapes 1 et 2)")
