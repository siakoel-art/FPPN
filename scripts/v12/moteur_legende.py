FAMILLES = [  # échelles complètes : la légende montre toute l'échelle, même si le tableau n'en utilise qu'une partie
    [("🔴", "Critique", "non-conformité réglementaire, impact stratégique majeur, risque élevé"), ("🟠", "Élevé", "impact opérationnel significatif, risque important"),
     ("🟡", "Moyen", "amélioration souhaitable, impact ponctuel"), ("🟢", "Faible", "opportunité d'optimisation ou de bonne pratique")],
    [("🔴", "Urgent", "à traiter en premier"), ("🟠", "Moyenne", "à traiter ensuite"), ("🟢", "Faible", "à traiter en dernier")],
    [("🔴", "Court terme", "moins de 6 mois"), ("🟠", "Moyen terme", "6 à 18 mois"), ("🟢", "Long terme", "plus de 18 mois")],
    [("🔴", "Non conforme", "norme non respectée"), ("🟠", "Non démontré", "non démontré par les documents reçus"), ("🟡", "En attente", "document du FPPN attendu"),
     ("⚪", "À apprécier", "norme non chiffrée"), ("🟢", "Conforme", "norme respectée")],
    [("🔴", "Non", "réponse négative à la question de la colonne"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Oui", "réponse positive à la question de la colonne")],
    [("🔴", "Non reçu", "document non reçu"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Reçu", "document reçu")],
    [("🔴", "Non remplie", "condition non remplie"), ("🟠", "Partiel", "en partie seulement"), ("🟢", "Remplie", "condition remplie")],
]


def legende_codes(slide, spec, cols, x0, y, largeur):
    """Tableau de signification des boules, sous le tableau : l'échelle complète de la famille de codes utilisée."""
    vus = [tuple(v) for v in (spec.get("legende") or [])]
    if not vus:
        presents = []
        for j in cols:
            for l in spec["lignes"]:
                cc = code_couleur(spec["entetes"][j], l[j]) if j < len(l) else None
                if cc and cc[2] not in [v[1] for v in presents]:
                    presents.append((BOULE.get(cc[0], ""), cc[2], cc[3]))
        libs = {v[1] for v in presents}
        fam = max(FAMILLES, key=lambda f_: sum(1 for v in f_ if v[1] in libs))
        if not any(v[1] in libs for v in fam):
            fam = FAMILLES[3]  # échelle de conformité par défaut (tableaux de formules)
        vus = list(fam) + [v for v in presents if v[1] not in {w[1] for w in fam}]
    if not vus:
        return y
    par_ligne = max(1, min(len(vus), int(largeur // 6.0)))
    lignes_ = [vus[k:k + par_ligne] for k in range(0, len(vus), par_ligne)]
    h = 0.45
    shp = slide.shapes.add_table(len(lignes_), 1 + 2 * par_ligne, Cm(x0), Cm(y), Cm(largeur), Cm(h * len(lignes_)))
    t = shp.table
    t._tbl.tblPr.set("firstRow", "0")
    t._tbl.tblPr.set("bandRow", "0")
    lc = (largeur - 2.2) / par_ligne
    t.columns[0].width = Cm(2.2)
    for k in range(par_ligne):
        t.columns[1 + 2 * k].width = Cm(0.7)
        t.columns[2 + 2 * k].width = Cm(lc - 0.7)
    for i, lg in enumerate(lignes_):
        t.rows[i].height = Cm(h)
        cells = [("Code couleur" if i == 0 else "", "", True)]
        for boule, lib, sens in lg:
            cells += [(boule, "", False), (lib, " : " + sens, True)]
        cells += [("", "", False)] * (1 + 2 * par_ligne - len(cells))
        for j, (texte, suite, g_) in enumerate(cells):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.06)
            c.margin_top = c.margin_bottom = Cm(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.text = ""
            if texte in BOULES and texte:
                p.alignment = PP_ALIGN.CENTER
            runs_texte(p, texte, 7, g_, NOIR)
            if suite:
                runs_texte(p, suite, 7, False, NOIR)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    return y + h * len(lignes_)


