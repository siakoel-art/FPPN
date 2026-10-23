def legende_codes(slide, spec, cols, x0, y, largeur):
    """Tableau de signification des boules, sous le tableau : boule, libellé, sens. spec["legende"] remplace la légende automatique."""
    vus = [tuple(v) for v in (spec.get("legende") or [])]
    if vus:  # légende imposée : seules les boules présentes dans le tableau
        tout_ = " ".join(str(c_) for l_ in spec["lignes"] for c_ in l_)
        presentes = {cc_[0] for j in cols for l in spec["lignes"] if j < len(l) for cc_ in [code_couleur(spec["entetes"][j], l[j])] if cc_}
        vus = [v for v in vus if v[0] in tout_ or v[0] in {BOULE.get(f_, "") for f_ in presentes}]
    if not vus:
        for j in cols:
            for l in spec["lignes"]:
                cc = code_couleur(spec["entetes"][j], l[j]) if j < len(l) else None
                if cc and cc[2] not in [v[1] for v in vus]:
                    vus.append((BOULE.get(cc[0], ""), cc[2], cc[3]))
    if not vus:
        return y
    par_ligne = max(1, min(len(vus), int(largeur // 7.8)))
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


