VERT_C, JAUNE_C = "70AD47", "FFD966"
# ---- v12 : codes couleur par boules (pas de remplissage). Police des boules : Segoe UI Emoji (Windows), repli automatique ailleurs
BOULE = {ROUGE: "🔴", ORANGE: "🟠", JAUNE_C: "🟡", VERT_C: "🟢", GRIS_C: "⚪"}
LECT83 = {}  # texte du bandeau « Comment lire » de la carte des risques


def runs_texte(p, texte, taille, gras, coul, surl=False):
    """Écrit un texte en runs : les boules dans une police d'émojis, le reste en Helvetica."""
    for seg in re.split("([" + BOULES + "])", str(texte)):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        emoji = seg in BOULES
        r.font.name = "Segoe UI Emoji" if emoji else "Helvetica"
        r.font.size = Pt(taille + 1 if emoji else taille)
        r.font.bold = False if emoji else gras
        r.font.color.rgb = RGBColor.from_string(coul)
        if surl and not emoji:
            surligner(r)


def encadre(slide, x, y, l, h, lignes, titre="Définition et lecture"):
    """Encadré gris à filet orange : titre, puis une ligne par définition (**gras** accepté)."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(l), Cm(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string("F4F4F4")
    sh.line.color.rgb = RGBColor.from_string(ORANGE)
    sh.line.width = Pt(1.25)
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Cm(0.3)
    tf.margin_top = tf.margin_bottom = Cm(0.15)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    runs_texte(p, titre, 10.5, True, ORANGE)
    for li in lignes:
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(3)
        for morceau, gras, surl in morceaux_balises(li):
            runs_texte(p, morceau, 10, gras, NOIR, surl)
    return y + h
