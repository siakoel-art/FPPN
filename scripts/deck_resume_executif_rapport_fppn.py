"""Résumé exécutif du rapport actuariel et statistique FPPN (Timesaver Adenka, contenu sur fond blanc).

Structure : contexte et objectifs (ce qui est attendu des analyses), constats et enseignements rattachés à chaque analyse
attendue avec renvoi aux diapositives du rapport principal, constats classés du plus critique au moins critique (boules),
recommandations par axe sur le modèle du point d'étape. Chiffres du rapport actuariel et statistique du 23/10/2026.
Moteur repris du point d'étape (deck_point_etape_management_fppn.py), lui-même repris du rapport (rapport_v12_fppn.py).

Usage : python3 deck_resume_executif_rapport_fppn.py <skill adenka-powerpoint> <sortie.pptx>
"""
import os
import re
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
SKILL, SORTIE = sys.argv[1:3]
SRC = open(os.path.join(ICI, "rapport_v12_fppn.py"), encoding="utf-8").read().split("\n")


def lignes(a, b):
    return "\n".join(SRC[a - 1:b])


ENTETE = "\n".join([lignes(82, 87), lignes(90, 91), lignes(94, 94), lignes(96, 104), lignes(136, 187), lignes(238, 375), lignes(396, 411), lignes(602, 635)])
debut = next(i for i, l in enumerate(SRC, 1) if l.startswith("tmp = SORTIE.replace"))
fin = next(i for i, l in enumerate(SRC, 1) if l.startswith("finaliser(SORTIE)"))
MOTEUR = lignes(debut, fin + 1)
for avant, apres in [('"Audit Comptable et financier FPPN", "Text Placeholder 4": "23 octobre 2026"', '"Audit Comptable et financier FPPN", "Text Placeholder 4": "Novembre 2026"'),
                     ('r.text = "Rapport actuariel et statistique"', 'r.text = "Résumé exécutif du rapport actuariel et statistique"'),
                     ("def remplir_gab(slide, num, d, analyse=None):\n", "def remplir_gab(slide, num, d, analyse=None):\n    if num == 11:\n        remplir_11(slide, d)\n        return\n"),
                     ('entrees = [("Résumé exécutif", 3)]', 'entrees = ENTREES'),
                     ('        sommaire(sl, entrees)', '        sommaire_noir(sl, entrees)'),
                     ('        if x.get("triangle"):', '        if x.get("dessin"):\n            DESSINS[x["dessin"]](sl)\n        if x.get("triangle"):')]:
    assert avant in MOTEUR, avant
    MOTEUR = MOTEUR.replace(avant, apres)

ACT_AIDES = r'''
# ---------------------------------------------------------------- constats par activité : menu cliquable puis une fiche par activité
FAM_ACT = [("Garanties", ["Maladie", "Décès", "Retraités (CGRAE)"]), ("Épargne", ["EFC", "PEP"]),
            ("Services aux policiers", ["Grenier", "Flotte automobile", "MEDOC", "Clinique (CMP)"]), ("Fonctions transverses", TRANSV)]
ORDRE_ACT = [a for _, l_ in FAM_ACT for a in l_]
FAMILLE_DE = {a: f for f, l_ in FAM_ACT for a in l_}
COURT_N = {"Retraités (CGRAE)": "Retraités", "Flotte automobile": "Flotte auto", "Clinique (CMP)": "Clinique", "Cotisations et précomptes": "Cotisations",
           "Comptabilité et provisions": "Compta.", "Trésorerie et placements": "Trésorerie", "Données et informatique": "Données", "Gouvernance": "Gouvern."}
BOULE_C = {"C": "🔴", "E": "🟠", "M": "🟡", "F": "🟢"}
COUL_C = {"C": "C00000", "E": "ED7D31", "M": "FFC000", "F": "70AD47"}


def tri_items(a):
    return sorted(ACT[a], key=lambda it: (ORDRE_C[it["crit"]], it["prio"]))


def bouton(slide, x, y, l, h, nom, fond, bord=None):
    sh = cadre(slide, x, y, l, h, fond, bord, ancre=MSO_ANCHOR.MIDDLE, marge=0.12)
    sh.name = nom
    return sh


def d_menu_act(slide):
    y = 3.95
    for fam, liste in FAM_ACT:
        b = cadre(slide, 1.4, y, 31.2, 0.5, BLEU, ancre=MSO_ANCHOR.MIDDLE, marge=0.15)
        run(b.text_frame.paragraphs[0], fam, 9.5, True, BLANC)
        y += 0.58
        n_ = len(liste)
        g_ = 0.2
        l_ = (31.2 - (n_ - 1) * g_) / n_
        for k, a in enumerate(liste):
            its = tri_items(a)
            sh = bouton(slide, 1.4 + k * (l_ + g_), y, l_, 1.62, "Nav " + a, BLANC, "112A40")
            p = sh.text_frame.paragraphs[0]
            run(p, "● ", 12, True, COUL_C[its[0]["crit"]])
            run(p, a, 10.5, True, NOIR)
            p1 = sh.text_frame.add_paragraph()
            run(p1, f"{len(its)} constats  ›  ouvrir la fiche", 8, True, "53565A")
            p2 = sh.text_frame.add_paragraph()
            run(p2, f"Le plus grave : {its[0]['num']}, {(its[0]['titre'] if its[0]['titre'][1:2].isupper() else its[0]['titre'][0].lower() + its[0]['titre'][1:])}", 8, False, "53565A")
        y += 1.72
    tb = slide.shapes.add_textbox(Cm(1.4), Cm(y + 0.1), Cm(31.2), Cm(0.5))
    p = tb.text_frame.paragraphs[0]
    run(p, "Comment lire : ", 9, True, NOIR)
    run(p, "en diaporama, cliquer sur une activité pour ouvrir sa fiche. Les fiches se lisent aussi dans l'ordre, de la maladie aux données. Boule : criticité du constat le plus grave.", 9, False, NOIR)


def d_nav(slide, act):
    k = ORDRE_ACT.index(act)
    y, h = 3.85, 0.72
    sh = bouton(slide, 1.4, y, 2.6, h, "Nav MENU", BLEU)
    run(sh.text_frame.paragraphs[0], "⌂ Toutes les activités", 7.5, True, BLANC)
    x = 4.1
    if k > 0:
        sh = bouton(slide, x, y, 0.95, h, "Nav " + ORDRE_ACT[k - 1], ORANGE)
        sh.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run(sh.text_frame.paragraphs[0], "◀", 10, True, NOIR)
    x += 1.03
    if k < len(ORDRE_ACT) - 1:
        sh = bouton(slide, x, y, 0.95, h, "Nav " + ORDRE_ACT[k + 1], ORANGE)
        sh.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run(sh.text_frame.paragraphs[0], "▶", 10, True, NOIR)
    x0 = x + 1.08
    n_ = len(ORDRE_ACT)
    g_ = 0.08
    l_ = (32.6 - x0 - (n_ - 1) * g_) / n_
    for j, a in enumerate(ORDRE_ACT):
        courant = a == act
        sh = bouton(slide, x0 + j * (l_ + g_), y, l_, h, ("Ici " if courant else "Nav ") + a, NOIR if courant else BLANC, None if courant else "BBBCBC")
        sh.text_frame.margin_left = sh.text_frame.margin_right = Cm(0.04)
        p = sh.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, "●", 6.5, True, COUL_C[tri_items(a)[0]["crit"]])
        run(p, " " + COURT_N.get(a, a), 6.5, True, BLANC if courant else NOIR)


DESSINS["menu_act"] = d_menu_act
'''
ACT_DIAPOS = r'''
_ntot = sum(len(ACT[a]) for a in ORDRE_ACT)
contenu({"titre": "Constats par activité", "sous_titre": f"{_ntot} constats sur 14 activités : l'épargne, la liquidité et les règles de solidité concentrent les plus graves",
         "synthese": "Chaque bouton ouvre la fiche de l'activité : couleur du constat le plus grave, nombre de constats et constat le plus grave"},
        "Menu des fiches. En diaporama, cliquer sur une activité, puis revenir par « Toutes les activités ».",
        dessin="menu_act", sommaire="Constats par activité", source=SRC_R + " ; résumé exécutif du 26/10/2026")
for a in ORDRE_ACT:
    its = tri_items(a)
    nb = {c: sum(1 for it in its if it["crit"] == c) for c in "CEMF"}
    det = ", ".join(f"{nb[c]} {lib}" for c, lib in (("C", "critique" + ("s" if nb["C"] > 1 else "")), ("E", "élevé" + ("s" if nb["E"] > 1 else "")),
                                                    ("M", "moyen" + ("s" if nb["M"] > 1 else "")), ("F", "faible" + ("s" if nb["F"] > 1 else ""))) if nb[c])
    contenu({"titre": f"{FAMILLE_DE[a]} : {a}", "sous_titre": SOUS[a],
             "synthese": f"{len(its)} constats ({det}). Le plus grave : {its[0]['num']}, {(its[0]['titre'] if its[0]['titre'][1:2].isupper() else its[0]['titre'][0].lower() + its[0]['titre'][1:]).rstrip('.')}"},
            "Lire chaque ligne de gauche à droite : le constat et son commentaire, la recommandation, la complexité, la priorité, puis la criticité. Les boutons du haut mènent aux autres activités.",
            tableau={"titre": None, "entetes": ["N°", "Constat et commentaire", "Recommandation", "Complexité", "Priorité", "Criticité"], "largeurs": [1.3, 15.4, 8.9, 2.1, 1.75, 1.75],
                     "taille": 8.5, "sans_code": True, "legende": LEG_CRIT,
                     "lignes": [[it["num"], f"**{it['titre']}** {it['texte']}", it["reco"], it["cplx"], it["prio"], BOULE_C[it["crit"]]] for it in its]},
            nav=a, source=SRC_R + " ; résumé exécutif du 26/10/2026")
'''
CONTENU = r'''
import copy
VALIDITE = {}
PBCM = {}
GAB = {"couverture": 1, "sommaire": 4, "contenu": 5, "merci": 110}
BLANC, GRIS_CLAIR, CARTE = "FFFFFF", "BBBCBC", "F2F2F2"
PRIO = {"P1": "C00000", "P2": "ED7D31", "P3": "FFC000", "P4": "70AD47"}
RAP = "Source : rapport actuariel et statistique du 23/10/2026"


def run(p, texte, taille, gras=False, couleur=NOIR, surl=False):
    r_ = p.add_run()
    r_.text = texte
    r_.font.name = "Helvetica"
    r_.font.size = Pt(taille)
    r_.font.bold = gras
    r_.font.color.rgb = RGBColor.from_string(couleur)
    if surl:
        surligner(r_)
    return r_


def balise(p, texte, taille, couleur=NOIR):
    for morceau, g_, s_ in morceaux_balises(texte):
        run(p, morceau, taille, g_, couleur, s_)


def cadre(slide, x, y, l, h, fond=None, bord=None, ancre=MSO_ANCHOR.TOP, marge=0.15, forme=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(forme, Cm(x), Cm(y), Cm(l), Cm(h))
    if fond:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fond)
    else:
        sh.fill.background()
    if bord:
        sh.line.color.rgb = RGBColor.from_string(bord)
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Cm(marge)
    tf.margin_top = tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = ancre
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    return sh


def icone_noire(blob):
    from PIL import Image
    im = Image.open(io.BytesIO(blob)).convert("RGBA")
    alpha = im.split()[3]
    if alpha.getextrema() == (255, 255):
        alpha = im.convert("L").point(lambda v: 255 if v < 128 else 0)
    noir = Image.new("RGBA", im.size, (6, 6, 6, 0))
    noir.putalpha(alpha)
    out = io.BytesIO()
    noir.save(out, "PNG")
    return out.getvalue()


def icone(slide, k, x, y, t, blanche=True):
    ic = icones[k % len(icones)]
    blob = icone_blanche(ic.image.blob) if blanche else icone_noire(ic.image.blob)
    slide.shapes.add_picture(io.BytesIO(blob), Cm(x), Cm(y), Cm(t), Cm(t))


def sommaire_noir(slide, entrees):
    """Sommaire sur fond noir, au modèle fourni : traits blancs, disques blancs à icône noire, cadres orange."""
    txt(slide, 1.0, 0.55, 12, 1.1, "Sommaire", 20, True, BLANC)
    icone(slide, 0, 2.15, 1.55, 1.1)
    pas = 1.45
    centres = [4.2 + i * pas for i in range(len(entrees))]
    trait = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.7), Cm(2.75), Cm(2.7), Cm(centres[-1]))
    trait.line.color.rgb = RGBColor.from_string(BLANC)
    trait.line.width = Pt(1)
    for i, (intitule, num) in enumerate(entrees):
        cy = centres[i]
        h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(2.7), Cm(cy), Cm(4.75), Cm(cy))
        h.line.color.rgb = RGBColor.from_string(BLANC)
        h.line.width = Pt(1)
        d = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(4.75), Cm(cy - 0.62), Cm(1.24), Cm(1.24))
        d.fill.solid()
        d.fill.fore_color.rgb = RGBColor.from_string(BLANC)
        d.line.fill.background()
        icone(slide, i + 3, 4.75 + 0.27, cy - 0.35, 0.7, blanche=False)
        for xx, l, t_, al in ((6.25, 23.0, intitule, PP_ALIGN.LEFT), (29.55, 2.5, str(num), PP_ALIGN.CENTER)):
            b = cadre(slide, xx, cy - 0.5, l, 1.0, None, ORANGE, MSO_ANCHOR.MIDDLE, 0.4)
            p = b.text_frame.paragraphs[0]
            p.alignment = al
            run(p, t_, 13, True, BLANC)


# ---------------------------------------------------------------- dessins des diapositives
def d_contexte(slide):
    cartes = [("Le FPPN", "La mutuelle des policiers de Côte d'Ivoire. Les cotisations sont prélevées sur la solde, ou sur la pension pour les retraités. Il rembourse les soins, verse une allocation au décès, gère deux épargnes (EFC et PEP) et prête aux policiers par le Grenier."),
              ("Son cadre", "Mutualité sociale de l'UEMOA, sous le contrôle de l'AIRMS. Il doit respecter 36 indicateurs de solidité (Décision n° 05/2022). Le Code CIMA ne sert que de référence technique."),
              ("La mission Adenka", "Audit comptable et financier du FPPN du 01/01/2020 au 31/08/2026, sur la base d'une offre technique validée par le FPPN. Réunions de lancement les 13 et 14/10/2026."),
              ("Le volet actuariel et statistique", "Vérifier ce que le FPPN devra payer à ses adhérents et s'il pourra le payer à temps. La valorisation et le rendement des placements relèvent aussi de l'expert financier.")]
    x0, y0, l, h, g = 1.4, 3.95, 7.65, 8.2, 0.2
    for k, (t_, c_) in enumerate(cartes):
        x = x0 + k * (l + g)
        b = cadre(slide, x, y0, l, 1.0, BLEU, ancre=MSO_ANCHOR.MIDDLE, marge=1.2)
        run(b.text_frame.paragraphs[0], t_, 11.5, True, BLANC)
        icone(slide, k + 4, x + 0.25, y0 + 0.15, 0.7)
        b = cadre(slide, x, y0 + 1.05, l, h - 1.05, CARTE, ancre=MSO_ANCHOR.TOP, marge=0.3)
        b.text_frame.margin_top = Cm(0.3)
        run(b.text_frame.paragraphs[0], c_, 12.5)
    chiffres = [("3,4 Md", "de FCFA de cotisations en 2025"), ("15,7 Md", "de FCFA d'épargne due au 31/08/2026"),
                ("31 198", "policiers cotisants en août 2026"), ("610", "retraités cotisants en août 2026")]
    for k, (v_, l_) in enumerate(chiffres):
        x = x0 + k * (l + g)
        b = cadre(slide, x, y0 + h + 0.3, l, 2.0, None, ORANGE, ancre=MSO_ANCHOR.MIDDLE)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, v_, 20, True, ORANGE)
        p2 = b.text_frame.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run(p2, l_, 10, False, NOIR)


def d_objectifs(slide):
    x0, y0 = 1.4, 3.95
    b = cadre(slide, x0, y0, 18.6, 0.8, BLEU, ancre=MSO_ANCHOR.MIDDLE)
    run(b.text_frame.paragraphs[0], "Les quatre questions de la mission", 11.5, True, BLANC)
    qs = [("ENG", "Engagements", "Combien le FPPN doit-il à ses adhérents, aujourd'hui et plus tard ?"),
          ("PROV", "Provisions", "Les sommes mises de côté suffisent-elles à payer les prestations déjà nées ?"),
          ("ALM", "Adossement", "L'argent du FPPN sera-t-il disponible au moment où il devra payer ?"),
          ("PLA", "Placements", "Les placements sont-ils sûrs, disponibles et assez rentables ?")]
    for k, (code, t_, q_) in enumerate(qs):
        y = y0 + 0.95 + k * 2.15
        b = cadre(slide, x0, y, 2.6, 2.0, ORANGE, ancre=MSO_ANCHOR.MIDDLE)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, code, 14, True, NOIR)
        b = cadre(slide, x0 + 2.7, y, 15.9, 2.0, CARTE, ancre=MSO_ANCHOR.MIDDLE, marge=0.3)
        run(b.text_frame.paragraphs[0], t_, 12, True)
        p2 = b.text_frame.add_paragraph()
        run(p2, q_, 11)
    x1, l1 = x0 + 19.0, 12.2
    b = cadre(slide, x1, y0, l1, 0.8, BLEU, ancre=MSO_ANCHOR.MIDDLE)
    run(b.text_frame.paragraphs[0], "Travaux ajoutés pour y répondre", 11.5, True, BLANC)
    aj = [("DON", "contrôle des données de base (offre technique, pôle 4)"), ("STAT", "volet statistique : effectifs, cotisations, soins, décès, mortalité"),
          ("REG", "36 indicateurs de solidité de l'UEMOA et gouvernance (pôle 4, risque R3)"), ("PROV", "autres provisions : prêts du Grenier, retraite du personnel"),
          ("ALM", "projection de l'équilibre technique (pôle 4)")]
    b = cadre(slide, x1, y0 + 0.95, l1, 6.2, CARTE, marge=0.3)
    tf = b.text_frame
    tf.margin_top = Cm(0.25)
    for k, (c_, t_) in enumerate(aj):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(7)
        run(p, c_ + "  ", 11, True, BLEU)
        run(p, t_, 11)
    b = cadre(slide, x1, y0 + 7.3, l1, 1.9, None, ORANGE, ancre=MSO_ANCHOR.MIDDLE, marge=0.3)
    p = b.text_frame.paragraphs[0]
    run(p, "Période : ", 11, True, ORANGE)
    run(p, "du 01/01/2020 au 31/08/2026, à chaque clôture annuelle et au 31/08/2026.", 11, False, NOIR)
    p2 = b.text_frame.add_paragraph()
    run(p2, "Codes : ", 11, True, ORANGE)
    run(p2, "repris dans le récapitulatif et les recommandations.", 11, False, NOIR)


def _texte_p(p_el, texte):
    rs = p_el.findall(qn("a:r"))
    for r_ in rs[1:]:
        p_el.remove(r_)
    if rs:
        rs[0].find(qn("a:t")).text = texte
    return p_el


def remplir_11(slide, d):
    """Deux cartes du Timesaver : en-tête noir (Actuariat) et orange (Stratégie), sous-titres en gras et puces."""
    formes = {sh.name: sh for sh in slide.shapes}
    for nom_t, nom_c, (titre, sections) in (("Rectangle 3", "Rectangle 20", d["cartes"][0]), ("Rectangle 16", "Rectangle 23", d["cartes"][1])):
        ecrire(formes[nom_t], [titre], [13], [True])
        corps = formes[nom_c]
        corps.height = Cm(7.4)
        txb = corps.text_frame._txBody
        paras = txb.findall(qn("a:p"))
        mod_t, mod_p = copy.deepcopy(paras[0]), copy.deepcopy(paras[1])
        for p_ in paras:
            txb.remove(p_)
        for sous, puces in sections:
            txb.append(_texte_p(copy.deepcopy(mod_t), sous))
            for x_ in puces:
                txb.append(_texte_p(copy.deepcopy(mod_p), x_))
        for p_ in corps.text_frame.paragraphs:
            for r_ in p_.runs:
                r_.font.size = Pt(10.5 if r_.font.bold else 10)


AXES = [("1", "Engagements et épargne", "ENG", [("Fixer le taux servi et comptabiliser les intérêts dus", "P1"), ("Tenir un compte par épargnant", "P2")]),
        ("2", "Provisions et comptes", "PROV", [("Calculer les provisions chaque année et reclasser l'excédent", "P1"), ("Déprécier les prêts du Grenier", "P3")]),
        ("3", "Liquidité et placements", "ALM  PLA", [("Constituer une réserve de liquidité et établir l'échéancier de l'épargne", "P1"), ("Voter une politique de placement", "P2")]),
        ("4", "Conformité et gouvernance", "REG", [("Présenter un plan de conformité à l'AIRMS", "P1"), ("Délibérer en conseil sur chaque placement", "P2")]),
        ("5", "Données et pilotage", "DON  STAT", [("Transmettre les données manquantes", "P1"), ("Rapprocher chaque mois les cotisations", "P2")])]


def bande_volet(slide, y, h, libelle, fond):
    b = cadre(slide, 1.4, y, 1.1, h, fond, ancre=MSO_ANCHOR.MIDDLE, marge=0.02)
    b.text_frame.word_wrap = True
    bp = b._element.find(qn("p:txBody")).find(qn("a:bodyPr"))
    bp.set("vert", "vert270")
    p = b.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run(p, libelle, 11, True, BLANC if fond != ORANGE else NOIR)


def d_recos(slide):
    x0, g = 2.7, 0.18
    l = (31.2 - 1.3 - 4 * g) / 5
    # bande actuariat
    y0, h_b = 3.95, 6.35
    bande_volet(slide, y0, h_b, "VOLET ACTUARIAT", NOIR)
    for k, (num, titre, codes, actions) in enumerate(AXES):
        x = x0 + k * (l + g)
        b = cadre(slide, x, y0, l, 1.35, BLEU, ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
        p = b.text_frame.paragraphs[0]
        run(p, f"Axe {num}  ", 9.5, True, ORANGE)
        run(p, codes, 8.5, False, GRIS_CLAIR)
        p2 = b.text_frame.add_paragraph()
        run(p2, titre, 11, True, BLANC)
        y = y0 + 1.45
        for t_, pr in actions:
            b = cadre(slide, x, y, l, 2.38, CARTE, ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
            b.text_frame.margin_top = Cm(0.5)
            run(b.text_frame.paragraphs[0], t_, 10.5)
            pil = cadre(slide, x + l - 1.2, y + 0.12, 1.0, 0.45, PRIO[pr], ancre=MSO_ANCHOR.MIDDLE, marge=0.02, forme=MSO_SHAPE.ROUNDED_RECTANGLE)
            q = pil.text_frame.paragraphs[0]
            q.alignment = PP_ALIGN.CENTER
            run(q, pr, 8.5, True, BLANC)
            y += 2.45
    # bande stratégie, à renseigner par l'équipe
    y1 = y0 + h_b + 0.3
    bande_volet(slide, y1, 5.3, "VOLET STRATÉGIE", ORANGE)
    for k in range(5):
        x = x0 + k * (l + g)
        b = cadre(slide, x, y1, l, 1.35, ORANGE, ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
        p = b.text_frame.paragraphs[0]
        run(p, f"Axe {k + 1}", 9.5, True, NOIR)
        p2 = b.text_frame.add_paragraph()
        run(p2, "Intitulé de l'axe", 11, True, NOIR)
        b = cadre(slide, x, y1 + 1.45, l, 3.85, None, "BBBCBC", ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, "Actions et priorités : à renseigner par l'équipe Stratégie", 10, False, "808080")
    tb = slide.shapes.add_textbox(Cm(1.4), Cm(y1 + 5.35), Cm(31.2), Cm(0.42))
    p = tb.text_frame.paragraphs[0]
    run(p, "Priorité : ", 8.5, True, NOIR)
    for pr, lib in (("P1", "dans les 3 mois"), ("P2", "dans les 6 mois"), ("P3", "de 6 à 18 mois"), ("P4", "au-delà de 18 mois")):
        run(p, "■ ", 10, True, PRIO[pr])
        run(p, f"{pr} {lib}     ", 8.5, False, NOIR)


ETAPES = [("Transmission des données manquantes", "FPPN", "Dès que possible"), ("Calculs finaux : provisions, épargne par policier, projection", "Adenka", "À réception"),
          ("Rapprochement avec l'expert financier", "Adenka", "Avant la restitution"), ("Restitution du rapport provisoire", "Adenka et FPPN", "23 et 24/11/2026"),
          ("Intégration des commentaires du FPPN", "Adenka", "Du 25 au 27/11/2026"), ("Remise du rapport définitif", "Adenka", "30/11/2026")]


def d_etapes(slide):
    x0, n = 2.7, len(ETAPES)
    pas = (31.2 - 1.3) / n
    y0 = 3.95
    # couloir actuariat
    bande_volet(slide, y0, 5.6, "ACTUARIAT", NOIR)
    ligne = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x0 + pas / 2), Cm(y0 + 0.5), Cm(x0 + pas * (n - 0.5)), Cm(y0 + 0.5))
    ligne.line.color.rgb = RGBColor.from_string(ORANGE)
    ligne.line.width = Pt(2)
    for k, (t_, qui, quand) in enumerate(ETAPES):
        cx = x0 + pas * (k + 0.5)
        d = cadre(slide, cx - 0.5, y0, 1.0, 1.0, ORANGE, ancre=MSO_ANCHOR.MIDDLE, marge=0.02, forme=MSO_SHAPE.OVAL)
        q = d.text_frame.paragraphs[0]
        q.alignment = PP_ALIGN.CENTER
        run(q, f"{k + 1:02d}", 11, True, NOIR)
        b = cadre(slide, cx - pas / 2 + 0.08, y0 + 1.1, pas - 0.16, 0.8, None, None, ancre=MSO_ANCHOR.MIDDLE, marge=0.03)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, quand, 11, True, BLEU)
        b = cadre(slide, cx - pas / 2 + 0.08, y0 + 2.0, pas - 0.16, 3.6, CARTE, ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, t_, 10.5, True)
        p2 = b.text_frame.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run(p2, qui, 9.5, False, "53565A")
    # couloir stratégie, à renseigner par l'équipe
    y1 = y0 + 5.9
    bande_volet(slide, y1, 3.3, "STRATÉGIE", ORANGE)
    for k in range(n):
        cx = x0 + pas * (k + 0.5)
        b = cadre(slide, cx - pas / 2 + 0.08, y1, pas - 0.16, 3.3, None, "BBBCBC", ancre=MSO_ANCHOR.MIDDLE, marge=0.2)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, "Étape, responsable et date : à renseigner par l'équipe Stratégie", 9.5, False, "808080")
    b = cadre(slide, 1.4, y1 + 3.5, 31.2, 1.1, None, ORANGE, ancre=MSO_ANCHOR.MIDDLE, marge=0.4)
    p = b.text_frame.paragraphs[0]
    run(p, "Point d'attention, volet actuariat : ", 10.5, True, ORANGE)
    run(p, "le calendrier dépend des données du FPPN. Sans elles, la restitution du 23/11/2026 se fera avec des verdicts partiels sur les engagements et les provisions. Avancement au 23/10/2026 : 63,85 %.", 10.5, False, NOIR)


DESSINS = {"contexte": d_contexte, "objectifs": d_objectifs, "recos": d_recos, "etapes": d_etapes}

# ---------------------------------------------------------------- récapitulatif codifié
RECAP_DEMANDES = [
    ["🔴 ENG-1", "Demande 1 : épargne (EFC, PEP)", "Inventaire des soldes comptables, intérêts simples sur l'encours moyen aux taux des documents du FPPN, sensibilité au taux",
     "Soldes de 7 exercices (2020 au 31/08/2026), 3 taux comparés (3,50 %, 5,50 %, 14 %)",
     "**La dette d'épargne est sous-évaluée de 3,3 milliards de FCFA.** Épargne due : 15 743 516 391 FCFA. Intérêts promis jamais comptabilisés : 3 278 913 238 FCFA (2 541 799 542 FCFA si l'EFC est servie à 3,50 %). Engagement réel : 19 022 429 629 FCFA."],
    ["🟠 ENG-2", "Demande 1 : décès et retraités", "États de fin d'année rapprochés du compte 41591000, décès déclarés après la clôture",
     "5 états annuels de décès (2021 à 2025), fichiers de 610 retraités",
     "**La dette envers les familles grossit, et la santé des retraités n'est pas mesurable.** Allocations décès dues : 489 419 669 FCFA, en hausse chaque année. Décès déclarés en retard : 7 100 000 FCFA. Date de naissance connue pour 105 retraités seulement."],
    ["🟠 PROV-1", "Demande 2 : provisions maladie et décès", "Règlements postérieurs à la clôture (écoulement mensuel), cadence des déclarations tardives (Code CIMA, art. 334-12). Chain Ladder écarté : soins réglés en quelques mois",
     "Journal des sinistres de 3 grands livres (2023 à 2025), 5 états décès, délais de paiement de 2020 à 2026",
     "**Les provisions sont figées et mal affectées, pas insuffisantes.** 4 829 298 290 FCFA au bilan depuis 2020, montant passé vers 2004 selon la direction comptable, pour un besoin mesuré de 19 581 392 FCFA. Rien n'est provisionné pour les intérêts de l'épargne."],
    ["🟢 PROV-3", "Demande 2 : prêts du Grenier", "Dépréciation par tranche de retard, aux taux de la politique de crédit (40 %, 80 %, 100 %)",
     "Balances âgées de 2 959 prêts au 31/12/2025 et au 31/08/2026",
     "**Les pertes probables sur prêts ne sont pas comptées.** 14 016 392 FCFA à déprécier, aucune dépréciation depuis 2022. Montant faible : le vrai risque est l'épargne prêtée, le FPPN finançant 87,66 % de l'encours."],
    ["🟢 PROV-2", "Demande 2 : retraite du personnel", "Unités de crédit projetées (méthode actuarielle), comparée aux droits acquis",
     "Calcul agent par agent à partir des déclarations de salaires (DISA) de 2021 à 2025, 79 agents en 2025",
     "**La provision du personnel est suffisante, et la méthode actuarielle est la plus juste.** Recalcul : 17 184 450 FCFA avec un départ à 60 ans (ordonnance n° 2012-03), pour 43 666 678 FCFA au compte du FPPN (droits acquis : 43 195 084 FCFA). Elle tient compte du départ possible avant la retraite, des hausses de salaire et du temps, et donne le coût juste de chaque année."],
    ["🔴 ALM-1", "Demande 3 : adossement actif-passif", "Classement des actifs par liquidité, écoulement de l'épargne au taux observé, écart de liquidité par échéance, tests de résistance, scénarios de couverture",
     "Actifs et dettes de 2 bilans classés par échéance, sorties d'épargne de 5 années, 3 scénarios (central, choc, choc sévère)",
     "**L'argent disponible ne couvre qu'un tiers des paiements de l'année.** Couverture de 34,42 % (26,22 % en choc sévère), manque de 12 596 052 822 FCFA à cinq ans. Couverture des dettes : 96,33 % sur le papier, 27,19 % sans les biens non évalués."],
    ["🔴 PLA-1", "Demande 4 : placements", "Encours pondéré par les jours, rendement, comparaison aux taux de la BCEAO, quatre conditions de l'art. 17",
     "Encours jour par jour de 4 placements dans les grands livres de 2021 à 2026",
     "**Les placements sont trop faibles pour payer les intérêts promis.** 763 752 135 FCFA placés en 2025, soit 5,53 % de l'épargne, pour un rendement de 4,90 %. Écart de 470 608 168 FCFA avec les intérêts dus en 2025. 2 356 399 123 FCFA dorment sur des comptes courants."],
]
RECAP_COMPL = [
    ["🔴 REG-1", "Indicateurs de solidité", "36 ratios de la Décision n° 05/2022, deux lectures (comme le FPPN et corrigée selon le PCMS), sensibilités et leviers",
     "36 ratios recalculés sur 7 exercices, en deux lectures, 9 tableaux de sensibilité",
     "**Le FPPN ne respecte pas 14 règles de solidité sur 36, ce qui expose son agrément.** 7 conformes, 15 en attente de pièces. Fonds propres : -879 498 700 FCFA. Liquidité : 0,13 (1,15 en lecture corrigée). Frais : 36,34 % des cotisations pour 20 % au plus."],
    ["🟠 STAT-2", "Sinistralité", "Taux de sinistralité réglementaire (Décision n° 05/2022), par garantie, ratio combiné, excédent par risque",
     "Prestations et cotisations de 7 exercices, par garantie",
     "**La garantie maladie perd de l'argent, et le décès la finance.** Soins : 6,00 % des cotisations maladie en 2021, 73,53 % en 2025, 97,40 % en 2026 (déficit de 702 millions). Le taux global de 75,20 % est dans la norme mais masque ce déficit."],
    ["🟠 REG-2", "Gouvernance", "Entretiens, logigrammes du FPPN, lecture des statuts et des décisions",
     "4 entretiens, 4 circuits redessinés à partir des logigrammes du FPPN",
     "**Les décisions qui engagent l'épargne ne sont pas encadrées.** Placements décidés par la direction générale. Projet de Cocody sans budget. Ni politique de placement, ni comité actif-passif. Clinique sans comptabilité séparée."],
    ["🟠 DON-1", "Données de base", "Extraction automatisée, empreinte numérique de chaque fichier, contrôle du format, des doublons, des lignes de total et des fichiers vides, rapprochement des fichiers envoyés et reçus",
     "4 287 fichiers lus, 4 454 416 lignes de cotisations contrôlées, 506 feuilles, 80 mois couverts",
     "**Des données essentielles manquent pour finir les calculs.** Cotisations mensuelles disponibles seulement depuis 2023 ou 2024. Fichiers de décembre 2024 vides. Ni base des prestations datée, ni dates de naissance, ni fichier envoyé à la Solde."],
    ["🟡 STAT-1", "Effectifs et cotisations", "Comptage mensuel des cotisants, mesure du taux réellement prélevé, tendance par régression",
     "Fichiers mensuels de 31 mois (janvier 2024 à août 2026), salaires de 8 mois",
     "**Les effectifs progressent, mais la répartition des cotisations diffère des statuts.** 23 119 cotisants en janvier 2024, 31 198 en août 2026 (+34,95 %). Taux prélevés : 3,77 % et 1,23 % au lieu de 3,50 % et 1,50 %. Cotisations : +45,76 % de 2020 à 2025."],
    ["🟡 STAT-3", "Mortalité et projection", "Tables CIMA, projection agrégée de l'équilibre en deux scénarios",
     "4 tables de mortalité, projection de 2026 à 2030 en 2 scénarios",
     "**Les calculs par âge restent à faire.** La mortalité par âge est impossible sans dates de naissance. La projection de l'équilibre n'est que globale : elle sera affinée à réception des données."],
]
LEG_CRIT = [("🔴", "Critique", "risque élevé ou non-conformité"), ("🟠", "Élevé", "impact significatif"), ("🟡", "Moyen", "amélioration souhaitable"), ("🟢", "Faible", "point maîtrisé ou de faible montant")]

# ---------------------------------------------------------------- recommandations par axe (modèle validé le 02/11/2026)
AXES_R = [("1", "Engagements et épargne", "ENG", [("Fixer par délibération le taux servi sur l'épargne", "P1"), ("Comptabiliser les intérêts dus et tenir un compte par épargnant", "P1"),
                                                   ("Étude actuarielle des risques longs tous les cinq ans", "P4")]),
          ("2", "Provisions et comptes", "PROV", [("Note de calcul annuelle des provisions, signée et présentée à l'Assemblée générale", "P1"),
                                                  ("Reclasser l'excédent après décision des organes et avis de l'AIRMS", "P1"), ("Déprécier les prêts du Grenier, comptabilité séparée de la clinique", "P2")]),
          ("3", "Liquidité et placements", "ALM  PLA", [("Réserve de liquidité égale aux sorties d'épargne d'une année", "P1"), ("Échéancier des sorties d'épargne (départs à la retraite)", "P1"),
                                                        ("Politique de placement votée, évaluation de la clinique et des terrains", "P2")]),
          ("4", "Conformité et gouvernance", "REG", [("Plan de mise en conformité chiffré présenté à l'AIRMS", "P1"), ("Délibération du conseil pour chaque placement, comité actif-passif", "P2"),
                                                     ("Recalculer la cotisation maladie, contrôler 5 % des factures", "P2")]),
          ("5", "Données et pilotage", "DON  STAT", [("Transmettre les données manquantes (dates de naissance, prestations datées)", "P1"),
                                                     ("Rapprocher chaque mois les fichiers de cotisations envoyés et reçus", "P2"), ("Tableau de bord semestriel des indicateurs de solidité", "P3")])]


def d_recos_rapport(slide):
    x0, g = 1.4, 0.18
    l = (31.2 - 4 * g) / 5
    y0, h_t, h_c, h_k, g_k = 3.95, 1.45, 0.6, 2.45, 0.1
    for k, (num, titre, codes, actions) in enumerate(AXES_R):
        x = x0 + k * (l + g)
        b = cadre(slide, x, y0, l, h_t, NOIR, ancre=MSO_ANCHOR.MIDDLE, marge=0.25)
        p = b.text_frame.paragraphs[0]
        run(p, f"Axe {num}", 9.5, True, ORANGE)
        p2 = b.text_frame.add_paragraph()
        run(p2, titre, 11, True, BLANC)
        b = cadre(slide, x, y0 + h_t + 0.05, l, h_c, ORANGE, ancre=MSO_ANCHOR.MIDDLE, marge=0.25)
        run(b.text_frame.paragraphs[0], "Constats : " + codes, 9, True, NOIR)
        y = y0 + h_t + 0.05 + h_c + 0.1
        for t_, pr in actions:
            b = cadre(slide, x, y, l, h_k, CARTE, ancre=MSO_ANCHOR.MIDDLE, marge=0.25)
            b.text_frame.margin_top = Cm(0.55)
            run(b.text_frame.paragraphs[0], t_, 10.5)
            pil = cadre(slide, x + l - 1.2, y + 0.15, 1.0, 0.45, PRIO[pr], ancre=MSO_ANCHOR.MIDDLE, marge=0.02, forme=MSO_SHAPE.ROUNDED_RECTANGLE)
            q = pil.text_frame.paragraphs[0]
            q.alignment = PP_ALIGN.CENTER
            run(q, pr, 8.5, True, BLANC)
            y += h_k + g_k
    tb = slide.shapes.add_textbox(Cm(2.2), Cm(y + 0.05), Cm(8.0), Cm(2.3))
    tf = tb.text_frame
    run(tf.paragraphs[0], "Priorité :", 8.5, True, NOIR)
    for pr, lib in (("P1", "dans les 3 mois"), ("P2", "dans les 6 mois"), ("P3", "de 6 à 18 mois"), ("P4", "au-delà de 18 mois")):
        q = tf.add_paragraph()
        run(q, "■ ", 9, True, PRIO[pr])
        run(q, f"{pr} {lib}", 8.5, False, NOIR)


DESSINS["recos_rapport"] = d_recos_rapport
#@@HELPERS_ACTIVITES@@

# ---------------------------------------------------------------- constats et enseignements par analyse attendue
VERD = {"🔴": "Insuffisant", "🟠": "Partiel", "🟡": "Moyen", "🟢": "Satisfaisant"}
ENS_DEMANDES = [
    ["ENG", "Demande 1 : engagements envers les adhérents et ayants droit",
     "Épargne due au 31/08/2026 : 15 743 516 391 FCFA. Intérêts promis jamais comptabilisés : 3 278 913 238 FCFA (2 541 799 542 FCFA si l'EFC est servie à 3,50 %). Allocations décès dues : 489 419 669 FCFA. Provision du personnel recalculée : 17 184 450 FCFA. Santé des retraités non mesurable.",
     "**L'épargne est le premier engagement, et sa valeur n'est pas démontrée.** Aucun taux n'est arrêté et les intérêts ne sont pas comptabilisés. Le verdict deviendra « Démontré » avec le taux servi, les relevés individuels et les âges des retraités.",
     "🟠", "5, 95 à 98, 112, 126, 133"],
    ["PROV", "Demande 2 : suffisance des provisions au regard des prestations servies",
     "Provisions maladie et obsèques : 4 829 298 290 FCFA, inchangées depuis 2020. Besoin mesuré : 19 581 392 FCFA. Intérêts d'épargne (3 278 913 238 FCFA) et pertes sur prêts du Grenier (14 016 392 FCFA) non provisionnés.",
     "**Les provisions ne sont pas insuffisantes en montant global, mais mal affectées et non calculées.** Elles doivent être recalculées chaque année (statuts, art. 89). L'excédent ne peut être réaffecté qu'après décision des organes et avis de l'AIRMS.",
     "🟠", "6, 117 à 129, 144"],
    ["ALM", "Demande 3 : adossement actif-passif",
     "Actif à moins d'un an au 31/08/2026 : 3 457 866 951 FCFA. Paiements attendus dans l'année, épargne comprise : 10 046 113 184 FCFA. Couverture : 34,42 % (26,22 % en test sévère). Manque cumulé à cinq ans : 12 596 052 822 FCFA.",
     "**L'épargne, exigible à tout moment, finance des actifs longs.** Terrains, clinique, créance sur l'hôpital et prêts du Grenier ne se vendent pas vite. Il faut un échéancier de l'épargne, une réserve de liquidité et une politique actif-passif votée.",
     "🔴", "7, 149 à 158"],
    ["PLA", "Demande 4 : adéquation des placements aux engagements",
     "Placements financiers moyens 2025 : 763 752 135 FCFA, soit 5,53 % de l'épargne due. Rendement moyen de 2021 à 2025 : 4,90 %. Produits financiers 2025 (117 912 488 FCFA) inférieurs aux intérêts d'une année sur l'EFC à 3,50 % (309 709 577 FCFA). Aucune politique de placement votée.",
     "**Le problème n'est pas d'abord le rendement, mais le volume et la liquidité.** L'essentiel de l'épargne est immobilisé dans des actifs qui ne produisent ni revenu financier ni liquidité.",
     "🔴", "8, 162 à 166"],
]
ENS_COMPL = [
    ["REG-1", "Conformité prudentielle (36 indicateurs de l'AIRMS)",
     "14 indicateurs sur 36 non conformes de 2022 à 2025, 7 conformes, 15 en attente de pièces. Fonds propres : -879 498 700 FCFA au 31/08/2026. Liquidité : 0,13 en 2025 (1,15 en lecture corrigée).",
     "**Le FPPN expose son agrément.** Deux exercices de suite non conformes permettent un retrait d'agrément (règlement n° 07/2009, art. 84).",
     "🔴", "19, 168 à 217"],
    ["STAT-2", "Sinistralité et résultat par garantie",
     "Sinistralité maladie : 6,00 % en 2021, 73,53 % en 2025, 97,40 % sur les huit mois de 2026. Ratio combiné maladie 2025, frais compris : 109,87 %.",
     "**La garantie maladie perd de l'argent, et le décès la finance.** La cotisation maladie ne couvre plus le coût complet de la garantie.",
     "🟠", "72 à 86, 175"],
    ["REG-2", "Gouvernance des décisions",
     "Placements décidés par la direction générale, sans délibération du conseil. Projet de Cocody sans budget. Ni politique de placement, ni comité actif-passif.",
     "**L'épargne des adhérents est employée sans décision des organes**, alors que les statuts l'exigent (art. 85).",
     "🟠", "157, 163, 164"],
    ["DON-1", "Données de base",
     "4 287 fichiers lus, 4 454 416 lignes de cotisations contrôlées. Ni base des prestations datée, ni dates de naissance. Fichiers maladie et décès de décembre 2024 vides.",
     "**Des données essentielles manquent pour finir les calculs.** Mortalité observée, santé des retraités et projection par âge restent impossibles.",
     "🟠", "45 à 58"],
    ["STAT-1", "Effectifs et cotisations",
     "23 119 cotisants en janvier 2024, 31 198 en août 2026 (+34,95 %). Taux prélevés : 3,77 % et 1,23 % au lieu de 3,50 % et 1,50 %. Cotisations : +45,76 % de 2020 à 2025.",
     "**Les effectifs progressent, mais la répartition des cotisations diffère des statuts.**",
     "🟡", "59 à 71"],
    ["STAT-3", "Tables de mortalité et projection de l'équilibre",
     "4 tables de mortalité comparées. Projection agrégée de l'équilibre de 2026 à 2030, en deux scénarios.",
     "**Les calculs par âge restent à faire.** La projection sera affinée à réception des dates de naissance.",
     "🟡", "218 à 228"],
]
LEG_VERD = [("🔴", "Insuffisant", "exigence non satisfaite ou risque élevé"), ("🟠", "Partiel", "démontré en partie, impact significatif"), ("🟡", "Moyen", "amélioration souhaitable"), ("🟢", "Satisfaisant", "point maîtrisé")]
DETAIL = {"ENG-1": "5, 95 à 98", "ENG-2": "110 à 112, 124 à 127", "PROV-1": "6, 117 à 127", "PROV-2": "131 à 142", "PROV-3": "129", "ALM-1": "7, 147 à 158", "PLA-1": "8, 159 à 166",
          "REG-1": "19, 168 à 217", "STAT-2": "72 à 86", "REG-2": "157, 163, 164", "DON-1": "45 à 58", "STAT-1": "59 à 71", "STAT-3": "218 à 228"}
SRC_R = "Source : rapport actuariel et statistique du 23/10/2026"

# ---------------------------------------------------------------- diapositives
SL.append({"type": "couverture", "notes": "Résumé exécutif du rapport actuariel et statistique : contexte et objectifs, constats et enseignements, recommandations."})
SL.append({"type": "sommaire", "notes": ""})
gab(31, {"titre": "Contexte de la mission", "sous_titre": "Le FPPN gère la prévoyance et l'épargne des policiers, l'analyse actuarielle dit s'il peut tenir ses promesses",
          "synthese": "Volet actuariel et statistique de l'audit comptable et financier du 01/01/2020 au 31/08/2026"},
    {"cartes": [("Le FPPN", "Mutuelle des policiers de Côte d'Ivoire : soins, allocation au décès, épargne EFC et PEP, prêts du Grenier. 31 198 policiers cotisants en août 2026."),
                ("Son cadre", "Mutualité sociale de l'UEMOA, sous le contrôle de l'AIRMS, avec 36 indicateurs de solidité à respecter (Décision n° 05/2022)."),
                ("La mission", "Audit comptable et financier du 01/01/2020 au 31/08/2026, sur la base d'une offre technique validée par le FPPN. Lancement les 13 et 14/10/2026."),
                ("Le rapport actuariel", "291 diapositives, établies sur les documents reçus jusqu'au 22/10/2026. Chaque chiffre est relié au classeur des calculs.")]},
    "Présenter le FPPN, son cadre, la mission et le rapport actuariel que ce résumé synthétise.",
    source="Source : offre technique Adenka ; statuts du FPPN édités le 15/06/2023 ; " + SRC_R[9:] + ", diapositives 41 à 44 et 61")
SL[-1]["sommaire"] = "Contexte et objectifs"
contenu({"titre": "Ce qui est attendu des analyses", "sous_titre": "Quatre questions posées par la mission, et les travaux nécessaires pour y répondre",
         "synthese": "Chaque analyse attendue est reprise dans la partie suivante, avec ses constats, ses enseignements et le renvoi au rapport principal"},
        "Lire chaque ligne : la question posée, ce que l'analyse doit établir, puis où trouver le détail dans le rapport principal.",
        tableau={"titre": None, "entetes": ["Code", "Analyse attendue", "Question posée", "Ce que l'analyse doit établir", "Détail dans le rapport"], "largeurs": [1.9, 6.2, 9.0, 9.6, 4.5], "taille": 9.5,
                 "sans_code": True, "lignes": [
                     ["ENG", "**Demande 1 :** engagements envers les adhérents et ayants droit", "Que doit le FPPN à ses adhérents : épargne et intérêts, allocations décès, soins, personnel ?", "La valeur de chaque engagement à la date d'arrêté, et ce qui manque pour la démontrer", "Diapositives 5, 87 à 113, 143"],
                     ["PROV", "**Demande 2 :** suffisance des provisions au regard des prestations servies", "Les provisions du bilan couvrent-elles ce qui reste à payer ?", "La comparaison, provision par provision, du montant au bilan et du besoin mesuré", "Diapositives 6, 114 à 146"],
                     ["ALM", "**Demande 3 :** adossement actif-passif", "Les actifs permettent-ils de payer les engagements à leur échéance ?", "La couverture des paiements de l'année et à cinq ans, avec des tests de résistance", "Diapositives 7, 147 à 158"],
                     ["PLA", "**Demande 4 :** adéquation des placements aux engagements", "Les placements sont-ils assez sûrs, liquides et rentables pour servir l'épargne ?", "Le volume, le rendement et la gouvernance des placements", "Diapositives 8, 159 à 166"],
                     ["DON  STAT  REG", "**Travaux complémentaires**, nécessaires aux quatre demandes", "Les données sont-elles fiables, la garantie équilibrée, les règles de l'AIRMS respectées ?", "La qualité des données, les effectifs, la sinistralité, les 36 indicateurs, les tables de mortalité et la projection", "Diapositives 45 à 86, 168 à 228"]]},
        source=SRC_R + ", diapositives 4 à 8 et 41 à 44 ; offre technique Adenka validée par le FPPN")
for k, (lot, titre_, sous_, synth_) in enumerate((
        (ENS_DEMANDES, "Constats et enseignements : les quatre demandes", "Engagements et provisions partiellement démontrés, adossement et placements insuffisants",
         "Pour chaque analyse attendue : ce que l'on a vu, ce que l'on en retient, le verdict et l'endroit où voir le détail"),
        (ENS_COMPL, "Constats et enseignements : travaux complémentaires", "Les travaux complémentaires confirment le risque sur l'agrément et sur la garantie maladie",
         "Du plus critique au moins critique : ce que l'on a vu, ce que l'on en retient et l'endroit où voir le détail"))):
    contenu({"titre": titre_, "sous_titre": sous_, "synthese": synth_},
            "Lire chaque ligne de gauche à droite : l'analyse attendue, les constats chiffrés, l'enseignement annoncé en gras, le verdict, puis les diapositives du rapport principal.",
            tableau={"titre": None, "entetes": ["Code", "Analyse attendue", "Ce que l'on a vu", "Ce que l'on en retient", "Détail (diapositives)", "Verdict"], "largeurs": [1.75, 4.4, 10.6, 9.4, 3.3, 1.75],
                     "taille": 8.5 if k == 0 else 8, "sans_code": True, "legende": LEG_VERD,
                     "lignes": [[r_[0], r_[1], r_[2], r_[3], r_[5], r_[4]] for r_ in lot]},
            sommaire="Constats et enseignements" if k == 0 else None, source=SRC_R)
TOUS = sorted(RECAP_DEMANDES + RECAP_COMPL, key=lambda r_: "🔴🟠🟡🟢".index(r_[0][0]))
contenu({"titre": "Constats classés par criticité", "sous_titre": "Quatre constats critiques : l'épargne, la liquidité, les placements et la conformité prudentielle",
         "synthese": "Les treize constats du plus critique au moins critique, avec la conclusion soulignée, le commentaire chiffré et le renvoi au rapport principal"},
        "Lire la conclusion soulignée de chaque constat, puis le commentaire chiffré. La criticité figure en dernière colonne.",
        tableau={"titre": None, "entetes": ["Code", "Analyse", "Constat", "Commentaires", "Détail", "Criticité"], "largeurs": [1.6, 3.7, 7.6, 13.9, 2.65, 1.75], "taille": 7.4,
                 "sans_code": True, "legende": LEG_CRIT,
                 "lignes": [[r_[0].split(" ", 1)[1], r_[1], "**" + r_[4].split("**")[1] + "**", r_[4].split("**", 2)[2].strip(), DETAIL[r_[0].split(" ", 1)[1]], r_[0].split(" ", 1)[0]] for r_ in TOUS]},
        source=SRC_R + ", diapositives 4 à 13, 45 à 228 ; plan de travail du 23/10/2026")
#@@CONSTATS_ACTIVITES@@
contenu({"titre": "Recommandations - Actuariat", "sous_titre": "Cinq axes de recommandations, dont huit actions à lancer dans les trois mois",
         "synthese": "Chaque axe répond aux constats codés du récapitulatif, avec une priorité de P1 à P4"},
        "Présenter axe par axe. Le détail des actions figure aux diapositives 25 à 27 du rapport principal.",
        dessin="recos_rapport", sommaire="Recommandations", source=SRC_R + ", diapositives 25 à 27 ; résumé exécutif du 26/10/2026")
SL.append({"type": "merci", "notes": ""})
ENTREES = [(x["sommaire"], i) for i, x in enumerate(SL, 1) if x.get("sommaire")]
'''

CONS = open(os.path.join(ICI, "deck_constats_activites_fppn.py"), encoding="utf-8").read()
_donnees = CONS[CONS.index("CRIT = {"):CONS.index("\n\n\ndef pire")] + "\n" + CONS[CONS.index("SOUS = {"):CONS.index("def h_bloc")]
CONTENU = CONTENU.replace("#@@HELPERS_ACTIVITES@@", _donnees + "\n" + ACT_AIDES).replace("#@@CONSTATS_ACTIVITES@@", ACT_DIAPOS)
MOTEUR = MOTEUR.replace('        if "tableau" in x:\n            tableau_avec_analyse(sl, x)', '        if x.get("nav"):\n            d_nav(sl, x["nav"])\n            tableau(sl, x["tableau"], 4.85, hauteur_max=BAS_CONTENU - 4.85)\n        elif "tableau" in x:\n            tableau_avec_analyse(sl, x)', 1)
assert 'd_nav(sl, x["nav"])' in MOTEUR
LIENS = r"""
# ---- navigation : chaque bouton est une forme unique qui porte l'action « aller à la diapositive »
prs = Presentation(SORTIE)
cible = {}
for sl in prs.slides:
    for sh in sl.shapes:
        if sh.name.startswith("Ici "):
            cible[sh.name[4:]] = sl
        if sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text.strip() == "Constats par activité":
            cible["MENU"] = sl
n_liens = 0
for sl in prs.slides:
    for sh in sl.shapes:
        if sh.name.startswith("Nav ") and sh.name[4:] in cible:
            sh.click_action.target_slide = cible[sh.name[4:]]
            n_liens += 1
prs.save(SORTIE)
finaliser(SORTIE)
print("liens de navigation :", n_liens, "| fiches d'activité :", sum(1 for k in cible if k != "MENU"))
"""
code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\n" + LIENS + "\nprint(SORTIE, len(SL), 'diapositives')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "deck_resume_executif_rapport_fppn (assemblé)", "exec"), g)

# conclusions soulignées dans la colonne « Constat »
from pptx import Presentation as _P
_prs = _P(SORTIE)
for _sl in _prs.slides:
    if any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text.strip() == "Constats classés par criticité" for sh in _sl.shapes):
        for _sh in _sl.shapes:
            if _sh.has_table:
                for _row in list(_sh.table.rows)[1:]:
                    for _p in _row.cells[2].text_frame.paragraphs:
                        for _r in _p.runs:
                            _r.font.underline = True
# remplissage coloré de la colonne « Statut » : vert si réalisé, orange si en cours
from pptx.dml.color import RGBColor as _RGB
for _sl in _prs.slides:
    if any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text.strip() == "Actions effectuées" for sh in _sl.shapes):
        for _sh in _sl.shapes:
            if _sh.has_table:
                for _row in list(_sh.table.rows)[1:]:
                    _c = _row.cells[4]
                    _t = _c.text_frame.text
                    _col = "C6E0B4" if _t.startswith("Réalisé") else "F8CBAD" if _t.startswith("En cours") else None
                    if _col:
                        _c.fill.solid()
                        _c.fill.fore_color.rgb = _RGB.from_string(_col)
_prs.save(SORTIE)

# version diaporama (.ppsx) : s'ouvre directement en mode diaporama, où les boutons répondent au simple clic
import zipfile as _zf
_ppsx = os.path.splitext(SORTIE)[0] + ".ppsx"
with _zf.ZipFile(SORTIE) as _zi, _zf.ZipFile(_ppsx, "w", _zf.ZIP_DEFLATED) as _zo:
    for _it in _zi.infolist():
        _d = _zi.read(_it.filename)
        if _it.filename == "[Content_Types].xml":
            _d = _d.replace(b"presentationml.presentation.main+xml", b"presentationml.slideshow.main+xml")
        _zo.writestr(_it, _d)
print(_ppsx)
