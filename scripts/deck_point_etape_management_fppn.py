"""Document de point d'étape pour le management : volet actuariel et statistique de la mission FPPN (fond noir).

Structure demandée : contexte et objectifs, récapitulatif des activités (tableau codifié par axe : méthodes, poids des
travaux, résultats), recommandations par axe (une diapositive), prochaines étapes (une diapositive).
Chiffres du rapport actuariel et statistique du 10/11/2026, du plan de travail du 23/10/2026 et de la liste des
documents demandés du 26/10/2026. Moteur de mise en page repris du script du rapport (rapport_v12_fppn.py).

Usage : python3 deck_point_etape_management_fppn.py <skill adenka-powerpoint> <sortie.pptx>
"""
import os
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
for avant, apres in [('"Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "10 novembre 2026"', '"Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "Novembre 2026"'),
                     ('r.text = "Rapport actuariel et statistique"', 'r.text = "Point d\'étape : volets Actuariat et Stratégie"'),
                     ("def remplir_gab(slide, num, d, analyse=None):\n", "def remplir_gab(slide, num, d, analyse=None):\n    if num == 11:\n        remplir_11(slide, d)\n        return\n"),
                     ('entrees = [("Résumé exécutif", 3)]', 'entrees = ENTREES'),
                     ('        sommaire(sl, entrees)', '        sommaire_noir(sl, entrees)'),
                     ('        if x.get("triangle"):', '        if x.get("dessin"):\n            DESSINS[x["dessin"]](sl)\n        if x.get("triangle"):')]:
    assert avant in MOTEUR, avant
    MOTEUR = MOTEUR.replace(avant, apres)

CONTENU = r'''
import copy
VALIDITE = {}
PBCM = {}
GAB = {"couverture": 1, "sommaire": 4, "contenu": 5, "merci": 110}
BLANC, GRIS_CLAIR, CARTE = "FFFFFF", "BBBCBC", "F2F2F2"
PRIO = {"P1": "C00000", "P2": "ED7D31", "P3": "FFC000", "P4": "70AD47"}
RAP = "Source : rapport actuariel et statistique du 10/11/2026"


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
    ["🔴 ACT-1", "Valeur des actifs", "Lecture des rapports du commissaire aux comptes, des PV du conseil et des courriers de l'AIRMS reçus le 09/11/2026, sensibilité des fonds propres",
     "4 actifs examinés (2 terrains, 2 créances), rapports du commissaire aux comptes de 2022 à 2025",
     "**Certains actifs ne sont pas correctement évalués.** Terrain d'Attoban sans titre (375 000 000 FCFA), créance sur l'Hôpital de la Police inchangée depuis 2020 (3 407 519 850 FCFA). Sans eux, les fonds propres passeraient à -4 662 018 550 FCFA au 31/08/2026."],
    ["🟠 ALM-2", "Engagements hors bilan", "Relevé des décisions des PV du conseil et de l'Assemblée générale de 2021 à 2026",
     "28 PV et rapports du conseil et de l'Assemblée générale",
     "**Des engagements hors bilan pèsent sur la trésorerie future.** Procès perdu sur l'EFC : 263 000 000 FCFA bloqués. Projet d'Azaguié : 19 539 377 111 FCFA, sans plan de financement. Projet d'entrée au capital de SOMAVIE."],
    ["🟡 GRE-1", "Comptes du Grenier", "Lecture des états financiers du Grenier 2022 à 2025, rapprochement avec les balances âgées",
     "États financiers de 4 exercices",
     "**Le Grenier ne déprécie aucun prêt et inscrit l'épargne PEP et EFC dans ses comptes.** Résultat 2025 : 30 543 376 FCFA avant toute perte. Épargne de 8 493 055 501 FCFA inscrite dans deux comptabilités."],
    ["🟡 STAT-3", "Mortalité et projection", "Tables CIMA, projection agrégée de l'équilibre en deux scénarios",
     "4 tables de mortalité, projection de 2026 à 2030 en 2 scénarios",
     "**Les calculs par âge restent à faire.** La mortalité par âge est impossible sans dates de naissance. La projection de l'équilibre n'est que globale : elle sera affinée à réception des données."],
]
LEG_CRIT = [("🔴", "Critique", "risque élevé ou non-conformité"), ("🟠", "Élevé", "impact significatif"), ("🟡", "Moyen", "amélioration souhaitable"), ("🟢", "Faible", "point maîtrisé ou de faible montant")]

# ---------------------------------------------------------------- diapositives
SL.append({"type": "couverture", "notes": "Point d'étape du volet actuariel et statistique pour le management : contexte, récapitulatif des travaux, recommandations, prochaines étapes."})
SL.append({"type": "sommaire", "notes": ""})
gab(31, {"titre": "Contexte de la mission", "sous_titre": "Le FPPN gère la prévoyance et l'épargne des policiers, l'audit dit s'il peut tenir ses promesses",
          "synthese": "Audit comptable et financier du 01/01/2020 au 31/08/2026, présenté ici pour ses volets Actuariat et Stratégie"},
    {"cartes": [("Le FPPN", "Mutuelle des policiers de Côte d'Ivoire : soins, allocation au décès, épargne EFC et PEP, prêts du Grenier. 31 198 policiers cotisants en août 2026."),
                ("Son cadre", "Mutualité sociale de l'UEMOA, sous le contrôle de l'AIRMS, avec 36 indicateurs de solidité à respecter (Décision n° 05/2022)."),
                ("La mission Adenka", "Audit comptable et financier du 01/01/2020 au 31/08/2026, sur la base d'une offre technique validée par le FPPN. Lancement les 13 et 14/10/2026."),
                ("Deux volets présentés", "Actuariat : ce que le FPPN devra payer et sa capacité à le payer. Stratégie : objectifs présentés par l'équipe Stratégie.")]},
    "Présenter le FPPN, son cadre, la mission Adenka et les deux volets traités dans ce document.",
    source="Source : offre technique Adenka ; statuts du FPPN édités le 15/06/2023 ; " + RAP[9:] + ", diapositive 68")
SL[-1]["sommaire"] = "Contexte et objectifs"
gab(11, {"titre": "Objectifs par volet", "sous_titre": "Deux volets complémentaires : mesurer les engagements, éclairer la stratégie",
         "synthese": "Les codes du volet Actuariat sont repris dans le récapitulatif des activités et les recommandations"},
    {"cartes": [("Volet Actuariat", [("Les quatre questions de la mission", ["ENG : engagements envers les adhérents", "PROV : suffisance des provisions", "ALM : adossement actif-passif", "PLA : adéquation des placements"]),
                                     ("Travaux complémentaires", ["DON : contrôle des données de base", "STAT : effectifs, cotisations, sinistralité", "REG : indicateurs de solidité et gouvernance"])]),
                ("Volet Stratégie", [("Objectifs", ["À renseigner par l'équipe Stratégie"]), ("Travaux", ["À renseigner par l'équipe Stratégie"])])]},
    "Carte noire : volet Actuariat. Carte orange : volet Stratégie, complétée par l'équipe en charge.",
    source="Source : offre technique Adenka validée par le FPPN, pôle 4 (p. 70) et profil de l'actuaire (p. 74) ; plan d'analyse statistique du 16/10/2026")
ORDRE_ACT = ["ENG-1", "ENG-2", "PROV-1", "PROV-2", "PROV-3", "ALM-1", "ALM-2", "PLA-1", "ACT-1", "DON-1", "STAT-1", "STAT-2", "STAT-3", "GRE-1", "REG-1", "REG-2"]
TOUT = {r_[0].split(" ", 1)[1]: r_ for r_ in RECAP_DEMANDES + RECAP_COMPL}
SRC_RECAP = RAP + ", diapositives 4 à 13, 48 à 53, 56 à 68, 81, 95 à 173, 185 et 203 ; plan de travail du 23/10/2026 ; liste des documents demandés du 26/10/2026"
STR_VIDE = [["STR-1", "Stratégie : à renseigner", "", ""], ["STR-2", "Stratégie : à renseigner", "", ""]]
STATUT = {"ENG-1": "**En cours.** Reste : comptes individuels d'épargne",
          "ENG-2": "**En cours.** Reste : dossiers non payés, dates de naissance des retraités",
          "PROV-1": "**En cours.** Reste : base des prestations datée",
          "PROV-2": "**Réalisé.** Fichiers de calcul du FPPN reçus le 09/11/2026",
          "PROV-3": "**Réalisé**",
          "ALM-1": "**En cours.** Reste : inventaire et échéancier signés des placements",
          "PLA-1": "**Réalisé**",
          "DON-1": "**En cours.** Reste : fichiers de décembre 2024, base des prestations",
          "STAT-1": "**Réalisé**",
          "STAT-2": "**Réalisé**",
          "STAT-3": "**En cours.** Reste : dates de naissance, projection par garantie",
          "REG-1": "**En cours.** Reste : ratios en attente de pièces",
          "REG-2": "**Réalisé**",
          "ACT-1": "**Réalisé.** Décision du CA attendue", "ALM-2": "**Réalisé.** Jugement et plan de financement attendus", "GRE-1": "**Réalisé**"}
for _k14, _ord14 in enumerate((ORDRE_ACT[:8], ORDRE_ACT[8:]), 1):
    contenu({"titre": f"Actions effectuées ({_k14}/2)", "sous_titre": "Seize activités menées sur le volet Actuariat, des quatre demandes aux travaux complémentaires",
             "synthese": "Pour chaque activité : la méthode, le volume traité et le statut. Lignes du volet Stratégie renseignées par l'équipe en charge"},
            "Lire chaque ligne de gauche à droite : le code, l'activité, la méthode retenue, le volume traité, puis le statut et ce qui reste à obtenir. Les quatre demandes d'abord, les travaux complémentaires ensuite.",
            tableau={"titre": None, "entetes": ["Code", "Activité", "Méthodes", "Volume traité", "Statut au 10/11/2026"], "largeurs": [1.75, 4.6, 11.2, 7.9, 5.75], "taille": 7.8,
                     "sans_code": True, "lignes": [[c_, TOUT[c_][1], TOUT[c_][2], TOUT[c_][3], STATUT[c_]] for c_ in _ord14] + ([["STR", "Stratégie : à renseigner par l'équipe en charge", "", "", ""]] if _k14 == 2 else [])},
            sommaire="Récapitulatif des activités" if _k14 == 1 else None, source=SRC_RECAP)
contenu({"titre": "Constats", "sous_titre": "Les engagements et les provisions sont partiellement démontrés, l'adossement et les placements insuffisants",
         "synthese": "Les quatre demandes d'abord, puis les travaux complémentaires, chacun du constat le plus critique au moins critique"},
        "Lire la conclusion soulignée de chaque constat, puis le commentaire chiffré. La criticité figure en dernière colonne.",
        tableau={"titre": None, "entetes": ["Code", "Activité", "Constat", "Commentaires", "Criticité"], "largeurs": [1.75, 4.4, 8.3, 15.0, 1.75], "taille": 6.9,
                 "sans_code": True,
                 "lignes": [[r_[0].split(" ", 1)[1], r_[1], "**" + r_[4].split("**")[1] + "**", r_[4].split("**", 2)[2].strip(), r_[0].split(" ", 1)[0]]
                            for lot in (RECAP_DEMANDES, RECAP_COMPL) for r_ in sorted(lot, key=lambda r_: "🔴🟠🟡🟢".index(r_[0][0]))]},
        source=SRC_RECAP)
contenu({"titre": "Recommandations", "sous_titre": "Des axes de recommandations par volet, avec une priorité de P1 à P4",
         "synthese": "Volet Actuariat : cinq axes reliés aux constats codés. Volet Stratégie : axes renseignés par l'équipe en charge"},
        "Présenter axe par axe, sans entrer dans le détail des actions, qui figure dans le rapport et le résumé exécutif.",
        dessin="recos", sommaire="Recommandations", source=RAP + ", diapositives 25 à 27 ; résumé exécutif du 26/10/2026")
contenu({"titre": "Prochaines étapes", "sous_titre": "Le rapport définitif est prévu le 30/11/2026, sous réserve des données du FPPN",
         "synthese": "Étapes du volet Actuariat jusqu'au rapport définitif, étapes du volet Stratégie renseignées par l'équipe en charge"},
        "Insister sur la dépendance aux données du FPPN et sur ce qui est déjà acquis.",
        dessin="etapes", sommaire="Prochaines étapes", source="Source : plan de travail du 23/10/2026")
SL.append({"type": "merci", "notes": ""})
ENTREES = [(x["sommaire"], i) for i, x in enumerate(SL, 1) if x.get("sommaire")]
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositives')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "deck_point_etape_management_fppn (assemblé)", "exec"), g)

# conclusions soulignées dans la colonne « Constat »
from pptx import Presentation as _P
_prs = _P(SORTIE)
for _sl in _prs.slides:
    if any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text.strip() == "Constats" for sh in _sl.shapes):
        for _sh in _sl.shapes:
            if _sh.has_table:
                for _row in list(_sh.table.rows)[1:]:
                    for _p in _row.cells[2].text_frame.paragraphs:
                        for _r in _p.runs:
                            _r.font.underline = True
# remplissage coloré de la colonne « Statut » : vert si réalisé, orange si en cours
from pptx.dml.color import RGBColor as _RGB
for _sl in _prs.slides:
    if any(sh.has_text_frame and sh.name == "Adenka_Titre" and sh.text_frame.text.strip() .startswith("Actions effectuées") for sh in _sl.shapes):
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
