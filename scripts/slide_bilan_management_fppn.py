"""Diapositive unique pour une réunion de management : bilan du volet actuariel et statistique de la mission FPPN.

Trois colonnes : le poids des travaux (volumétrie), les premiers résultats commentés, les décisions à prendre ; une
bande sur les méthodes, l'avancement et la suite. Chiffres du rapport actuariel et statistique du 23/10/2026, du plan
de travail du 23/10/2026 et de la liste des documents demandés du 26/10/2026.
Moteur de mise en page repris du script du rapport (rapport_v12_fppn.py).

Usage : python3 slide_bilan_management_fppn.py <skill adenka-powerpoint> <sortie.pptx>
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
avant = '        if x.get("triangle"):'
assert avant in MOTEUR
MOTEUR = MOTEUR.replace(avant, '        if x.get("bilan"):\n            dessin_bilan(sl)\n' + avant)

CONTENU = r'''
VALIDITE = {}
PBCM = {}
GAB = {"contenu": 5}


def run(p, texte, taille, gras=False, couleur=NOIR):
    r_ = p.add_run()
    r_.text = texte
    r_.font.name = "Helvetica"
    r_.font.size = Pt(taille)
    r_.font.bold = gras
    r_.font.color.rgb = RGBColor.from_string(couleur)
    return r_


def cadre(slide, x, y, l, h, fond=None, bord=None, ancre=MSO_ANCHOR.TOP, marge=0.15):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(l), Cm(h))
    if fond:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fond)
    else:
        sh.fill.background()
    if bord:
        sh.line.color.rgb = RGBColor.from_string(bord)
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Cm(marge)
    tf.margin_top = tf.margin_bottom = Cm(0.06)
    tf.vertical_anchor = ancre
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    return sh


def entete_col(slide, x, y, l, num, titre):
    b = cadre(slide, x, y, l, 0.62, BLEU, ancre=MSO_ANCHOR.MIDDLE)
    p = b.text_frame.paragraphs[0]
    run(p, f"{num}  ", 12, True, ORANGE)
    run(p, titre, 11.5, True, "FFFFFF")


def dessin_bilan(slide):
    y0 = 3.95
    y_bas = y0 + 0.75 + 4 * 1.95 + 3 * 0.12
    # ---- colonne 1 : le poids des travaux
    x1, l1 = 1.4, 9.2
    entete_col(slide, x1, y0, l1, "1", "Le poids des travaux")
    tuiles = [("4 454 416", "lignes de cotisations contrôlées"), ("80 mois", "couverts, de 2020 au 31/08/2026"),
              ("4 287", "fichiers du FPPN extraits et inventoriés"), ("2 885", "chiffres du rapport vérifiés, 0 écart"),
              ("99", "onglets de calcul à formules visibles"), ("36", "ratios réglementaires recalculés"),
              ("88", "documents demandés, 4 entretiens"), ("64", "constats, 10 risques notés")]
    lt, ht, g = (l1 - 0.2) / 2, 1.95, 0.2
    for k, (val, lib) in enumerate(tuiles):
        x = x1 + (k % 2) * (lt + g)
        y = y0 + 0.75 + (k // 2) * (ht + 0.12)
        b = cadre(slide, x, y, lt, ht, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, val, 20, True, BLEU)
        p2 = b.text_frame.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run(p2, lib, 9, False, "53565A")
    # ---- colonne 2 : les premiers résultats
    x2, l2 = x1 + l1 + 0.4, 11.6
    entete_col(slide, x2, y0, l2, "2", "Les premiers résultats")
    res = [("C00000", "Adossement : insuffisant", "34,42 %", "des paiements de l'année couverts par l'argent disponible. Manque de 12,6 milliards de FCFA à cinq ans.", "Risque principal : la liquidité de l'épargne."),
           ("C00000", "Placements : insuffisants", "5,53 %", "de l'épargne réellement placée. Rendement de 4,90 %, sous le taux de 5,50 % promis sur l'EFC.", "L'épargne a financé clinique, terrains et prêts."),
           ("ED7D31", "Engagements : partiellement démontrés", "19,0 Md", "de FCFA dus aux épargnants, dont 3,3 milliards d'intérêts jamais comptabilisés.", "La dette réelle est sous-évaluée."),
           ("ED7D31", "Provisions : partiellement démontrées", "4,8 Md", "de FCFA figés depuis 2020, très au-dessus du besoin mesuré (19,6 millions).", "L'argent existe, mais il est mal affecté.")]
    hr = (y_bas - (y0 + 0.75) - 3 * 0.12) / 4
    for k, (coul, titre, chiffre, txt, com) in enumerate(res):
        y = y0 + 0.75 + k * (hr + 0.12)
        cadre(slide, x2, y, 0.18, hr, coul)
        b = cadre(slide, x2 + 0.18, y, 2.9, hr, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE, marge=0.08)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, chiffre, 17, True, coul)
        b = cadre(slide, x2 + 3.15, y, l2 - 3.15, hr, None, "D9D9D9", ancre=MSO_ANCHOR.MIDDLE)
        p = b.text_frame.paragraphs[0]
        run(p, titre, 10.5, True)
        p2 = b.text_frame.add_paragraph()
        run(p2, txt, 9.5)
        p3 = b.text_frame.add_paragraph()
        run(p3, "→ " + com, 9.5, True, BLEU)
    # ---- colonne 3 : ce que cela implique
    x3, l3 = x2 + l2 + 0.4, 31.2 - (l1 + l2 + 0.8)
    entete_col(slide, x3, y0, l3, "3", "Les décisions à prendre")
    b = cadre(slide, x3, y0 + 0.75, l3, 2.15, "FBE5C8", ancre=MSO_ANCHOR.MIDDLE)
    p = b.text_frame.paragraphs[0]
    run(p, "Enjeu majeur : ", 10, True, "C00000")
    run(p, "14 règles de solidité sur 36 non respectées et fonds propres négatifs (-879 millions de FCFA). Deux exercices non conformes de suite permettent à l'AIRMS de retirer l'agrément.", 9.5)
    dec = [("Fixer le taux servi", "et comptabiliser les intérêts dus"), ("Calculer les provisions", "chaque année et reclasser l'excédent"),
           ("Constituer une réserve de liquidité", "et cesser de financer des biens durables avec l'épargne"),
           ("Présenter un plan de conformité", "chiffré à l'AIRMS"), ("Obtenir les données manquantes", "pour finir les calculs")]
    b = cadre(slide, x3, y0 + 3.02, l3, y_bas - (y0 + 3.02), None, "D9D9D9", ancre=MSO_ANCHOR.MIDDLE)
    tf = b.text_frame
    p = tf.paragraphs[0]
    run(p, "Cinq actions à lancer dans les trois mois", 10.5, True, BLEU)
    p.space_after = Pt(6)
    for k, (t_, s_) in enumerate(dec, 1):
        q = tf.add_paragraph()
        q.space_after = Pt(6)
        run(q, f"{k}  ", 11, True, ORANGE)
        run(q, t_ + " ", 10.5, True)
        run(q, s_, 10)
    # ---- bande du bas : méthodes, avancement, suite
    yb = y_bas + 0.2
    b = cadre(slide, x1, yb, 31.2, 1.55, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE)
    tf = b.text_frame
    p = tf.paragraphs[0]
    run(p, "Méthodes : ", 9, True, BLEU)
    run(p, "écart de liquidité par échéance, écoulement de l'épargne, tests de résistance, provisions par règlements postérieurs à la clôture et cadence des déclarations tardives, intérêts sur encours moyen, rendement pondéré par les jours, 36 ratios de la Décision n° 05/2022.", 9)
    p2 = tf.add_paragraph()
    run(p2, "Avancement : ", 9, True, BLEU)
    run(p2, "63,85 % au 23/10/2026. Les conclusions sur la liquidité et les placements sont acquises. Les montants des engagements et des provisions seront précisés à réception des données du FPPN (base des prestations datée, dates de naissance, comptes individuels d'épargne).", 9)


contenu({"titre": "Volet actuariel et statistique : bilan des travaux",
         "sous_titre": "Le FPPN est excédentaire, mais ne peut pas démontrer qu'il remboursera l'épargne à temps",
         "synthese": "Plus de quatre millions de lignes contrôlées sur 80 mois, quatre réponses à la mission, cinq décisions à prendre dans les trois mois"},
        "Lire de gauche à droite : ce qui a été fait, ce qui en ressort, ce qu'il faut décider. Message clé : le risque principal est la liquidité de l'épargne, l'enjeu réglementaire est l'agrément.",
        bilan=True,
        source="Source : rapport actuariel et statistique du 23/10/2026, diapositives 4 à 8, 13 et 239 ; plan de travail du 23/10/2026 ; liste des documents demandés du 26/10/2026")
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositive(s)')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "slide_bilan_management_fppn (assemblé)", "exec"), g)
