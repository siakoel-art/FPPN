"""Diapositive unique pour une réunion de management : bilan du volet actuariel et statistique de la mission FPPN.

Trois colonnes à la charte Adenka : le poids des travaux (volumétrie), sept résultats (chiffre clé, cause, conséquence
pour le FPPN, en langage simple), les décisions à prendre ; une bande sur les méthodes et l'avancement. Icônes Adinkra
de la planche du Timesaver dans les en-têtes de colonne.
Chiffres du rapport actuariel et statistique du 23/10/2026, du plan de travail du 23/10/2026 et de la liste des
documents demandés du 26/10/2026. Moteur de mise en page repris du script du rapport (rapport_v12_fppn.py).

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
    tf.margin_top = tf.margin_bottom = Cm(0.04)
    tf.vertical_anchor = ancre
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    return sh


def entete_col(slide, x, y, l, num, titre, k_icone):
    b = cadre(slide, x, y, l, 0.66, BLEU, ancre=MSO_ANCHOR.MIDDLE, marge=0.8)
    b.text_frame.word_wrap = False
    p = b.text_frame.paragraphs[0]
    run(p, f"{num}  ", 11, True, ORANGE)
    run(p, titre, 10, True, "FFFFFF")
    ic = icones[k_icone % len(icones)]
    slide.shapes.add_picture(io.BytesIO(icone_blanche(ic.image.blob)), Cm(x + 0.15), Cm(y + 0.08), Cm(0.5), Cm(0.5))


RESULTATS = [
    ("C00000", "34,42 %", "Liquidité : insuffisante", "des paiements de l'année couverts. Il manquerait 12,6 milliards de FCFA en cinq ans.",
     "l'épargne a financé clinique, terrains et prêts, qui ne se revendent pas vite.",
     "il paie les sorties d'épargne avec les nouvelles cotisations : une vague de départs le mettrait en difficulté."),
    ("C00000", "14 sur 36", "Règles de solidité : non respectées", "indicateurs réglementaires non conformes de 2022 à 2025.",
     "épargne mal couverte, fonds propres négatifs, frais à 36,34 % des cotisations (maximum 20 %).",
     "deux années non conformes de suite permettent à l'AIRMS de retirer l'agrément."),
    ("C00000", "5,53 %", "Placements : trop faibles", "de l'épargne réellement placée en 2025, rendement de 4,90 % par an.",
     "l'épargne est immobilisée dans des biens qui ne rapportent rien, sans politique de placement.",
     "les placements ne paient pas les intérêts promis : 470 millions de FCFA d'écart en 2025."),
    ("ED7D31", "19,0 Md", "Épargne : dette sous-évaluée", "de FCFA dus aux épargnants, dont 3,3 milliards d'intérêts jamais comptabilisés.",
     "aucun taux fixé officiellement (3,50 %, 5,50 % ou 14 % selon le document).",
     "la dette est sous-évaluée, et un épargnant peut contester ce qu'il reçoit."),
    ("ED7D31", "4,8 Md", "Provisions : un montant hérité", "de FCFA figés depuis 2020, pour un besoin mesuré de 19,6 millions.",
     "[[montant passé vers 2004 selon le chef comptable]], jamais recalculé, décès dus déjà comptés ailleurs.",
     "pas un trou : l'argent existe mais il est mal rangé, et rien n'est prévu pour les intérêts."),
    ("ED7D31", "-879 M", "Fonds propres : négatifs", "de FCFA au 31/08/2026, après -6,7 milliards fin 2020.",
     "pertes passées pas encore effacées, puis déficit de 441 millions sur huit mois de 2026.",
     "aucun coussin pour un coup dur. Comptes remis en ordre : environ +637 millions."),
    ("ED7D31", "97,40 %", "Garantie maladie : déficitaire", "des cotisations maladie consommées par les soins en 2026 (73,53 % en 2025).",
     "les soins augmentent beaucoup plus vite que les cotisations.",
     "le décès paie le déficit (702 millions en 2026) : cotisation ou remboursements à revoir."),
]


def dessin_bilan(slide):
    y0 = 3.95
    y_haut, y_bas = y0 + 0.74, 15.55
    # ---- colonne 1 : le poids des travaux
    x1, l1 = 1.4, 6.0
    entete_col(slide, x1, y0, l1, "1", "Le poids des travaux", 0)
    tuiles = [("4 454 416", "lignes de cotisations contrôlées"), ("80 mois", "couverts, 2020 au 31/08/2026"),
              ("4 287", "fichiers du FPPN extraits"), ("2 885", "chiffres vérifiés, 0 écart"),
              ("99", "onglets de calcul"), ("36", "ratios réglementaires recalculés"),
              ("88", "documents demandés, 4 entretiens"), ("64", "constats, 10 risques notés")]
    g = 0.12
    lt = (l1 - g) / 2
    ht = (y_bas - y_haut - 3 * g) / 4
    for k, (val, lib) in enumerate(tuiles):
        x = x1 + (k % 2) * (lt + g)
        y = y_haut + (k // 2) * (ht + g)
        b = cadre(slide, x, y, lt, ht, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE, marge=0.08)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, val, 15, True, BLEU)
        p2 = b.text_frame.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run(p2, lib, 7.5, False, "53565A")
    # ---- colonne 2 : les premiers résultats
    x2, l2 = x1 + l1 + 0.3, 17.8
    entete_col(slide, x2, y0, l2, "2", "Les premiers résultats", 1)
    n = len(RESULTATS)
    hr = (y_bas - y_haut - (n - 1) * 0.06) / n
    for k, (coul, chiffre, titre, constat, cause, effet) in enumerate(RESULTATS):
        y = y_haut + k * (hr + 0.06)
        cadre(slide, x2, y, 0.16, hr, coul)
        b = cadre(slide, x2 + 0.16, y, 2.45, hr, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE, marge=0.05)
        p = b.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run(p, chiffre, 13, True, coul)
        b = cadre(slide, x2 + 2.66, y, l2 - 2.66, hr, None, "D9D9D9", ancre=MSO_ANCHOR.MIDDLE, marge=0.12)
        tf = b.text_frame
        p = tf.paragraphs[0]
        run(p, titre + " : ", 8.5, True)
        run(p, constat, 8)
        p2 = tf.add_paragraph()
        run(p2, "Pourquoi : ", 8, True, "53565A")
        balise(p2, cause, 8, "53565A")
        p3 = tf.add_paragraph()
        run(p3, "→ Pour le FPPN : ", 8, True, BLEU)
        run(p3, effet, 8, False, BLEU)
    # ---- colonne 3 : les décisions à prendre
    x3 = x2 + l2 + 0.3
    l3 = 31.2 + 1.4 - x3
    entete_col(slide, x3, y0, l3, "3", "Les décisions à prendre", 2)
    b = cadre(slide, x3, y_haut, l3, 2.3, "FBE5C8", ancre=MSO_ANCHOR.MIDDLE)
    p = b.text_frame.paragraphs[0]
    run(p, "Message clé : ", 9, True, "C00000")
    run(p, "le FPPN n'est pas en faillite, il dégage des excédents. Son risque principal est la liquidité de l'épargne, et son enjeu réglementaire est l'agrément.", 9)
    dec = [("Fixer le taux de l'épargne", "et comptabiliser les intérêts dus"), ("Calculer les provisions", "chaque année et reclasser l'excédent"),
           ("Constituer une réserve de liquidité", "et cesser de financer des biens durables avec l'épargne"),
           ("Présenter un plan de conformité", "chiffré à l'AIRMS"), ("Transmettre les données manquantes", "pour finir les calculs")]
    b = cadre(slide, x3, y_haut + 2.4, l3, y_bas - y_haut - 2.4, None, "D9D9D9", ancre=MSO_ANCHOR.MIDDLE)
    tf = b.text_frame
    p = tf.paragraphs[0]
    run(p, "Cinq actions à lancer dans les trois mois", 9.5, True, BLEU)
    p.space_after = Pt(5)
    for k, (t_, s_) in enumerate(dec, 1):
        q = tf.add_paragraph()
        q.space_after = Pt(5)
        run(q, f"{k}  ", 10.5, True, ORANGE)
        run(q, t_ + " ", 9, True)
        run(q, s_, 8.5)
    # ---- bande du bas : méthodes et avancement
    b = cadre(slide, x1, y_bas + 0.08, 31.2, 0.72, "F2F2F2", ancre=MSO_ANCHOR.MIDDLE)
    tf = b.text_frame
    p = tf.paragraphs[0]
    run(p, "Méthodes : ", 7.5, True, BLEU)
    run(p, "écart de liquidité par échéance, écoulement de l'épargne, tests de résistance, provisions par règlements postérieurs à la clôture et cadence des déclarations tardives, intérêts sur encours moyen, 36 ratios de la Décision n° 05/2022.", 7.5)
    p2 = tf.add_paragraph()
    run(p2, "Avancement : ", 7.5, True, BLEU)
    run(p2, "63,85 % au 23/10/2026. Conclusions acquises sur la liquidité et les placements. Montants des engagements et des provisions à préciser avec les données du FPPN.", 7.5)


contenu({"titre": "Volet actuariel et statistique : bilan des travaux",
         "sous_titre": "Le FPPN est excédentaire, mais ne peut pas démontrer qu'il remboursera l'épargne à temps",
         "synthese": "Plus de quatre millions de lignes contrôlées sur 80 mois, sept résultats clés et cinq décisions à prendre dans les trois mois"},
        "Lire de gauche à droite : ce qui a été fait, ce qui en ressort et pourquoi, ce qu'il faut décider. Message clé : le risque principal est la liquidité de l'épargne, l'enjeu réglementaire est l'agrément.",
        bilan=True,
        source="Source : rapport actuariel et statistique du 23/10/2026, diapositives 4 à 13, 98, 118, 150, 161, 165, 178, 195 et 239 ; plan de travail du 23/10/2026 ; liste des documents demandés du 26/10/2026")
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositive(s)')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "slide_bilan_management_fppn (assemblé)", "exec"), g)
