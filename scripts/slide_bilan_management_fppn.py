"""Diapositive unique pour une réunion de management : bilan du volet actuariel et statistique de la mission FPPN.

Gabarit Timesaver n° 25 (9 blocs icône + titre orange + texte) sous l'en-tête Adenka : le poids des travaux, sept
résultats (constat, cause, conséquence pour le FPPN, en langage simple) et les décisions à prendre.
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
avant = "def remplir_gab(slide, num, d, analyse=None):\n"
assert avant in MOTEUR
MOTEUR = MOTEUR.replace(avant, avant + "    if num == 25:\n        remplir_25(slide, d)\n        return\n")

CONTENU = r'''
VALIDITE = {}
PBCM = {}
GAB = {"contenu": 5}


def remplir_25(slide, d):
    """Neuf blocs : titre orange (paragraphe du gabarit) puis paragraphes « Constat / Cause / Pour le FPPN »."""
    boites = sorted([sh for sh in slide.shapes if sh.name == "TextBox 49"], key=lambda sh: (round(sh.top / 360000), sh.left))
    icones = sorted([sh for sh in slide.shapes if sh.shape_type == 13 and sh.width < Cm(3)], key=lambda sh: (round(sh.top / 360000), sh.left))
    # grille régulière sous l'en-tête : 3 colonnes (icône + texte) et 3 rangées
    xs_ic, xs_tx, larg = (1.3, 12.0, 22.7), (3.1, 13.8, 24.5), (8.6, 8.6, 8.2)
    ys, h_r, t_ic = (3.85, 8.0, 12.15), 4.1, 1.5
    for k, ic in enumerate(icones[:9]):
        r_, c_ = divmod(k, 3)
        ic.width, ic.height = Cm(t_ic), Cm(t_ic)
        ic.left, ic.top = Cm(xs_ic[c_]), Cm(ys[r_] + (h_r - t_ic) / 2)
    for k, (sh, (titre, parties)) in enumerate(zip(boites, d["blocs"])):
        r_, c_ = divmod(k, 3)
        sh.left, sh.top, sh.width, sh.height = Cm(xs_tx[c_]), Cm(ys[r_]), Cm(larg[c_]), Cm(h_r)
        sh.text_frame.margin_top = Cm(0.02)
        tf = sh.text_frame
        tf.word_wrap = True
        ecrire(sh, [titre], [10], [True])
        for lib, texte in parties:
            p = tf.add_paragraph()
            p.space_before = Pt(1.5)
            if lib:
                r_ = p.add_run()
                r_.text = lib + " : "
                r_.font.name, r_.font.size, r_.font.bold = "Helvetica", Pt(8), True
                r_.font.color.rgb = RGBColor.from_string(BLEU)
            for morceau, g_, s_ in morceaux_balises(texte):
                r_ = p.add_run()
                r_.text = morceau
                r_.font.name, r_.font.size, r_.font.bold = "Helvetica", Pt(8), g_
                r_.font.color.rgb = RGBColor.from_string(NOIR)
                if s_:
                    surligner(r_)


BLOCS = [
    ("Le poids des travaux", [
        ("", "**4 454 416 lignes** de cotisations contrôlées sur **80 mois** (2020 au 31/08/2026), **4 287 fichiers** du FPPN extraits, balances et grands livres relus."),
        ("", "**99 onglets** de calcul, **36 ratios** réglementaires recalculés, **2 885 chiffres** vérifiés sans écart, 88 documents demandés, 4 entretiens, 64 constats."),
        ("Avancement", "63,85 % au 23/10/2026.")]),
    ("Liquidité : insuffisante", [
        ("Constat", "l'argent disponible dans l'année couvre 34,42 % de ce que le FPPN doit payer dans l'année. Il manquerait 12,6 milliards de FCFA en cinq ans."),
        ("Cause", "l'épargne des policiers a financé la clinique, des terrains et des prêts, qui ne se revendent pas vite."),
        ("Pour le FPPN", "il paie les sorties d'épargne avec les nouvelles cotisations. Si les départs à la retraite augmentent, il ne pourra pas payer tout le monde à temps.")]),
    ("Épargne : dette sous-évaluée", [
        ("Constat", "15,7 milliards de FCFA d'épargne due au 31/08/2026, 19,0 milliards avec les intérêts promis."),
        ("Cause", "aucun taux n'est fixé officiellement (3,50 %, 5,50 % ou 14 % selon le document) et aucun intérêt n'est comptabilisé depuis 2020, soit 3,3 milliards."),
        ("Pour le FPPN", "ses comptes sous-estiment ce qu'il doit aux policiers, et un épargnant peut contester le montant reçu à son départ.")]),
    ("Provisions : un montant hérité", [
        ("Constat", "4,8 milliards de FCFA mis de côté pour la maladie et le décès, le même montant depuis 2020, pour un besoin mesuré de 19,6 millions."),
        ("Cause", "[[un montant passé vers 2004 selon le chef comptable]], jamais recalculé. Les allocations décès dues sont déjà comptées ailleurs (489 millions)."),
        ("Pour le FPPN", "ce n'est pas un trou : l'argent existe mais il est mal rangé. Il gonfle les dettes et cache l'absence de réserve pour les intérêts de l'épargne.")]),
    ("Placements : trop faibles", [
        ("Constat", "5,53 % de l'épargne réellement placée en 2025, pour un rendement de 4,90 % par an."),
        ("Cause", "presque toute l'épargne est immobilisée dans des biens qui ne rapportent ni intérêts ni liquidités, sans politique de placement votée."),
        ("Pour le FPPN", "les placements ne rapportent pas de quoi payer les intérêts promis : 470 millions de FCFA d'écart en 2025.")]),
    ("Fonds propres : négatifs", [
        ("Constat", "-879 millions de FCFA au 31/08/2026, après -6,7 milliards fin 2020."),
        ("Cause", "des pertes passées que les excédents de 2020 à 2025 n'ont pas encore effacées, puis un déficit de 441 millions sur les huit premiers mois de 2026."),
        ("Pour le FPPN", "aucun coussin pour absorber un coup dur. Comptes remis en ordre (provisions, intérêts, prêts), ils seraient positifs d'environ 637 millions.")]),
    ("Règles de solidité : non respectées", [
        ("Constat", "14 indicateurs réglementaires sur 36 non conformes de 2022 à 2025 : réserve, liquidité, fonds propres, frais."),
        ("Cause", "les points précédents, et des frais de fonctionnement de 36,34 % des cotisations pour un maximum de 20 %."),
        ("Pour le FPPN", "deux années de suite non conformes permettent à l'AIRMS de retirer l'agrément. Ce n'est pas automatique, mais c'est l'enjeu principal.")]),
    ("Garantie maladie : déficitaire", [
        ("Constat", "les soins consomment 73,53 % des cotisations maladie en 2025 et 97,40 % en 2026. Déficit de 702 millions sur huit mois de 2026."),
        ("Cause", "les soins augmentent beaucoup plus vite que les cotisations. Le niveau bas de 2020 à 2022 venait d'une couverture limitée et de factures impayées."),
        ("Pour le FPPN", "la garantie décès paie le déficit. Sans mesure, il faudra revoir la cotisation ou les remboursements.")]),
    ("Décisions à prendre dans les trois mois", [
        ("1", "fixer le taux de l'épargne et comptabiliser les intérêts dus."),
        ("2", "calculer les provisions chaque année et reclasser l'excédent."),
        ("3", "constituer une réserve de liquidité et cesser de financer des biens durables avec l'épargne."),
        ("4", "présenter un plan de mise en conformité à l'AIRMS."),
        ("5", "transmettre les données manquantes : dates de naissance, prestations datées, comptes individuels.")]),
]

gab(25, {"titre": "Bilan du volet actuariel et statistique",
         "sous_titre": "Le FPPN est excédentaire, mais ne peut pas démontrer qu'il remboursera l'épargne à temps",
         "synthese": "Plus de quatre millions de lignes contrôlées sur 80 mois, sept résultats clés et cinq décisions à prendre dans les trois mois"},
    {"blocs": BLOCS},
    "Lire de gauche à droite et de haut en bas : le poids des travaux, sept résultats avec leur cause et leur conséquence pour le FPPN, puis les décisions. Message clé : le risque principal est la liquidité de l'épargne, l'enjeu réglementaire est l'agrément.",
    source="Source : rapport actuariel et statistique du 23/10/2026, diapositives 4 à 13, 98, 118, 150, 161, 165, 178, 195 et 239 ; plan de travail du 23/10/2026 ; liste des documents demandés du 26/10/2026")
'''

code = ENTETE + "\n" + CONTENU + "\n" + MOTEUR + "\nprint(SORTIE, len(SL), 'diapositive(s)')\n"
g = {"__name__": "__main__", "__file__": os.path.join(ICI, "rapport_v12_fppn.py"), "SKILL": SKILL, "SORTIE": SORTIE}
exec(compile(code, "slide_bilan_management_fppn (assemblé)", "exec"), g)
