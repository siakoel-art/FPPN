# -*- coding: utf-8 -*-
"""
FPPN - Note dirigeant : ce qui change dans le rapport actuariel par rapport à la version du 09/11/2026 (éléments en bleu).
Deux pages au plus. Les numéros de diapositives sont lus dans le rapport (titres), les chiffres viennent des mêmes sources
que le rapport (assets/sinistralite_v15.json, constats de contenu_actuariel_fppn.py).
Utilisation : python3 note_mises_a_jour_fppn.py <rapport.pptx> <sortie.docx>
"""
import json
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation

ICI = os.path.dirname(os.path.abspath(__file__))
RAPPORT, SORTIE = sys.argv[1], sys.argv[2]
SQ = {int(k): v for k, v in json.load(open(os.path.join(ICI, "assets", "sinistralite_v15.json"), encoding="utf-8")).items()}


def nb(x, d=0):
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def pc(x):
    return nb(100 * x, 2) + " %"


def F(x):
    return nb(round(x)) + " FCFA"


# numéros de diapositives par titre
NUM = {}
for i, sl in enumerate(Presentation(RAPPORT).slides, 1):
    for sh in sl.shapes:
        if sh.has_text_frame and sh.name == "Adenka_Titre":
            NUM.setdefault(sh.text_frame.text.strip(), i)


def dia(*titres):
    n = sorted(NUM[t] for t in titres if t in NUM)
    if not n:
        return "n.d."
    return ", ".join(str(x) for x in n) if len(n) < 3 or n[-1] - n[0] != len(n) - 1 else f"{n[0]} à {n[-1]}"


s25, s26 = SQ[2025], SQ[2026]
c25 = s25["cats"]
HON25 = c25["Honoraires d'un cabinet (factures trimestrielles)"]
MESSAGES = [
    ("La hausse de la sinistralité n'est pas une bonne nouvelle en soi.",
     f" Elle corrige un sous-service ancien, mais une partie n'est pas du soin : sans les dépenses étrangères aux soins, le taux réglementaire 2025 passe de {pc(s25['sin'])} à {pc(s25['sin_soins'])}, sous la norme de 70 %. "
     f"La maladie reste déficitaire : ratio combiné des soins seuls de {pc(s26['comb_soins'])} au 31/08/2026."),
    ("Plus d'un quart des « prestations maladie » de 2025 ne sont pas des soins.",
     f" {F(s25['non_soins'])} ({pc(s25['non_soins'] / s25['prest'])}) : dons {F(c25['Dons (autorités, élus, délégués)'])}, gadgets et imprimés {F(c25['Gadgets, imprimés et cartes'])}, "
     f"honoraires trimestriels d'un cabinet {F(HON25)}."),
    ("Le taux proche de zéro de 2020 à 2022 ne mesurait pas le coût réel des soins.",
     " Des dépenses de santé étaient inscrites à l'actif : clinique du FPPN, créance sur l'Hôpital de la Police (3 407 519 850 FCFA), créance sur l'État (3 932 124 730 FCFA). En 2022, les charges nettes de la clinique (469 980 456 FCFA) auraient porté le taux maladie de 6,47 % à 32,86 %, ce que confirme le mémo du service comptable."),
    ("La valeur de l'actif est surestimée.",
     " 3 782 519 850 FCFA sont à sortir de l'actif selon le commissaire aux comptes (terrain d'Attoban, créance sur l'Hôpital de la Police). La réserve de réévaluation du siège (2 899 400 000 FCFA) égale sa valeur entière estimée en 2019, et non une plus-value."),
    ("Des engagements hors bilan pèsent sur la trésorerie.",
     " 263 000 000 FCFA bloqués à la BOA après un procès perdu, et un projet immobilier à Azaguié de 19 539 377 111 FCFA approuvé sans plan de financement."),
    ("Un tiers des salariés n'est pas dans la provision retraite.",
     " En août 2026, 33 salariés sur 100, surtout de la clinique, sont déclarés sans numéro CNPS et absents du calcul de la provision."),
]
JAUNE_MSG = {}

CHANGEMENTS = [
    ("Sinistralité (nouveau, 11/11)", "Analyse dans les deux sens : contenu réel du compte, taux des soins seuls, explication du taux bas, synthèse",
     "Piloter le ratio combiné par garantie, sortir les dépenses hors soins, justifier les honoraires",
     dia("Sinistralité : lire le taux dans les deux sens", "Prestations maladie : ce que contient le compte", "Sinistralité des soins seuls",
         "Pourquoi le taux était proche de zéro de 2020 à 2022", "Sinistralité : ce qu'il faut en retenir")),
    ("Actifs (ACT-1, nouveau)", "Deux terrains sans titre foncier, deux créances sans accord de remboursement, sensibilité des fonds propres",
     "Fonds propres et couverture de l'épargne plus faibles qu'affiché", dia("Actifs dont la valeur n'est pas établie")),
    ("Hors bilan (ALM-2, nouveau)", "Procès perdu (fonds bloqués), projet d'Azaguié, prise de participation",
     "Étude de faisabilité et plan de financement avant tout engagement", dia("Engagements et projets hors bilan")),
    ("Grenier du Policier (GRE-1, nouveau)", "États financiers propres 2024 et 2025 : 8 493 055 501 FCFA d'épargne PEP et EFC à son passif, aucun prêt déprécié",
     "Fixer le propriétaire de l'épargne et supprimer la double inscription", dia("Grenier du Policier : ses propres états financiers")),
    ("Provision du personnel", "Le compte 196 est le calcul du FPPN en droits acquis, sans actualisation ni mortalité",
     "Recalcul actuariel maintenu (écart de 1,08 % sur les droits acquis 2025)", dia("Provision du personnel : méthode du FPPN et compte 196")),
    ("Ratios prudentiels", "Calculs du FPPN comparés au recalcul : mêmes sinistralité et frais, réserve, liquidité et fonds propres divergents",
     "Réconcilier les ratios avant transmission à l'AIRMS", dia("Ratios prudentiels : calculs du FPPN et d'Adenka")),
    ("Documents et demandes", "980 fichiers nouveaux ou modifiés contrôlés, réponses partielles et six nouvelles questions au FPPN",
     "Liste des demandes et PBC mis à jour", dia("Documents reçus le 09/11/2026")),
    ("Siège et comptes 47 et 27 (nouveau, 11/11)", "Réserve de réévaluation égale à la valeur entière du siège (2 899 400 000 FCFA). Mémo comptable : 11 406 323 425 FCFA de créances sur la clinique, l'hôpital et le Grenier",
     "Produire l'écriture d'origine, justifier ou déprécier les créances", dia("Réévaluation du siège : le compte 10620000", "Comptes 47 et 27 : ce qu'en dit le service comptable")),
    ("Grenier : liaison (nouveau, 11/11)", "Écart de 70 191 500 FCFA entre les comptes de liaison, épargne inscrite au Grenier sans actif réel",
     "Rapprocher chaque mois, supprimer l'épargne inscrite au Grenier", dia("Grenier du Policier : liaison avec le FPPN")),
    ("Personnel (nouveau, 11/11)", "33 salariés sur 100 sans numéro CNPS en août 2026, hors provision retraite",
     "Déclarer ces salariés et les intégrer dans la provision", dia("Personnel : salariés hors provision retraite")),
    ("Placements (11/11)", "Le FCP de 500 000 000 FCFA ouvert le 04/11/2025 provient du dépôt à terme de l'épargne PEP",
     "Vérifier sa valeur liquidative et son règlement", dia("Calcul pas à pas : couverture de l'épargne")),
    ("Gouvernance (mention brève)", "Prorogation des mandats jugée nulle par l'AIRMS, mise en demeure de transmettre les comptes 2025",
     "Relève du volet juridique, sans effet sur les chiffres actuariels", "—"),
]
DECISIONS = [
    "Reclasser hors des prestations les dons, gadgets et imprimés, et produire les pièces des honoraires trimestriels du cabinet.",
    "Facturer la clinique du FPPN au coût réel, avec une comptabilité séparée, et recalculer la cotisation maladie (3,5 % du salaire de base).",
    "Décider du sort du terrain d'Attoban et de la créance sur l'Hôpital de la Police, et déprécier ce qui n'est pas recouvrable.",
    "Suspendre tout engagement sur le projet d'Azaguié jusqu'à une étude de faisabilité et un plan de financement.",
    "Produire l'écriture d'origine de la réévaluation du siège et corriger la réserve, puis rapprocher le compte de liaison avec le Grenier.",
    "Déclarer à la CNPS les 33 salariés sans numéro et les intégrer dans la provision retraite.",
]

# ------------------------------------------------------------------ mise en page
d = Document()
sec = d.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(1.6)
st = d.styles["Normal"]
st.font.name, st.font.size = "Helvetica", Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Helvetica")
st.paragraph_format.space_after = Pt(4)
NOIR, GRIS = RGBColor(0x06, 0x06, 0x06), RGBColor(0x59, 0x59, 0x59)


def para(t, taille=10, gras=False, couleur=NOIR, apres=4):
    p = d.add_paragraph()
    r = p.add_run(t)
    r.font.size, r.bold, r.font.color.rgb = Pt(taille), gras, couleur
    p.paragraph_format.space_after = Pt(apres)
    return p


def fond(cell, hexa):
    tc = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexa)
    tc.append(sh)


para("FPPN : audit du volet actuariel", 9, couleur=GRIS, apres=0)
para("Ce qui change dans le rapport du 11/11/2026", 16, True, apres=2)
para("Note pour la direction. Elle résume les ajouts et corrections écrits en bleu dans le rapport, par rapport à la version du 09/11/2026 (documents reçus les 09/11 et 11/11/2026). "
     "Les éléments surlignés en jaune restent à confirmer par le FPPN.", 9, couleur=GRIS, apres=8)

para("Messages clés", 12, True, apres=3)
for k, (g, t) in enumerate(MESSAGES, 1):
    p = d.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(0.6), Cm(-0.6)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"{k}. {g}")
    r.bold = True
    p.add_run(t)
    if k in JAUNE_MSG:
        r = p.add_run(JAUNE_MSG[k])
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW

para("Ce qui a changé dans le rapport", 12, True, apres=3)
tb = d.add_table(rows=1, cols=4)
tb.alignment = WD_TABLE_ALIGNMENT.CENTER
tb.style = "Table Grid"
tb.autofit = False
larg = [Cm(3.4), Cm(6.6), Cm(5.4), Cm(2.4)]
for j, h in enumerate(["Sujet", "Ce qui change", "Ce que cela implique", "Diapositives"]):
    c = tb.rows[0].cells[j]
    c.text = ""
    r = c.paragraphs[0].add_run(h)
    r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), RGBColor(0xFF, 0xFF, 0xFF)
    fond(c, "060606")
for i, ligne in enumerate(CHANGEMENTS):
    cells = tb.add_row().cells
    for j, v in enumerate(ligne):
        cells[j].text = ""
        r = cells[j].paragraphs[0].add_run(v)
        r.font.size, r.bold = Pt(8.5), j == 0
        if j == 0:
            fond(cells[j], "D9D9D9")
        elif i % 2:
            fond(cells[j], "F2F2F2")
for row in tb.rows:
    for j, c in enumerate(row.cells):
        c.width = larg[j]
para("Source : rapport actuariel du 11/11/2026 ; documents de la dataroom reçus les 09/11 et 11/11/2026 ; grands livres des comptes 2021 au 31/08/2026.", 7.5, couleur=GRIS, apres=8)

para("Décisions attendues du FPPN", 12, True, apres=3).paragraph_format.keep_with_next = True
for k, t in enumerate(DECISIONS, 1):
    p = d.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(0.6), Cm(-0.6)
    p.paragraph_format.space_after = Pt(3)
    p.add_run(f"{k}. {t}")
para("Sigles : AIRMS : Agence ivoirienne de régulation de la mutualité sociale. BOA : Bank of Africa. CMP : Clinique médicale du policier. CNPS : Caisse nationale de prévoyance sociale. "
     "EFC : épargne de fin de carrière. FCP : fonds commun de placement. FPPN : Fonds de prévoyance de la Police nationale. PBC : liste des documents demandés. PEP : plan d'épargne du policier.",
     7.5, couleur=GRIS, apres=0)
d.save(SORTIE)
print("note écrite :", SORTIE)
