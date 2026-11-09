# -*- coding: utf-8 -*-
"""Remplit les parties du volet actuariel dans le résumé exécutif général de la mission (document Word fourni).

Parties remplies, sans toucher au reste du document (styles, numérotation et mise en forme d'origine conservés) :
  2.1.2 Approche actuarielle ; 2.2.2 Limites actuarielles ;
  3.3.1 Analyse actuarielle (sous l'audit comptable et financier) ;
  3.4.1 Ratios prudentiels : travaux du volet actuariel ;
  4.4.1 Ratios prudentiels et régulateurs et 4.4.2 Analyse actuarielle : plans d'actions priorisés.

Usage : python3 remplir_resume_general_fppn.py <résumé général.docx> <sortie.docx>
"""
import copy
import os
import random
import re
import sys

import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contenu_actuariel_fppn as CA  # noqa: E402

SRC, SORTIE = sys.argv[1:3]
d = docx.Document(SRC)
COUL = {"🔴": "C00000", "🟠": "ED7D31", "🟡": "FFC000", "🟢": "70AD47"}
TAILLE = Pt(11)


def num_id(p):
    n = p._p.find(qn("w:pPr"))
    n = n.find(qn("w:numPr")) if n is not None else None
    return n.find(qn("w:numId")).get(qn("w:val")) if n is not None and n.find(qn("w:numId")) is not None else None


def trouver(texte, nid, rang=0):
    res = [p for p in d.paragraphs if p.text.strip() == texte and num_id(p) == nid]
    return res[rang]


# modèles de mise en forme pris dans le document
P_NORMAL = next(p for p in d.paragraphs if p.text.startswith("Le Fonds de Prévoyance de la Police Nationale (FPPN) est une mutuelle"))
P_PUCE = next(p for p in d.paragraphs if p.text.startswith("Prise de connaissance"))
P_TITRE4 = trouver("Environnement de contrôle", "13")


# ---------------------------------------------------------------- numérotation des nouveaux titres (3.3.x et 3.4.x)
numbering = d.part.numbering_part.element


def nouvelle_liste(texte_niveau, modele_num="13"):
    abs_map = {n.get(qn("w:numId")): n.find(qn("w:abstractNumId")).get(qn("w:val")) for n in numbering.findall(qn("w:num"))}
    absn = next(a for a in numbering.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == abs_map[modele_num])
    nouv = copy.deepcopy(absn)
    nid_abs = str(max(int(a.get(qn("w:abstractNumId"))) for a in numbering.findall(qn("w:abstractNum"))) + 1)
    nouv.set(qn("w:abstractNumId"), nid_abs)
    ns = nouv.find(qn("w:nsid"))
    if ns is not None:
        ns.set(qn("w:val"), "%08X" % random.randint(0, 0xFFFFFFFF))
    nouv.findall(qn("w:lvl"))[0].find(qn("w:lvlText")).set(qn("w:val"), texte_niveau)
    numbering.findall(qn("w:num"))[0].addprevious(nouv)
    nid = str(max(int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))) + 1)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), nid)
    an = OxmlElement("w:abstractNumId")
    an.set(qn("w:val"), nid_abs)
    num.append(an)
    numbering.append(num)
    return nid


NUM_33 = nouvelle_liste("3.3.%1.")
NUM_34 = nouvelle_liste("3.4.%1.")


# ---------------------------------------------------------------- écriture
class Curseur:
    def __init__(self, apres):
        self.el = apres._p if hasattr(apres, "_p") else apres

    def poser(self, el):
        self.el.addnext(el)
        self.el = el
        return el


def runs(p, texte, taille=TAILLE, gras=None, italique=None, couleur=None):
    """Ajoute le texte avec **gras** et [[surligné jaune]]."""
    for morceau in re.split(r"(\[\[.+?\]\])", texte):
        if not morceau:
            continue
        surl = morceau.startswith("[[")
        morceau = morceau[2:-2] if surl else morceau
        for m2 in re.split(r"(\*\*.+?\*\*)", morceau):
            if not m2:
                continue
            g = m2.startswith("**") and m2.endswith("**")
            r = p.add_run(m2[2:-2] if g else m2)
            r.font.name = "Helvetica"
            r.font.size = taille
            r.bold = True if (g or gras) else None
            r.italic = italique
            if couleur:
                r.font.color.rgb = RGBColor.from_string(couleur)
            if surl:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW


def paragraphe(cur, modele, texte, num=None, **kw):
    p_el = OxmlElement("w:p")
    ppr = modele._p.find(qn("w:pPr"))
    if ppr is not None:
        p_el.append(copy.deepcopy(ppr))
    cur.poser(p_el)
    p = docx.text.paragraph.Paragraph(p_el, modele._parent)
    if num:
        p_el.find(qn("w:pPr")).find(qn("w:numPr")).find(qn("w:numId")).set(qn("w:val"), num)
    runs(p, texte, **kw)
    return p


def normal(cur, texte, **kw):
    return paragraphe(cur, P_NORMAL, texte, **kw)


def puce(cur, texte):
    return paragraphe(cur, P_PUCE, texte)


def titre(cur, texte, num):
    p = paragraphe(cur, P_TITRE4, texte, num=num)
    return p


def fond(cellule, couleur):
    tcpr = cellule._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), couleur)
    tcpr.append(shd)


def tableau(cur, entetes, lignes, largeurs, taille=8.5, boules=()):
    t = d.add_table(rows=len(lignes) + 1, cols=len(entetes))
    tblpr = t._tbl.tblPr
    bords = OxmlElement("w:tblBorders")
    for cote in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement("w:" + cote)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "BBBCBC")
        bords.append(b)
    tblpr.append(bords)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblpr.append(lay)
    for i, ligne in enumerate([entetes] + lignes):
        for j, val in enumerate(ligne):
            c = t.rows[i].cells[j]
            c.width = Cm(largeurs[j])
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if i == 0:
                fond(c, "060606")
                runs(p, val, Pt(taille), gras=True, couleur="FFFFFF")
            elif j in boules and val:
                b, _, lib = val.partition(" ")
                runs(p, "● ", Pt(taille + 2), gras=True, couleur=COUL.get(b, "000000"))
                runs(p, lib, Pt(taille), gras=True)
            else:
                if i % 2 == 0:
                    fond(c, "F2F2F2")
                runs(p, val, Pt(taille), gras=(j == 0) or None)
    # ligne d'en-tête répétée sur chaque page
    trpr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trpr.append(th)
    for k, w in enumerate(largeurs):
        t.columns[k].width = Cm(w)
    cur.poser(t._tbl)
    return t


def legende(cur, texte):
    return normal(cur, texte, taille=Pt(9), italique=True)


def source(cur, texte):
    return normal(cur, "Source : " + texte, taille=Pt(8), italique=True, couleur="53565A")


# ---------------------------------------------------------------- 2.1.2 approche actuarielle
cur = Curseur(trouver("Approche actuarielle", "9"))
normal(cur, CA.APPROCHE_INTRO)
normal(cur, "Les travaux se sont déroulés en six étapes :")
for e in CA.APPROCHE_ETAPES:
    puce(cur, e)
normal(cur, "**Les méthodes ont été choisies selon les données disponibles.** Chaque méthode retenue est la plus précise que permettent les documents transmis :")
legende(cur, "Tableau : méthodes actuarielles retenues et raison du choix")
tableau(cur, ["Analyse", "Méthode retenue", "Pourquoi ce choix", "Données utilisées"], [list(m) for m in CA.METHODES], [3.0, 5.4, 4.4, 3.2], taille=8)
source(cur, CA.RAP + ", diapositives 47, 48 et 238 à 245.")

# ---------------------------------------------------------------- 2.2.2 limites actuarielles
cur = Curseur(trouver("Limites actuarielles", "10"))
normal(cur, CA.LIMITES_INTRO)
for e in CA.LIMITES:
    puce(cur, e)
normal(cur, CA.LIMITES_EFFET)

# ---------------------------------------------------------------- 3.3.1 analyse actuarielle (sous l'audit comptable et financier)
cur = Curseur(trouver("Audit Comptable et Financier", "11"))
titre(cur, "Analyse actuarielle", NUM_33)
normal(cur, "**Le FPPN connaît ses dettes de court terme, mais pas la valeur de ses engagements d'épargne.** Faute de taux arrêté, d'échéancier et d'intérêts comptabilisés, deux des quatre questions de la mission ne sont que partiellement démontrées et les deux autres reçoivent une réponse insuffisante.")
legende(cur, "Tableau : réponse aux quatre questions du volet actuariel")
tableau(cur, ["Question", "Verdict", "En une phrase", "Détail"], [[q, f"{b} {v}", t, dt] for q, b, v, t, dt in CA.VERDICTS], [3.6, 2.6, 7.8, 2.0], taille=8.5, boules=(1,))
source(cur, CA.RAP + ", diapositives 4 à 8.")
for titre_q, msg, paras, renvoi in CA.RESULTATS:
    normal(cur, f"**{titre_q} : {msg[0].lower() + msg[1:]}**")
    for x in paras:
        normal(cur, x)
    normal(cur, renvoi + ".", taille=Pt(9), italique=True, couleur="53565A")
normal(cur, "**Les autres analyses confirment ces fragilités :**")
for e in CA.AUTRES_CONSTATS:
    puce(cur, e)

# ---------------------------------------------------------------- 3.4.1 ratios prudentiels : travaux du volet actuariel
cur = Curseur(trouver("La vérification du respect des ratios prudentiels et réglementaires applicables", "11"))
titre(cur, "Ratios prudentiels : travaux du volet actuariel", NUM_34)
normal(cur, CA.RATIOS_INTRO)
for x in CA.RATIOS_PARAS:
    normal(cur, x)
legende(cur, "Tableau : principaux indicateurs de solidité en 2025 (Décision n° 05/2022)")
tableau(cur, ["Indicateur", "Norme", "Valeur 2025", "Statut"],
        [[a, b, c, f"{s} {'Conforme' if s == '🟢' else 'Non conforme'}"] for a, b, c, s in CA.RATIOS_TABLE], [5.0, 4.2, 4.0, 2.8], taille=8.5, boules=(3,))
source(cur, CA.RAP + ", " + CA.RATIOS_RENVOI.split(", ", 1)[1] + ".")

# ---------------------------------------------------------------- 4.4.1 et 4.4.2 plans d'actions priorisés
CONST = {c[1]: c for c in CA.CONSTATS}
ENT = ["N°", "Constat", "Risque", "Recommandation", "Priorité", "Responsable", "Échéance"]
LARG = [0.7, 3.0, 3.3, 3.4, 1.4, 2.5, 1.7]


def ligne_plan(k, action):
    axe, act, codes, prio, qui = action
    c = CONST[codes.split(",")[0].strip()]
    return [str(k), f"**{c[1]}** {c[2]}", c[3], act, f"{CA.BOULE_P[prio]} {prio}", qui, CA.ECHEANCE[prio]]


ordre = {"P1": 0, "P2": 1, "P3": 2, "P4": 3}
plan_ratios = [a for a in CA.PLAN if a[2].startswith("REG-1")]
plan_act = sorted([a for a in CA.PLAN if a not in plan_ratios], key=lambda a: ordre[a[3]])

cur = Curseur(trouver("Ratios prudentiels et régulateurs", "15"))
normal(cur, "**Travaux du volet actuariel.** Le plan de mise en conformité est la première action : sans lui, le FPPN reste exposé au retrait de son agrément.")
legende(cur, "Tableau : plan d'actions priorisé, ratios prudentiels (volet actuariel)")
tableau(cur, ENT, [ligne_plan(k, a) for k, a in enumerate(sorted(plan_ratios, key=lambda a: ordre[a[3]]), 1)], LARG, taille=7.5, boules=(4,))
source(cur, CA.RAP + ", " + CA.PLAN_RENVOI.split(", ", 1)[1] + ".")

cur = Curseur(trouver("Analyse actuarielle", "15"))
normal(cur, "**Huit actions du volet actuariel sont à engager dans les trois mois, dont sept dans ce tableau.** Elles portent sur le taux et les intérêts de l'épargne, le calcul des provisions, la réserve de liquidité et la transmission des données. Elles conditionnent toutes les autres.")
legende(cur, "Tableau : plan d'actions priorisé, analyse actuarielle")
tableau(cur, ENT, [ligne_plan(k, a) for k, a in enumerate(plan_act, len(plan_ratios) + 1)], LARG, taille=7.5, boules=(4,))
p_ = normal(cur, "Priorité : ", taille=Pt(8.5), gras=True)
for pr_, lib_ in (("P1", "dans les 3 mois"), ("P2", "dans les 6 mois"), ("P3", "de 6 à 18 mois"), ("P4", "au-delà de 18 mois")):
    runs(p_, "● ", Pt(10), gras=True, couleur=COUL[CA.BOULE_P[pr_]])
    runs(p_, f"{pr_} {lib_}    ", Pt(8.5))
source(cur, CA.RAP + ", " + CA.PLAN_RENVOI.split(", ", 1)[1] + ".")

d.save(SORTIE)
print(SORTIE)
