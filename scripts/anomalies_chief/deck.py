# -*- coding: utf-8 -*-
"""Deck de 6 diapositives (charte Adenka) sur les deux anomalies : dépenses hors soins et provision de la clinique.
Utilisation : python deck.py <dossier de travail> <sortie.pptx>"""
import json
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Pt

SK = os.path.join(os.path.expanduser("~"), ".claude", "skills", "adenka-powerpoint")
sys.path.insert(0, os.path.join(SK, "scripts"))
import pptx_tools as T  # noqa: E402

import pandas as pd  # noqa: E402

W, OUT = sys.argv[1], sys.argv[2]
if os.path.exists(OUT):
    raise SystemExit("le fichier existe déjà : " + OUT)
R = json.load(open(os.path.join(W, "resultats.json"), encoding="utf-8"))
TB = pd.read_parquet(os.path.join(W, "t1_avant_apres.parquet"))
DET = pd.read_parquet(os.path.join(W, "t1_detail.parquet"))
ANS = list(TB.index)
CAT = [str(a) if a != 2026 else "2026 (8 mois)" for a in ANS]
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
GL = R["sources"]["grands_livres"]
BAL = R["sources"]["balances"]
f = lambda v: f"{v:,.0f}".replace(",", " ")
p = lambda v: f"{100 * v:.2f}".replace(".", ",")

T.selection(os.path.join(SK, "assets", "Timesaver_Adenka.pptx"), OUT, [5] * 6)
prs = Presentation(OUT)
SW, SH = prs.slide_width, prs.slide_height


def texte(slide, x, y, l, h, blocs, taille=13, align="justifie", nom=None):
    """Blocs séparés par \n. **gras**, [[surligné jaune]]."""
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(l), Cm(h))
    if nom:
        tb.name = nom
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    al = {"justifie": PP_ALIGN.JUSTIFY, "gauche": PP_ALIGN.LEFT, "centre": PP_ALIGN.CENTER}[align]
    for i, bloc in enumerate(blocs.split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = al
        para.space_after = Pt(7)
        for morceau in re.split(r"(\*\*.+?\*\*|\[\[.+?\]\])", bloc):
            if not morceau:
                continue
            r = para.add_run()
            gras, jaune = morceau.startswith("**"), morceau.startswith("[[")
            r.text = morceau[2:-2] if (gras or jaune) else morceau
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = gras
            r.font.color.rgb = RGBColor.from_string("060606")
            if jaune:
                rpr = r._r.get_or_add_rPr()
                hl = etree.SubElement(rpr, A + "highlight")
                etree.SubElement(hl, A + "srgbClr").set("val", "FFFF00")
                # ordre du schéma : highlight avant latin
                lat = rpr.find(A + "latin")
                if lat is not None:
                    rpr.remove(hl)
                    lat.addprevious(hl)
    return tb


def tableau(slide, x, y, l, lignes, larg, taille=12, h_ligne=1.0, gras_col0=True):
    nr, nc = len(lignes), len(lignes[0])
    gf = slide.shapes.add_table(nr, nc, Cm(x), Cm(y), Cm(l), Cm(h_ligne * nr))
    t = gf.table
    tbl = gf._element.graphic.graphicData.tbl
    tp = tbl.tblPr
    for k in ("bandRow", "firstRow"):
        tp.set(k, "0")
    st = tp.find(A + "tableStyleId")
    if st is not None:
        st.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # sans style, quadrillage
    for j, w in enumerate(larg):
        t.columns[j].width = Cm(w)
    for i, row in enumerate(lignes):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.15)
            c.margin_top = c.margin_bottom = Cm(0.08)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            jaune = isinstance(v, str) and v.startswith("[[")
            v = v[2:-2] if jaune else v
            tf.paragraphs[0].text = ""
            r = tf.paragraphs[0].add_run()
            r.text = str(v)
            r.font.name = "Helvetica"
            r.font.size = Pt(taille)
            r.font.bold = (i == 0) or (j == 0 and gras_col0)
            r.font.color.rgb = RGBColor.from_string("FFFFFF" if i == 0 else "060606")
            tf.paragraphs[0].alignment = PP_ALIGN.LEFT if j == 0 or not re.match(r"^[\d\s,%-]+$", str(v)) else PP_ALIGN.RIGHT
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string("060606" if i == 0 else ("FFFF00" if jaune else "FFFFFF"))
            tcPr = c._tc.get_or_add_tcPr()
            for bord in ("lnL", "lnR", "lnT", "lnB"):
                ln = etree.SubElement(tcPr, A + bord, w="6350")
                sf = etree.SubElement(ln, A + "solidFill")
                etree.SubElement(sf, A + "srgbClr").set("val", "BBBCBC")
            # les bordures doivent précéder le remplissage dans tcPr
            fill = tcPr.find(A + "solidFill")
            if fill is not None:
                tcPr.remove(fill)
                tcPr.append(fill)
    return gf


def pied(slide, src, sigles):
    T.source(slide, src)
    texte(slide, 1.02, 16.55, 31.7, 0.5, "Surligné en jaune : point à confirmer par une pièce ou une réponse écrite du FPPN. Sigles : " + sigles, taille=7, align="gauche")


def base(n, titre, st, syn):
    s = prs.slides[n]
    T.vider(s, SW, SH)
    T.entete(s, titre, st, syn, SW, SH)
    return s


SIG = "FPPN : Fonds de prévoyance de la Police nationale"
t26 = TB.loc[2026]
t25 = TB.loc[2025]
e = R["effets_2026"]
M = R["clinique_mois"]
G25, G26, B26 = GL["2025"], GL["2026"], BAL["2026"]
XL = "« FPPN_Anomalies_soins_provision_20261117.xlsx »"

# ------------------------------------------------------------------ 1 Synthèse
s = base(0, "Deux anomalies à corriger", "Les dépenses de soins sont gonflées par des dons et par des factures comptées deux fois",
         f"Corrigée, la maladie consomme {p(e['sc_mal'][2])} % de ses cotisations sur huit mois de 2026, et non {p(e['sc_mal'][0])} %")
tableau(s, 1.4, 4.0, 18.6, [
    ["Anomalie", "Montant", "Effet sur le ratio maladie", "À confirmer"],
    ["1. Dons, cadeaux, imprimés et honoraires payés comme des soins", f"{f(t25.hors_soins)} FCFA en 2025\n{f(t26.hors_soins)} FCFA sur 8 mois de 2026",
     f"2025 : de {p(t25.sc_mal_avant)} % à {p(t25.sc_mal_apres)} %", "Pièces des honoraires du cabinet"],
    ["2. Factures de la clinique d'octobre à décembre 2025 comptées en 2025 puis en 2026", f"{f(R['provision'])} FCFA",
     f"2026 : de {p(e['sc_mal'][0])} % à {p(e['sc_mal'][1])} %", "[[Lettrage du compte 40110000]]"],
    ["Les deux ensemble, 2026 (8 mois)", f"{f(R['provision'] + t26.hors_soins)} FCFA", f"2026 : de {p(e['sc_mal'][0])} % à {p(e['sc_mal'][2])} %",
     "[[Une facture mensuelle de la clinique à retrouver en 2026]]"]],
    [5.6, 4.6, 4.2, 4.2], h_ligne=2.4)
texte(s, 20.6, 4.0, 11.6, 12.3,
      f"**Les deux anomalies grossissent les soins et donnent une image trop sombre de la garantie maladie.** Sans elles, la maladie dépense "
      f"{p(e['sc_mal'][2])} FCFA de soins pour 100 FCFA de cotisations sur huit mois de 2026, soit {p(e['sc_mal'][2])} % des cotisations maladie, au lieu de {p(e['sc_mal'][0])} %.\n"
      "La première anomalie est un mauvais classement : dons, cadeaux, imprimés et honoraires sont payés sur les comptes de soins. Ils restent des dépenses du Fonds, mais ce ne sont pas des soins.\n"
      f"La seconde est un oubli d'écriture : la provision (argent mis de côté) de fin 2025 n'a pas été annulée quand les vraies factures sont arrivées en 2026. "
      f"Les dépenses de 2026 et la dette envers les cliniques sont trop lourdes de {f(R['provision'])} FCFA.")
pied(s, f"Sources : grands livres « {GL['2021']} » à « {G26} », balances « {BAL['2021']} » à « {B26} ». Calculs Adenka, classeur {XL}", SIG)

# ------------------------------------------------------------------ 2 Hors soins par type
s = base(1, "Dépenses hors soins par année", f"Les dépenses hors soins payées comme des soins atteignent {round(t25.hors_soins / 1e6)} millions de FCFA en 2025",
         f"En 2025, elles font {p(t25.part_hs)} FCFA sur 100 FCFA de prestations maladie, soit {p(t25.part_hs)} % des prestations maladie")
ty = DET[DET.categorie != "Soins"].pivot_table(index="categorie", columns="an", values="montant", aggfunc="sum").reindex(columns=ANS).fillna(0)
COUL = {"Dons (autorités, élus, délégués)": "EFAC62", "Gadgets, imprimés et cartes": "060606", "Honoraires d'un cabinet": "BBBCBC", "Régularisations comptables": "595959"}
texte(s, 1.4, 3.85, 18.2, 0.7, "Dépenses hors soins rangées dans les comptes de soins, par type", taille=12, align="centre").text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("EFAC62")
texte(s, 1.4, 4.45, 8, 0.5, "En millions de FCFA", taille=9, align="gauche")
T.graphique(s, {"type": "barres_empilees", "categories": CAT,
                "series": [{"nom": c, "valeurs": [round(ty.loc[c, a] / 1e6, 1) for a in ANS], "couleur": COUL[c]} for c in COUL],
                "format": "0", "x": 1.4, "y": 5.0, "largeur": 18.2, "hauteur": 11.3, "legende": "bas", "etiquettes": False, "taille_police": 11, "axe_min": 0})
c25 = ty[2025]
HON = "Honoraires d'un cabinet"
texte(s, 20.6, 4.0, 11.6, 12.3,
      f"**Les dons sont la plus grosse dépense hors soins : {f(c25['Dons (autorités, élus, délégués)'])} FCFA en 2025.** S'y ajoutent "
      f"{f(c25['Gadgets, imprimés et cartes'])} FCFA de gadgets, imprimés et cartes, et {f(c25[HON])} FCFA d'honoraires d'un cabinet, facturés tous les trois mois.\n"
      "Chaque ligne du grand livre est classée par son libellé (DON, CADEAU, GADGET, IMPRIME, facture n° 00444...). Les dons à caractère médical (dialyse, médicaments, accouchement) restent dans les soins.\n"
      "Avant 2023, les dons passaient surtout par un autre compte (6583 : 61 894 670 FCFA en 2022). "
      f"Une variante plus large (hadj, lobbying, cérémonies, autres imprimeurs) ajoute {f(t25.variante)} FCFA en 2025 sans changer la conclusion.")
pied(s, f"Sources : grands livres « {GL['2021']} » à « {G26} », comptes 6011 à 6080, classement Adenka par libellé. Classeur {XL}, onglets 1_Par_type et 1_Detail", SIG)

# ------------------------------------------------------------------ 3 Sinistralité avant / après
s = base(2, "Sinistralité avant et après", "Sans les dépenses hors soins, la maladie consomme 53 % de ses cotisations en 2025, et non 74 %",
         f"En 2025, {p(t25.sc_mal_apres)} FCFA de soins pour 100 FCFA de cotisations maladie, soit {p(t25.sc_mal_apres)} %, contre {p(t25.sc_mal_avant)} % dans les comptes")
texte(s, 1.4, 3.85, 18.2, 0.7, "Soins payés pour 100 FCFA de cotisations maladie", taille=12, align="centre").text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("EFAC62")
texte(s, 1.4, 4.45, 8, 0.5, "En % des cotisations maladie", taille=9, align="gauche")
T.graphique(s, {"type": "barres", "categories": CAT,
                "series": [{"nom": "Comptes tels quels", "valeurs": [round(100 * TB.loc[a, "sc_mal_avant"], 2) for a in ANS], "couleur": "BBBCBC"},
                           {"nom": "Sans les dépenses hors soins", "valeurs": [round(100 * TB.loc[a, "sc_mal_apres"], 2) for a in ANS], "couleur": "EFAC62"}],
                "format": "0.0", "x": 1.4, "y": 5.0, "largeur": 18.2, "hauteur": 11.3, "legende": "bas", "taille_police": 11})
texte(s, 20.6, 4.0, 11.6, 12.3,
      f"**La garantie maladie se porte mieux que ne le disent les comptes.** En 2025, le ratio maladie passe de {p(t25.sc_mal_avant)} % à {p(t25.sc_mal_apres)} % des cotisations maladie. "
      f"Sur huit mois de 2026, il passe de {p(t26.sc_mal_avant)} % à {p(t26.sc_mal_apres)} %.\n"
      f"Le ratio global du texte (toutes prestations pour 100 FCFA de toutes cotisations, norme 70 % à 80 %) passe de {p(t25.sc_glob_avant)} % à {p(t25.sc_glob_apres)} % en 2025. Il était déjà sous la norme.\n"
      "Ces dépenses ne disparaissent pas. Le Fonds les a payées et elles doivent être classées en frais de fonctionnement. Le résultat de l'année ne change pas. Pour 2026, la seconde anomalie s'ajoute (diapositive 5).")
pied(s, f"Sources : grands livres « {GL['2021']} » à « {G26} » (comptes 601 et 60800007), balances « {BAL['2021']} » à « {B26} » (comptes 7010 et 70). "
        f"Décision n° 05/2022/COM/UEMOA, annexe, § 1.2.1.2.5. Classeur {XL}, onglet 1_Avant_apres", SIG + " | UEMOA : Union économique et monétaire ouest-africaine")

# ------------------------------------------------------------------ 4 Clinique mois par mois
s = base(3, "Provision de la clinique fin 2025", "Les factures de la clinique d'octobre à décembre 2025 sont comptées en 2025 puis en 2026",
         f"{f(R['provision'])} FCFA mis de côté fin 2025 pour {f(R['factures'])} FCFA de vraies factures, sans annulation en 2026")
texte(s, 1.4, 3.85, 18.2, 0.7, "Argent mis de côté et vraie facture, par mois de soins", taille=12, align="centre").text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("EFAC62")
texte(s, 1.4, 4.45, 8, 0.5, "En millions de FCFA", taille=9, align="gauche")
T.graphique(s, {"type": "barres", "categories": [m["mois"] for m in M],
                "series": [{"nom": "Mis de côté le 31/12/2025 (OD 805, 806, 807)", "valeurs": [round(m["provision"] / 1e6, 2) for m in M], "couleur": "060606"},
                           {"nom": "Vraie facture passée le 31/01/2026 (SIN 462, 552, 619)", "valeurs": [round(m["facture"] / 1e6, 2) for m in M], "couleur": "EFAC62"}],
                "format": "0.00", "x": 1.4, "y": 5.0, "largeur": 18.2, "hauteur": 11.3, "legende": "bas", "taille_police": 11, "axe_min": 0})
texte(s, 20.6, 4.0, 11.6, 12.3,
      "**Le 31/12/2025, le Fonds attend les factures de sa clinique pour octobre, novembre et décembre.** Il met de côté 85 000 000 FCFA par mois, soit "
      f"{f(R['provision'])} FCFA, en dépense de 2025 et en dette envers la clinique (OD 805, 806 et 807). C'est la bonne pratique.\n"
      f"Le 31/01/2026, les vraies factures arrivent : {f(M[0]['facture'])}, {f(M[1]['facture'])} et {f(M[2]['facture'])} FCFA, soit {f(R['factures'])} FCFA (SIN 462, 552 et 619). "
      "Le Fonds les passe en dépense de 2026 sans annuler la provision. Le grand livre 2026 ne contient aucune écriture d'annulation.\n"
      f"Les mêmes soins sont donc comptés deux fois. La provision dépassait les factures de {f(R['boni_2025'])} FCFA, écart normal pour une estimation. "
      "[[À confirmer par le lettrage du compte 40110000, qui relie chaque facture à son paiement.]]", taille=12.5)
pied(s, f"Sources : « {G25} », OD 805, 806 et 807 du 31/12/2025, comptes 60120000 et 40110000. « {G26} », pièces SIN 462, 552 et 619 du 31/01/2026. "
        f"Classeur {XL}, onglets 2_Clinique_mois et 2_Ecritures", SIG + " | OD : journal des opérations diverses | SIN : journal des factures de soins")

# ------------------------------------------------------------------ 5 Effet 2026
s = base(4, "Effet sur 2026", "Corrigée des deux anomalies, la maladie consomme 60 % de ses cotisations en 2026, et non 97 %",
         f"Sur huit mois de 2026, {p(e['sc_mal'][2])} FCFA de soins pour 100 FCFA de cotisations maladie, soit {p(e['sc_mal'][2])} % des cotisations maladie")
pm = [t26.prest_maladie, t26.prest_maladie - R["provision"], t26.prest_maladie - R["provision"] - t26.hors_soins_dans_maladie]
tableau(s, 1.4, 4.0, 18.6, [
    ["Au 31/08/2026", "Comptes tels quels", "Sans la provision comptée deux fois", "Et sans les dépenses hors soins"],
    ["Ratio maladie (%)"] + [p(v) for v in e["sc_mal"]],
    ["Ratio global, norme 70 % à 80 % (%)"] + [p(v) for v in e["sc_glob"]],
    ["Prestations maladie (FCFA)"] + [f(v) for v in pm],
    ["Factures des cliniques à payer, compte 40110000 (FCFA)", f(690234451), f(690234451 - R["provision"]), f(690234451 - R["provision"])]],
    [6.2, 4.0, 4.2, 4.2], h_ligne=1.6)
texte(s, 1.4, 12.6, 18.6, 3.6,
      f"Les autres factures de 2025 passées en 2026 sont faibles : {f(R['autres_2025'])} FCFA sûres, et {f(R['autres_2025_mixtes'])} FCFA qui mêlent 2025 et 2026. "
      "C'est un minimum, car beaucoup de libellés n'écrivent pas l'année des soins.", taille=12)
texte(s, 20.6, 4.0, 11.6, 12.3,
      f"**L'annulation de la provision retire {f(R['provision'])} FCFA des dépenses de 2026.** La maladie passe de {p(e['sc_mal'][0])} % à {p(e['sc_mal'][1])} % de ses cotisations. "
      f"Sans les dépenses hors soins ({f(t26.hors_soins)} FCFA), elle descend à {p(e['sc_mal'][2])} %. Le ratio global passe sous la norme : {p(e['sc_glob'][2])} %.\n"
      f"La dette envers les cliniques est trop lourde du même montant : {f(690234451 - R['provision'])} FCFA au lieu de {f(690234451)} FCFA.\n"
      "[[Pour huit mois de 2026, le grand livre ne montre que sept factures mensuelles de la clinique.]] Si une facture manque, une provision est à passer au 31/08/2026, "
      "et le ratio corrigé de 2026 serait alors trop bas.")
pied(s, f"Sources : « {G26} » (comptes 60 et 40110000), « {B26} » (comptes 70 et 40110000, solde au 31/08/2026). Calculs Adenka, classeur {XL}, onglets 3_Ensemble_2026 et 2_Autres_2025", SIG)

# ------------------------------------------------------------------ 6 Sources et méthode
s = base(5, "Sources et méthode", "Chaque chiffre vient des grands livres et des balances du Fonds, recalculé ligne par ligne",
         "Les prestations recalculées sont égales aux balances du Fonds, compte par compte, de 2021 à 2026")
tableau(s, 1.4, 4.0, 31.0, [
    ["Chiffre", "Fichier exact du Fonds", "Comment il est obtenu"],
    ["Prestations par année", f"« {GL['2021']} » à « {G26} », un grand livre par année (liste exacte dans l'onglet Sources du classeur)",
     "Somme des lignes des comptes 60, à-nouveaux exclus. Total égal au mouvement de la balance de chaque année."],
    ["Cotisations par année", f"« {BAL['2021']} » à « {B26} »", "Mouvement crédit moins mouvement débit des comptes 7010 (maladie) et 70 (toutes)."],
    ["Dépenses hors soins", "Mêmes grands livres, colonne « Libellé écriture »", "Classement de chaque ligne des comptes 6011 à 6080 par les mots du libellé. Liste ligne par ligne dans le classeur."],
    ["Provision et factures de la clinique", f"« {G25} » et « {G26} »", "OD 805, 806 et 807 du 31/12/2025, pièces SIN 462, 552 et 619 du 31/01/2026. Recherche d'une annulation dans tout 2026 : aucune."],
    ["Dette envers les cliniques", f"« {B26} »", "Solde du compte 40110000 au 31/08/2026, lu tel quel : 690 234 451 FCFA."],
    ["Norme de sinistralité", "Décision n° 05/2022/COM/UEMOA, annexe, § 1.2.1.2.5", "Ratio global de 70 % à 80 % des cotisations."]],
    [6.0, 12.0, 13.0], h_ligne=1.3)
texte(s, 1.4, 15.0, 31.0, 1.4,
      f"**Les calculs sont faits en Python puis repris en formules dans le classeur {XL}.** Chaque mémo de cellule explique le chiffre en mots simples. "
      "Le FPPN ne calcule pas ces ratios corrigés : ce sont des estimations Adenka.", taille=12)
pied(s, "Source : documents du FPPN cités dans le tableau. Traitements Adenka", SIG + " | UEMOA : Union économique et monétaire ouest-africaine")

prs.save(OUT)
T.finaliser(OUT)
print("écrit", OUT)
