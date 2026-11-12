# -*- coding: utf-8 -*-
"""
FPPN - Diapositives v17 (12/11/2026) : évolution des dettes par type, argent disponible face aux paiements de l'année
(lectures A et B), intérêts promis sur l'épargne. Règles : REGLES_REDACTION_FPPN.md (titres-messages, langage simple,
tableaux en 12 pt, graphiques natifs modifiables, sources numérotées en bas de page, ajouts en bleu).

Module réutilisable : ajouter(prs, fond, dettes) ajoute les diapositives à une présentation Adenka.
Exécution directe : python3 diapos_dettes_v17_fppn.py <rapport.pptx> <classeur.xlsx> <sortie.pptx>
  construit un aperçu autonome (mêmes fond et charte que le rapport).
"""
import copy
import io
import re
import sys

import openpyxl
import pandas as pd
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

BLEU, NOIR, GRIS, ORANGE = "0070C0", "060606", "808080", "EFAC62"
SERIES = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948"]
TYPES = ["Épargne des policiers (PEP et EFC)", "Provisions maladie et décès (figées)", "Factures des cliniques et hôpitaux",
         "Factures de fonctionnement", "Fournisseurs d'investissement", "Allocations décès et frais funéraires dus",
         "Anciennes opérations immobilières", "Impôts, CNPS, personnel et autres"]
FICHIERS = {2020: "BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx", 2021: "BALANCE FPPN FORMAT EXCEL DU 01-01-2021 AU 31-12-2026.xlsx",
            2022: "BALANCE FPPN 2022.xlsx", 2023: "BALANCE FPPN 2023.xlsx", 2024: "BALANCE FPPN 2024.xlsx", 2025: "BALANCE FPPN 2025.xlsx",
            2026: "BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx"}
MISE_EN_PAGE = {}  # mise en page des diapositives de contenu du rapport (fond blanc), fixée par apercu()
SIGLES = "Sigles : FPPN : Fonds de prévoyance de la Police nationale. PEP : Plan Épargne Plus. EFC : Entraide fin de carrière. CNPS : Caisse nationale de prévoyance sociale."


def nb(x, d=0):
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def pc(x):
    return nb(100 * x, 2) + " %"


# ------------------------------------------------------------------ données
def type_compte(c):
    if c.startswith("41195"):
        return TYPES[0]
    if c.startswith(("15801", "15802", "19011", "19012")):
        return TYPES[1]
    if c[0] != "4":
        return None
    if c.startswith("40110"):
        return TYPES[2]
    if c.startswith(("401", "403", "408")):
        return TYPES[3]
    if c.startswith(("481", "482")):
        return TYPES[4]
    if c.startswith("4159"):
        return TYPES[5]
    if c.startswith(("472", "47111200", "47120")):
        return TYPES[6]
    return TYPES[7]


def lire_dettes(classeur):
    ws = openpyxl.load_workbook(classeur, read_only=True, data_only=True)["D1_balances"]
    rows = list(ws.iter_rows(min_row=10, values_only=True))
    df = pd.DataFrame(rows[1:], columns=list(rows[0]))
    df["Compte"] = df.Compte.astype(str)
    df["s"] = df["Solde débit"].fillna(0) - df["Solde crédit"].fillna(0)
    df["type"] = df.Compte.map(type_compte)
    d = df[df.type.notna()].copy()
    d["dette"] = (-d.s).clip(lower=0)
    t = d.pivot_table(index="type", columns="Exercice", values="dette", aggfunc="sum").fillna(0)
    return t.loc[TYPES]


# ------------------------------------------------------------------ mise en page
def _runs(p, texte, taille, couleur, gras_defaut=False):
    """**gras**, [[surligné jaune]]."""
    for morceau in re.split(r"(\*\*.+?\*\*|\[\[.+?\]\])", texte):
        if not morceau:
            continue
        gras = gras_defaut or morceau.startswith("**")
        jaune = morceau.startswith("[[")
        t = morceau.strip("*") if morceau.startswith("**") else (morceau[2:-2] if jaune else morceau)
        r = p.add_run()
        r.text = t
        f = r.font
        f.name, f.size, f.bold = "Helvetica", Pt(taille), gras
        f.color.rgb = RGBColor.from_string(couleur)
        if jaune:
            rpr = r._r.get_or_add_rPr()
            hl = etree.SubElement(rpr, qn("a:highlight"))
            etree.SubElement(hl, qn("a:srgbClr")).set("val", "FFFF00")


def texte(slide, x, y, w, h, lignes, taille, couleur=BLEU, gras=False, nom=None, align=PP_ALIGN.LEFT, ecart=4):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    if nom:
        tb.name = nom
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.1)
    for i, l in enumerate(lignes if isinstance(lignes, list) else [lignes]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(ecart)
        _runs(p, l, taille, couleur, gras)
    return tb


def nouvelle(prs, fond, titre, sous_titre, sources, sigles=SIGLES):
    s = prs.slides.add_slide(MISE_EN_PAGE.get("layout") or next(l for l in prs.slide_layouts if l.name == "Title Only"))
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    if fond is not None:
        s.shapes.add_picture(io.BytesIO(fond), 0, 0, prs.slide_width, MISE_EN_PAGE.get("h_fond") or prs.slide_height)
    texte(s, 1.52, 0.4, 31.24, 1.1, titre, 20, BLEU, True, "Adenka_Titre")
    texte(s, 2.07, 1.55, 31.24, 0.85, sous_titre, 16, BLEU, True, "Adenka_SousTitre")
    texte(s, 1.02, 15.55, 31.75, 1.5, ["Sources : " + sources[0]] + sources[1:], 8, GRIS, nom="Adenka_Source", ecart=1)
    texte(s, 1.02, 17.05, 31.75, 0.5, sigles, 7.5, GRIS, ecart=0)
    texte(s, 0.62, 17.77, 12.93, 0.73, "Empreinte Africaine, Expertise Internationale", 11, NOIR, ecart=0)
    return s


def tableau(slide, x, y, w, entetes, lignes, largeurs, taille=12, h_ligne=0.95):
    n, m = len(lignes) + 1, len(entetes)
    gt = slide.shapes.add_table(n, m, Cm(x), Cm(y), Cm(w), Cm(h_ligne * n))
    t = gt.table
    for j, l in enumerate(largeurs):
        t.columns[j].width = Cm(l)
    for i in range(n):
        t.rows[i].height = Cm(h_ligne)
        for j in range(m):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Cm(0.15)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            v = entetes[j] if i == 0 else lignes[i - 1][j]
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 or i == 0 and j == 0 else (PP_ALIGN.CENTER if i == 0 else PP_ALIGN.RIGHT)
            _runs(p, str(v), taille, "FFFFFF" if i == 0 else BLEU, i == 0 or j == 0)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string("060606" if i == 0 else ("D9D9D9" if j == 0 else ("F2F2F2" if i % 2 == 0 else "FFFFFF")))
    return gt


def courbes(slide, x, y, w, h, categories, series, couleurs, titre_axe):
    cd = CategoryChartData(number_format="#,##0")
    cd.categories = categories
    for nom, vals in series:
        cd.add_series(nom, [round(v) for v in vals])
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Cm(x), Cm(y), Cm(w), Cm(h), cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.RIGHT
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(12), "Helvetica"
    ch.font.size, ch.font.name = Pt(12), "Helvetica"
    va = ch.value_axis
    va.tick_labels.number_format = '#,##0,,'
    va.tick_labels.number_format_is_linked = False
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor.from_string("E5E5E5")
    va.format.line.color.rgb = RGBColor.from_string("BBBCBC")
    va.has_title = True
    va.axis_title.text_frame.text = titre_axe
    va.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
    ch.category_axis.format.line.color.rgb = RGBColor.from_string("BBBCBC")
    for s, c in zip(ch.plots[0].series, couleurs):
        s.smooth = False
        s.format.line.color.rgb = RGBColor.from_string(c)
        s.format.line.width = Pt(2.25)
        s.marker.style = XL_MARKER_STYLE.CIRCLE
        s.marker.size = 7
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = RGBColor.from_string(c)
        s.marker.format.line.color.rgb = RGBColor.from_string("FFFFFF")
    return gf


# ------------------------------------------------------------------ diapositives
def ajouter(prs, fond, t):
    ans = list(t.columns)
    cats = [str(a) for a in ans[:-1]] + ["31/08/2026"]
    v = lambda typ, a: float(t.loc[typ, a])
    src_bal = "« " + "», « ".join(FICHIERS[a] for a in ans) + " »"
    ep, pr, cl, al, inv, imp = (t.loc[x] for x in (TYPES[0], TYPES[1], TYPES[2], TYPES[5], TYPES[4], TYPES[7]))

    # 1. toutes les dettes
    s = nouvelle(prs, fond, "La plus grosse dette du FPPN, c'est l'épargne des policiers",
                 f"Fin août 2026, le FPPN doit {nb(ep[2026])} FCFA d'épargne aux policiers, {pc(ep[2026] / ep[2023] - 1)} de plus que fin 2023",
                 [f"(1) Balances des comptes du FPPN, {src_bal} : soldes créditeurs des comptes de dettes (classe 4) et de provisions (15801000, 15802000, 19011000, 19012000), au 31/12 de chaque année et au 31/08/2026 (balance provisoire).",
                  "(2) Épargne : comptes 41195100 (PEP) et 41195200 (EFC) des mêmes balances."])
    courbes(s, 1.4, 2.6, 31.2, 8.6, cats, [(n, t.loc[n].values) for n in TYPES], SERIES, "Millions de FCFA")
    texte(s, 1.4, 11.4, 31.2, 4.0,
          [f"**L'épargne des policiers est de loin la plus grosse dette du FPPN (2).** Elle a baissé jusqu'en 2023, puis elle a fortement monté : {nb(ep[2023])} FCFA fin 2023, {nb(ep[2026])} FCFA fin août 2026.",
           f"**Les provisions pour soins et décès n'ont jamais bougé depuis 2020 (1) :** {nb(pr[2026])} FCFA chaque année. Le FPPN ne les calcule pas, elles ne suivent donc pas ce qu'il doit vraiment.",
           "**Toutes les autres dettes, prises ensemble, restent sous 2,2 milliards FCFA depuis 2022 (1).** La diapositive suivante les montre de plus près."], 13)

    # 2. zoom dettes de court terme
    s = nouvelle(prs, fond, "Les factures de soins et les allocations décès dues repartent à la hausse",
                 f"Factures des cliniques à payer : {nb(cl[2025])} FCFA fin 2025, {nb(cl[2026])} FCFA fin août 2026",
                 [f"(1) Balances des comptes du FPPN, {src_bal} : comptes 40110000 (cliniques), 401, 403 et 408 (fonctionnement), 481 (investissement), 41591000 et 41592000 (allocations décès et frais funéraires dus), 472 (anciennes opérations), 42 à 44 (impôts, CNPS, personnel)."])
    courbes(s, 1.4, 2.6, 31.2, 8.6, cats, [(n, t.loc[n].values) for n in TYPES[2:]], SERIES[2:], "Millions de FCFA")
    texte(s, 1.4, 11.4, 31.2, 4.1,
          [f"**Les factures des cliniques et hôpitaux à payer ont presque triplé en huit mois (+{pc(cl[2026] / cl[2025] - 1)}) (1).** Le FPPN doit expliquer cette hausse : retard de paiement ou hausse des soins.",
           f"**Les familles de policiers décédés attendent de plus en plus d'argent :** {nb(al[2021])} FCFA dus fin 2021, {nb(al[2026])} FCFA fin août 2026, soit {pc(al[2026] / al[2021] - 1)} de plus (1).",
           f"**Deux mouvements restent à expliquer par le FPPN :** un pic de {nb(inv[2024])} FCFA de factures d'investissement fin 2024, presque soldé en 2025, et [[{nb(imp[2021] - imp[2022])} FCFA de dettes d'impôts et de cotisations sociales disparues entre 2021 et 2022, sans savoir si elles ont été payées ou annulées]] (1)."], 13)

    # 3 et 4. tableaux découpés
    for k, (a0, a1) in enumerate(((0, 4), (4, 7)), 1):
        sel = ans[a0:a1]
        s = nouvelle(prs, fond, f"Ce que le FPPN doit, type de dette par type de dette ({k}/2)",
                     "Montants exacts en FCFA, au 31 décembre de chaque année" + (" et au 31 août 2026" if k == 2 else ""),
                     [f"(1) Balances des comptes du FPPN, {'« ' + '», « '.join(FICHIERS[a] for a in sel) + ' »'} : soldes créditeurs des comptes cités à la diapositive précédente."])
        lignes = [[n] + [nb(t.loc[n, a]) for a in sel] for n in TYPES] + [["Total"] + [nb(t[a].sum()) for a in sel]]
        ent = ["Type de dette (1)"] + [("31/08/2026" if a == 2026 else f"31/12/{a}") for a in sel]
        lw = [11.2] + [20.0 / len(sel)] * len(sel)
        tableau(s, 1.4, 2.8, 31.2, ent, lignes, lw, 12, 1.15)

    # 5. lectures A et B
    dispo, dettes_ct, prov, besoin, sorties = 3457866951, 2188438207, 4829298290, 19581392, 3028376687
    a_payer, b_payer = dettes_ct + prov + sorties, dettes_ct + besoin + sorties
    s = nouvelle(prs, fond, "Le FPPN ne peut pas payer tout ce qu'il doit dans l'année",
                 f"Pour 100 FCFA à payer dans l'année, le FPPN n'en a que 66 ({pc(dispo / b_payer)}), et 34 selon son bilan",
                 ["(1) « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx » : banques, caisse, dépôts à terme, fonds commun de placement et créances à moins d'un an (actif disponible), dettes de la classe 4 hors épargne, comptes 15801000 et 15802000 (provisions).",
                  "(2) Estimation Adenka des soins et décès déjà survenus et pas encore payés, sur l'historique des paiements (onglets 19b et 20_risques_courts du classeur des calculs). Le FPPN ne calcule pas ce montant.",
                  "(3) Épargne due au 31/08/2026 (comptes 41195100 et 41195200) multipliée par le taux moyen de retrait de 2021 à 2025 (19,24 %), balances 2021 à 2026 (onglet 24a_ecoulement_epargne du classeur des calculs)."])
    tableau(s, 1.4, 2.8, 31.2, ["", "A. Selon le bilan du FPPN", "B. Ce qui sera vraiment payé"],
            [["Dettes à payer dans l'année, hors épargne (1)", nb(dettes_ct), nb(dettes_ct)],
             ["Soins et décès en attente de paiement", f"{nb(prov)} (provisions figées) (1)", f"{nb(besoin)} (estimation) (2)"],
             ["Épargne que les policiers vont retirer (3)", nb(sorties), nb(sorties)],
             ["**Total à payer dans l'année**", f"**{nb(a_payer)}**", f"**{nb(b_payer)}**"],
             ["Argent disponible dans l'année (1)", nb(dispo), nb(dispo)],
             ["**Pour 100 FCFA à payer, le FPPN a**", f"**{nb(100 * dispo / a_payer)} FCFA ({pc(dispo / a_payer)})**", f"**{nb(100 * dispo / b_payer)} FCFA ({pc(dispo / b_payer)})**"],
             ["**Argent qui manque**", f"**{nb(a_payer - dispo)}**", f"**{nb(b_payer - dispo)}**"]],
            [11.2, 10.0, 10.0], 12, 1.0)
    texte(s, 1.4, 11.1, 31.2, 4.3,
          ["**Les provisions pour soins et décès du bilan ne correspondent à aucun calcul (1).** Elles sont figées depuis 2020. Les soins et décès vraiment en attente sont estimés à 19 581 392 FCFA (2), et les sommes dues aux familles sont déjà dans les dettes.",
           f"**Ce qu'il faut retenir, c'est la lecture B :** il manque environ un tiers de ce qu'il faudra payer dans l'année, soit {nb(b_payer - dispo)} FCFA. La lecture A sert aux ratios du régulateur, calculés sur le bilan tel qu'il est tenu.",
           "[[Une troisième lecture ajoutera les intérêts promis sur l'épargne retirée dans l'année, une fois le taux fixé par le FPPN.]]"], 13)

    # 6. intérêts promis
    s = nouvelle(prs, fond, "Le FPPN a promis des intérêts sur l'épargne, sans jamais les inscrire",
                 "Si la promesse est tenue, le FPPN doit entre 2 541 799 542 et 3 278 913 238 FCFA d'intérêts aux policiers",
                 ["(1) « MANUEL DE PRODUIT D'EPARGNE FPPN VF VF.docx » (certificat d'adoption non signé), « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx », « CAHIER DE CHARGEEFC & PEP 2021.pdf », reçus le 22/10/2026. Taux de 3,5 % déclaré par la direction comptable le 19/10/2026, sans document.",
                  "(2) Balances 2020 au 31/08/2026 : aucun compte de charge ni de dette d'intérêts sur l'épargne. Seuls figurent des agios bancaires (63710200) et des intérêts d'emprunt en 2020 et 2021 (67120001).",
                  "(3) Calcul Adenka sur les soldes moyens des comptes 41195100 et 41195200, du 01/01/2020 au 31/08/2026, sans capitalisation (onglet 21 du classeur des calculs)."])
    tableau(s, 1.4, 2.8, 31.2, ["Ce que disent les documents (1)", "Taux promis", "Intérêts dus au 31/08/2026 (3)"],
            [["Tout à 3,5 % (direction comptable, à l'oral)", "PEP et EFC : 3,5 %", "2 541 799 542"],
             ["Cartographie des prestations : EFC à 5,5 % à partir de 2024", "PEP 3,5 %, EFC 3,5 % puis 5,5 %", "[[2 954 084 955]]"],
             ["Manuel des produits d'épargne", "PEP 3,5 %, EFC 5,5 %", "3 278 913 238"],
             ["Cahier des charges de 2021", "EFC : 14 % d'intérêt simple, période non dite", "Non calculé"]],
            [12.6, 9.6, 9.0], 12, 1.05)
    texte(s, 1.4, 8.4, 31.2, 7.0,
          ["**Le FPPN a promis par écrit de rémunérer l'épargne des policiers (1).** Trois de ses documents donnent un taux, mais aucun n'est validé par une décision signée de ses organes.",
           "**Rien n'a jamais été inscrit dans ses comptes pour ces intérêts (2).** La règle comptable veut qu'une somme due soit inscrite l'année où elle naît, même si elle n'est payée que plus tard.",
           "**Si cette promesse est tenue, le FPPN doit entre 2,5 et 3,3 milliards FCFA aux policiers, selon le taux (3).** Cela s'ajoute aux 15 743 516 391 FCFA d'épargne déjà dus.",
           "**Ce que le FPPN doit faire :** ses organes fixent le taux. Ensuite, soit il inscrit ces intérêts dans ses comptes, soit il modifie ses documents et en informe les policiers.",
           "[[Vérification en cours : comparer ce que les policiers partis à la retraite ont versé et ce qu'ils ont reçu, pour savoir si des intérêts sont réellement payés.]]"], 13)


def apercu(rapport, classeur, sortie):
    prs = Presentation(rapport)
    fond = None
    for sl in prs.slides:
        if not any(sh.has_text_frame and sh.name == "Adenka_SousTitre" for sh in sl.shapes) or not any(getattr(sh, "has_table", False) and sh.has_table for sh in sl.shapes):
            continue
        pics = [sh for sh in sl.shapes if sh.shape_type == 13 and abs(sh.left) < 50000 and abs(sh.top) < 50000 and sh.width >= prs.slide_width * 0.98]
        if pics:
            fond = pics[0].image.blob
            MISE_EN_PAGE["layout"] = sl.slide_layout
            MISE_EN_PAGE["h_fond"] = pics[0].height
            break
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        prs.part.drop_rel(sid.get(qn("r:id")))
        lst.remove(sid)
    ajouter(prs, fond, lire_dettes(classeur))
    prs.save(sortie)
    # nettoyage de structure (compteur de diapositives, auteurs sans commentaire) avec le vérificateur Adenka
    import glob
    import subprocess
    v = glob.glob("/root/.claude/skills/synced/*/adenka-verification-pptx/scripts/verif_pptx.py")
    if v:
        subprocess.run([sys.executable, v[0], sortie, "--reparer", "--sortie", sortie], capture_output=True)
    print(sortie, len(prs.slides), "diapositives")


if __name__ == "__main__":
    apercu(*sys.argv[1:4])
