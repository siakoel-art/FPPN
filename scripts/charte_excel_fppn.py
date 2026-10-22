# -*- coding: utf-8 -*-
"""
FPPN - Charte graphique Adenka appliquée au classeur des calculs (version 8)
===========================================================================

Règles reprises de la charte des tableaux Adenka (skill adenka-rapport-word, adenka_docx.py) et de la palette
Adenka (noir 060606, bleu nuit 112A40, orange EFAC62, gris BBBCBC) :
- police Helvetica partout ; titre de l'onglet 14 pt gras noir ; note sous le titre 9 pt italique gris ;
- tableau : titre numéroté au-dessus (« Tableau : … »), source en dessous ;
- en-tête noir (000000), texte blanc gras centré ; première colonne grise (CCCCCC) en gras ; lignes alternées gris
  clair (E8E8E8) ; bordures fines noires ; chiffres alignés à droite ;
- pourcentages à deux décimales, montants en FCFA sans décimale, années sans séparateur ;
- cellules modifiables (paramètres) en orange (EFAC62) ; lignes de section en gris (CCCCCC) ;
- statuts : « Conforme » sur fond bleu nuit, « Non conforme » sur fond rouge (C00000), texte blanc gras ;
- quadrillage masqué, impression en paysage sur une page de large.

Utilisation : appliquer(wb) depuis classeur_v2_fppn.py, ou python charte_excel_fppn.py <classeur.xlsx> (modifié en place).
"""
import re
import sys

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

POLICE = "Helvetica"
NOIR, BLEU, ORANGE, GRIS_TXT, ROUGE = "060606", "112A40", "EFAC62", "53565A", "C00000"
F_ENTETE, F_COL1, F_ALT = "000000", "CCCCCC", "E8E8E8"
ENTETES = {"00000000", "FF000000", "00060606", "FF060606", "00112A40", "FF112A40"}
PARAMS = {"00FFF2CC", "FFFFF2CC", "00EFAC62", "FFEFAC62"}
SECTIONS = {"00F2F2F2", "FFF2F2F2", "00CCCCCC", "FFCCCCCC"}
FIN = Side(style="thin", color="000000")
BORD = Border(left=FIN, right=FIN, top=FIN, bottom=FIN)
RUBRIQUES_EXPLIC = ("À quoi sert cet onglet", "D'où viennent les chiffres", "Comment le calcul est fait", "Comment lire le résultat", "Valable pour", "Limites")
GRIS_RUB = "53565A"

SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 ; traitements Adenka"
SRC_BAL = "Source : balances des comptes du FPPN (Sage 100), 2020 au 31/08/2026 (2026 provisoire) ; traitements Adenka"
SRC_GRE = "Source : exports du logiciel de prêts du Grenier du Policier reçus le 15/10/2026 ; traitements Adenka"
SRC_DEC = "Source : balances du FPPN ; Décision n° 05/2022/COM/UEMOA, annexe ; règlement d'exécution n° 004/2022/COM/UEMOA ; calculs Adenka"
TITRES = {
    "1_couverture": ("Couverture des fichiers de précomptes reçus, par poste et par sens", SRC_PREC),
    "1_aller_retour": ("Rapprochement des fichiers aller et retour, par poste", SRC_PREC),
    "1_rejets": ("Lignes écartées des traitements, par motif", SRC_PREC),
    "2_effectifs_mensuels": ("Cotisants et retraités précomptés, par mois", SRC_PREC),
    "2_entrees_sorties": ("Entrées et sorties mensuelles de cotisants", SRC_PREC),
    "2_taux_936": ("Taux du précompte maladie (poste 936) rapporté au salaire de base", SRC_PREC),
    "2_rapport_8Z6_936": ("Rapport des précomptes décès (8Z6) et maladie (936)", SRC_PREC),
    "2_anciennete": ("Ancienneté des agents cotisants", SRC_PREC),
    "2_cotisations": ("Cotisations comptabilisées et précomptées, par exercice", SRC_BAL + " ; fichiers de précomptes"),
    "3_sinistralite_resultat": ("Sinistralité et résultat, par exercice", SRC_BAL),
    "3_prestations_maladie_nature": ("Prestations maladie par nature (FCFA)", SRC_BAL),
    "4_bilan": ("Postes du bilan utilisés par le rapport (FCFA)", SRC_BAL),
    "4_pep_versements": ("Versements au Plan Épargne Plus", SRC_PREC + " ; balances (compte 41195100)"),
    "4_efc_versements": ("Versements à l'Entraide de Fin de Carrière", SRC_PREC + " ; balances (compte 41195200)"),
    "4_poste_998_references": ("Références du poste 998 (Plan Épargne Plus)", SRC_PREC),
    "5_grenier_encours": ("Encours des prêts du Grenier", SRC_GRE),
    "5_grenier_retards": ("Encours du Grenier par tranche de retard", SRC_GRE),
    "5_grenier_par_historique": ("Portefeuille en retard du Grenier (plus de 30 et de 90 jours)", SRC_GRE),
    "5_grenier_production": ("Production du Grenier par type de prêt (FCFA)", SRC_GRE),
    "5_grenier_echeances_precomptes": ("Échéances du Grenier et précomptes du poste 9X4", SRC_GRE + " ; fichiers de précomptes"),
    "6_cgrae": ("Retraités précomptés par la CGRAE, par mois", SRC_PREC),
    "6_produits_annexes": ("Produits annexes : flotte automobile et MEDOC", SRC_PREC),
    "7_rendement_placements": ("Rendement apparent des placements", SRC_BAL),
    "8_inflation_reel": ("Cotisations et prestations en valeur réelle (base 2020)", "Source : Banque mondiale, FP.CPI.TOTL.ZG ; " + SRC_BAL[9:]),
    "8_tcam": ("Taux de croissance annuel moyen des principaux indicateurs", SRC_BAL),
    "8_tendances": ("Tendances statistiques (régressions linéaires)", SRC_BAL),
    "8_sensibilite_retraits": ("Sensibilité de la trésorerie à des retraits d'épargne", SRC_BAL),
    "8_sensibilite_pertes_grenier": ("Sensibilité aux pertes sur les prêts du Grenier", SRC_GRE),
    "8_provisions_mois": ("Provisions techniques en mois de prestations", SRC_BAL),
    "8_tables_mortalite": ("Indicateurs des tables de mortalité CIMA", "Source : tables CIMA H, CIMA F, TD et TV CIMA (fichiers fournis le 18/10/2026) ; calculs Adenka"),
    "10_ratios": ("Ratios prudentiels de la Décision n° 05/2022, 2020 au 31/08/2026", SRC_DEC),
    "10_ratios_detail": ("Numérateurs, dénominateurs, valeurs et statuts des ratios", SRC_DEC),
    "10_classement_comptes": ("Classement des comptes selon le PCMS et selon la nature des risques", SRC_BAL + " ; PCMS (mai 2012)"),
    "10_marge_solvabilite": ("Marge de solvabilité exigée et disponible", SRC_DEC),
    "10_affectation_excedents": ("Affectation des excédents et dotation à la réserve obligatoire", SRC_DEC),
    "10_excedent_par_risque": ("Excédent par risque (maladie, décès)", SRC_DEC),
    "10_controle_pcms": ("Contrôle du classement des comptes selon le PCMS", "Source : PCMS (mai 2012) ; balances du FPPN"),
    "10_actif_mobilisation": ("Actif par délai de mobilisation et passif à moins d'un an", SRC_BAL),
    "10_agregats": ("Agrégats comptables utilisés par les ratios (FCFA)", SRC_BAL),
    "12_provisions_inventaire": ("Inventaire des provisions et dépréciations (FCFA)", SRC_BAL),
    "12_psap_rapprochement": ("Provision maladie, dettes envers les cliniques et prestations", SRC_BAL),
    "12_grenier_tranches": ("Encours du Grenier par tranche de retard", SRC_GRE),
    "12_gna_47111106": ("Compte de tiers GNA (47111106)", SRC_BAL),
    "12_pidr_196": ("Provision pour départ à la retraite du personnel (compte 196)", SRC_BAL),
    "12_sinistralite_variantes": ("Sinistralité : formule du texte et variantes", SRC_DEC),
    "9_tracabilite": ("Traçabilité : chaque chiffre du rapport, sa diapositive et sa valeur exacte", "Source : rapport PowerPoint et contrôle automatique des chiffres"),
    "9_figures_tableaux": ("Liste des figures et des tableaux du rapport", "Source : rapport PowerPoint"),
    "11_formules": ("Registre des formules du rapport", "Source : formules_fppn.py ; textes cités dans chaque ligne"),
}
SANS_FORMAT = {"10_ratios", "10_ratios_detail", "9_tracabilite", "9_figures_tableaux", "11_formules", "10_controle_pcms"}
PCT = ("taux", "part", "sc_", "rendement", "couverture", "tcam", "croissance", "par30", "par90", "cle_", "fidelisation",
       "recouvrement", "psap_sur", "ic95", "inflation")
ENTIER = ("annee", "exercice", "mois", "poste", "numero", "numéro", "diapositive", "m")
SIGLES_ENT = {"pep": "PEP", "efc": "EFC", "cgrae": "CGRAE", "pcms": "PCMS", "fppn": "FPPN", "psap": "PSAP", "sc": "S/C", "ic95": "IC95",
              "r2": "R²", "par30": "PAR30", "par90": "PAR90", "tcam": "TCAM", "cmp": "CMP", "gna": "GNA", "medoc": "MEDOC", "dat": "DAT",
              "cot": "cotisations", "prest": "prestations", "deces": "décès", "hao": "HAO", "resultat": "résultat", "epargne": "épargne",
              "prets": "prêts", "echeances": "échéances", "debloque": "débloqué", "precompte": "précompte", "precomptes": "précomptes",
              "retraites": "retraités", "annee": "année", "tresorerie": "trésorerie", "reserve": "réserve", "exigee": "exigée",
              "reserves": "réserves", "obligatoire": "obligatoire", "dotation": "dotation", "excedent": "excédent", "debit": "débit",
              "credit": "crédit", "depreciation": "dépréciation", "interets": "intérêts", "periode": "période", "entrees": "entrées",
              "anciennete": "ancienneté", "elements": "éléments", "immediat": "immédiat", "cloture": "clôture", "arrete": "arrêté",
              "denominateur": "dénominateur", "numerateur": "numérateur", "reelles": "réelles", "reels": "réels", "nominales": "nominales",
              "dedies": "dédiés", "ecart": "écart", "rubrique": "rubrique", "intitule": "intitulé", "variation": "variation"}


def _long_affiche(v, fmt):
    if v is None:
        return 0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if fmt == "0.00%":
            return len(f"{100 * v:,.2f}") + 2
        if fmt in ("#,##0",):
            return len(f"{v:,.0f}")
        if fmt == "0":
            return len(str(int(v)))
        return len(f"{v:,.2f}")
    return len(str(v))


def _rgb(c):
    try:
        return c.fill.fgColor.rgb if c.fill and c.fill.fill_type == "solid" else None
    except Exception:
        return None


def _police(c, taille=None, gras=None, italique=None, couleur=None):
    f = c.font
    c.font = Font(name=POLICE, size=taille or (f.sz if f and f.sz else 10), bold=f.b if gras is None else gras,
                  italic=f.i if italique is None else italique, color=couleur or (f.color if f and f.color else None))


def lisible(nom):
    if not isinstance(nom, str) or " " in nom or not re.match(r"^[a-z0-9_]+$", nom):
        return nom
    mots = [SIGLES_ENT.get(m, m) for m in nom.split("_")]
    t = " ".join(mots)
    return t[0].upper() + t[1:]


def _format_colonne(nom, valeurs):
    n = str(nom).lower()
    nums = [v for v in valeurs if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not nums:
        return None
    if n in ENTIER or n.startswith("annee") or n.startswith("exercice"):
        return "0"
    if any(k in n for k in PCT) and all(abs(v) < 50 for v in nums):
        return "0.00%"
    if all(float(v).is_integer() for v in nums) or max(abs(v) for v in nums) >= 1000:
        return "#,##0"
    return "0.00"


def preparer_brut(ws):
    """Onglet issu de pandas (en-tête en ligne 1) : titre, bloc d'explication en mots simples, en-têtes lisibles, formats imposés."""
    titre, source = TITRES.get(ws.title, (ws.title.split("_", 1)[-1].replace("_", " ").capitalize(), "Source : traitements Adenka"))
    try:
        from explications_excel_fppn import EXPLIC
    except ImportError:
        EXPLIC = {}
    rubs = [(r_, EXPLIC[ws.title][r_]) for r_ in RUBRIQUES_EXPLIC if EXPLIC.get(ws.title, {}).get(r_)]
    o = len(rubs) + 1 if rubs else 0
    ncol, nlig = ws.max_column, ws.max_row
    filtre = ws.auto_filter.ref
    gel = ws.freeze_panes
    ws.insert_rows(1, 2 + o)
    for j in range(1, ncol + 1):
        c = ws.cell(3 + o, j)
        nom = c.value
        if ws.title not in SANS_FORMAT:
            fmt = _format_colonne(nom, [ws.cell(i, j).value for i in range(4 + o, nlig + 3 + o)])
            if fmt:
                for i in range(4 + o, nlig + 3 + o):
                    if isinstance(ws.cell(i, j).value, (int, float)):
                        ws.cell(i, j).number_format = fmt
        c.value = lisible(nom)
        c.fill = PatternFill("solid", fgColor=F_ENTETE)
    ws.cell(1, 1, f"Tableau : {titre}")
    ws.cell(1, 1).font = Font(name=POLICE, size=12, bold=True, color=NOIR)
    derniere = openpyxl.utils.get_column_letter(max(ncol, 6))
    for k, (rub, texte) in enumerate(rubs):
        r = 2 + k
        a = ws.cell(r, 1, rub)
        a.font = Font(name=POLICE, size=9, bold=True, color="FFFFFF")
        a.fill = PatternFill("solid", fgColor=GRIS_RUB)
        a.alignment = Alignment(vertical="top", wrap_text=True)
        ws.merge_cells(f"B{r}:{derniere}{r}")
        b = ws.cell(r, 2, texte)
        b.font = Font(name=POLICE, size=9, color=NOIR)
        b.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = max(15, 12.5 * (len(texte) // 110 + 1))
    s = ws.cell(nlig + 4 + o, 1, source)
    s.font = Font(name=POLICE, size=8, italic=True, color=GRIS_TXT)
    if filtre:
        a, b = filtre.split(":")
        ca, cb = re.sub(r"\d", "", a), re.sub(r"\d", "", b)
        ws.auto_filter.ref = f"{ca}{3 + o}:{cb}{int(re.sub(r'[^0-9]', '', b)) + 2 + o}"
    if gel:
        col = re.sub(r"\d", "", gel) or "A"
        ws.freeze_panes = f"{col}{4 + o}"
    else:
        ws.freeze_panes = f"B{4 + o}"
    for j in range(1, ncol + 1):
        lettre = openpyxl.utils.get_column_letter(j)
        entete_ = len(str(ws.cell(3 + o, j).value or ""))
        long_ = max([_long_affiche(ws.cell(i, j).value, ws.cell(i, j).number_format) for i in range(4 + o, min(nlig + 3 + o, 400))] + [0])
        actuelle = ws.column_dimensions[lettre].width or 0
        voulue = min(50, max(10, long_ * 1.15 + 3, min(entete_, 22) * 1.05 + 2))
        ws.column_dimensions[lettre].width = max(actuelle, voulue) if ws.title in SANS_FORMAT else voulue
    if rubs:
        ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 24)
    ws.row_dimensions[3 + o].height = 30


def styliser_tableaux(ws):
    """Repère chaque tableau par sa ligne d'en-tête (fond noir ou bleu nuit) et applique la charte."""
    fus = {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                fus[(r, c)] = (m.min_row, m.min_col)
    vus = set()
    for r in range(1, ws.max_row + 1):
        c = 1
        while c <= ws.max_column:
            if _rgb(ws.cell(r, c)) in ENTETES and (r, c) not in vus:
                c2 = c
                while c2 + 1 <= ws.max_column and _rgb(ws.cell(r, c2 + 1)) in ENTETES:
                    c2 += 1
                for j in range(c, c2 + 1):
                    x = ws.cell(r, j)
                    vus.add((r, j))
                    x.fill = PatternFill("solid", fgColor=F_ENTETE)
                    x.font = Font(name=POLICE, size=10, bold=True, color="FFFFFF")
                    x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    x.border = BORD
                rr = r + 1
                while rr <= ws.max_row:
                    vals = [ws.cell(rr, j).value for j in range(c, c2 + 1)]
                    if all(v is None or v == "" for v in vals) and not any((rr, j) in fus and fus[(rr, j)] != (rr, j) for j in range(c, c2 + 1)):
                        break
                    if any(_rgb(ws.cell(rr, j)) in ENTETES for j in range(c, c2 + 1)):
                        break
                    section = _rgb(ws.cell(rr, c)) in SECTIONS
                    for j in range(c, c2 + 1):
                        x = ws.cell(rr, j)
                        fx = _rgb(x)
                        x.border = BORD
                        rouge = x.font and x.font.color is not None and str(x.font.color.rgb).endswith(ROUGE)
                        if fx in PARAMS:
                            x.fill = PatternFill("solid", fgColor=ORANGE)
                            x.font = Font(name=POLICE, size=10, bold=True, color=NOIR)
                        elif section:
                            x.fill = PatternFill("solid", fgColor=F_COL1)
                            x.font = Font(name=POLICE, size=10, bold=True, color=NOIR)
                        else:
                            if j == c:
                                x.fill = PatternFill("solid", fgColor=F_COL1)
                            elif (rr - r) % 2 == 0:
                                x.fill = PatternFill("solid", fgColor=F_ALT)
                            else:
                                x.fill = PatternFill(fill_type=None)
                            x.font = Font(name=POLICE, size=10, bold=True if j == c else (x.font.b if x.font else False),
                                          color=ROUGE if rouge else NOIR)
                        v = x.value
                        num = isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=") and x.number_format not in ("General", "@"))
                        x.alignment = Alignment(horizontal="right" if num else (x.alignment.horizontal or "left"),
                                                vertical="top", wrap_text=x.alignment.wrap_text if x.alignment else False)
                    rr += 1
                c = c2 + 1
            else:
                c += 1


def formats_ratios(wb):
    """Onglets 10_ratios et 10_ratios_detail : format de chaque valeur selon l'unité de la norme du ratio."""
    if "10_ratios" not in wb.sheetnames:
        return
    ws = wb["10_ratios"]
    h = next(r for r in range(1, 40) if "Code" in [ws.cell(r, j).value for j in range(1, 4)])
    ent = {ws.cell(h, j).value: j for j in range(1, ws.max_column + 1)}
    unite = {}
    for i in range(h + 1, ws.max_row + 1):
        code, norme = ws.cell(i, ent["Code"]).value, str(ws.cell(i, ent["Norme"]).value or "")
        if not code:
            continue
        pct = "%" in norme and "mois" not in norme
        unite[code] = pct
        for nom, j in ent.items():
            v = ws.cell(i, j).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and nom not in ("Priorité",):
                ws.cell(i, j).number_format = "0.00%" if pct else ("#,##0" if abs(v) >= 1000 else "0.00")
    if "10_ratios_detail" in wb.sheetnames:
        wd = wb["10_ratios_detail"]
        hd = next(r for r in range(1, 40) if wd.cell(r, 1).value == "Code")
        e = {wd.cell(hd, j).value: j for j in range(1, wd.max_column + 1)}
        for i in range(hd + 1, wd.max_row + 1):
            code = wd.cell(i, e["Code"]).value
            for nom in ("Numérateur", "Dénominateur"):
                if isinstance(wd.cell(i, e[nom]).value, (int, float)):
                    wd.cell(i, e[nom]).number_format = "#,##0"
            v = wd.cell(i, e["Valeur"]).value
            if isinstance(v, (int, float)):
                wd.cell(i, e["Valeur"]).number_format = "0.00%" if unite.get(code) else ("#,##0" if abs(v) >= 1000 else "0.00")
            if isinstance(wd.cell(i, e["Exercice"]).value, (int, float)):
                wd.cell(i, e["Exercice"]).number_format = "0"


def appliquer(wb):
    for ws in wb.worksheets:
        brut = ws.title in TITRES and ws.title not in ("10_tableau_de_bord",)
        if brut:
            preparer_brut(ws)
        styliser_tableaux(ws)
        for row in ws.iter_rows():
            for x in row:
                if x.value is not None and (x.font is None or x.font.name != POLICE):
                    _police(x)
        if not brut and isinstance(ws.cell(1, 1).value, str) and ws.title != "0_lisez_moi":
            ws.cell(1, 1).font = Font(name=POLICE, size=14, bold=True, color=NOIR)
            if isinstance(ws.cell(2, 1).value, str):
                ws.cell(2, 1).font = Font(name=POLICE, size=9, italic=True, color=GRIS_TXT)
        if ws.title == "10_tableau_de_bord":
            ws["B2"].font = Font(name=POLICE, size=14, bold=True, color=NOIR)
            ws["B3"].font = Font(name=POLICE, size=9, italic=True, color=GRIS_TXT)
        dim = ws.dimensions
        if ws.title.startswith("13"):
            ws.conditional_formatting.add(dim, FormulaRule(formula=['A1="Conforme"'], fill=PatternFill("solid", fgColor=BLEU), font=Font(name=POLICE, color="FFFFFF", bold=True)))
            ws.conditional_formatting.add(dim, FormulaRule(formula=['A1="Non conforme"'], fill=PatternFill("solid", fgColor=ROUGE), font=Font(name=POLICE, color="FFFFFF", bold=True)))
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    formats_ratios(wb)
    if "10_tableau_de_bord" in wb.sheetnames:
        for c in "FGHIJ":
            wb["10_tableau_de_bord"].column_dimensions[c].width = 17
    if "13i_rendement" in wb.sheetnames:
        for c in "CDEFGHIJKL":
            wb["13i_rendement"].column_dimensions[c].width = 15
    lm = wb["0_lisez_moi"] if "0_lisez_moi" in wb.sheetnames else None
    if lm:
        lm["A1"].font = Font(name=POLICE, size=14, bold=True, color=NOIR)
        for r in range(2, lm.max_row + 1):
            lm.cell(r, 1).font = Font(name=POLICE, size=10, color=NOIR)
            lm.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        lm.sheet_view.showGridLines = False
        for nom in wb.sheetnames:
            if nom.startswith(("13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25")):
                wb[nom].sheet_properties.tabColor = ORANGE
            elif nom.startswith(("00", "10", "11", "12")):
                wb[nom].sheet_properties.tabColor = BLEU


if __name__ == "__main__":
    w = openpyxl.load_workbook(sys.argv[1])
    appliquer(w)
    w.save(sys.argv[1])
