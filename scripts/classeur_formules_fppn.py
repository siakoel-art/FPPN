# -*- coding: utf-8 -*-
"""
FPPN - Classeur des calculs : données de base et formules vivantes
==================================================================
Remplace les onglets de valeurs 1 à 12 par des onglets à formules qui puisent dans des onglets de données brutes :

  D1_balances               : une ligne par compte et par exercice (à-nouveaux, mouvements, soldes), fichier source
  D2_fichiers_precomptes    : une ligne par fichier de précompte et par mois (poste, sens, lignes, montant)
  D3_precomptes_mensuels    : totaux par poste, sens et mois ; nombres d'agents distincts (calcul Python)
  D4_grenier_prets          : une ligne par prêt du Grenier et par arrêté (encours, retard, échéance)
  D5_grenier_production     : une ligne par prêt débloqué
  D6_rejets                 : lignes écartées des fichiers, par motif
  D7_externe                : inflation (Banque mondiale)
  D8_tables_mortalite       : nombre de survivants lx des tables CIMA H, CIMA F, TD et TV
  10_agregats               : chaque agrégat comptable = somme des comptes cités (formules SUMIFS), par exercice

Les colonnes qui exigent les lignes individuelles des fichiers (nombre d'agents distincts, médianes, entrées et sorties)
sont calculées par Python : les 4 454 416 lignes de précomptes retenues dépassent la limite d'Excel (1 048 576 lignes
par feuille). Ces colonnes sont sur fond bleu clair et nomment le script qui les produit.

Chaque valeur recalculée est comparée, dans l'onglet 0_controle, à la valeur publiée dans le rapport.
Utilisation : appelé par classeur_v2_fppn.py (fonction reconstruire).
"""
import math
import os
import re

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

POLICE = "Helvetica"
NOIR, BLEU, ORANGE, GRIS_TXT, ROUGE = "060606", "112A40", "EFAC62", "53565A", "C00000"
F_COL1, F_ALT, F_PY, F_RUB = "CCCCCC", "E8E8E8", "DCE6F1", "53565A"
FIN = Side(style="thin", color="000000")
BORD = Border(left=FIN, right=FIN, top=FIN, bottom=FIN)
EX = list(range(2020, 2027))
FCFA, PCT, NB2, ENT, DATE = "#,##0", "0.00%", "0.00", "0", "DD/MM/YYYY"
SCRIPT_BASE = "base_fppn.py (lecture et contrôle des fichiers)"
SCRIPT_ANA = "analyses_fppn.py"
RUBRIQUES = ("À quoi sert cet onglet", "D'où viennent les chiffres", "Comment le calcul est fait", "Exemple chiffré",
             "Comment lire le résultat", "Valable pour", "Limites")
CTRL = []          # (onglet, cellule, valeur du rapport) pour l'onglet 0_controle
CREES = []         # onglets créés par ce module (exclus de la mise en forme automatique)
TEXTES = {}        # explications de chaque onglet créé (reprises dans l'onglet 0_guide)


def q(nom):
    return "'" + nom + "'"


# ---------------------------------------------------------------------------
# Écriture d'un onglet : titre, explication, bandeau Python, tableau, source
# ---------------------------------------------------------------------------
class Onglet:
    def __init__(self, wb, nom, titre, index=None, explic=None, python=None, ncol=8):
        self.ws = wb.create_sheet(nom, index) if index is not None else wb.create_sheet(nom)
        self.nom, self.ws.sheet_view.showGridLines = nom, False
        CREES.append(nom)
        TEXTES[nom] = dict(explic or {}, **({"Calcul Python": python} if python else {}))
        ws = self.ws
        ws.cell(1, 1, f"Tableau : {titre}").font = Font(name=POLICE, size=12, bold=True, color=NOIR)
        r = 2
        derniere = L(max(ncol, 6))
        for rub in RUBRIQUES:
            t = (explic or {}).get(rub)
            if not t:
                continue
            a = ws.cell(r, 1, rub)
            a.font, a.fill = Font(name=POLICE, size=9, bold=True, color="FFFFFF"), PatternFill("solid", fgColor=F_RUB)
            a.alignment = Alignment(vertical="top", wrap_text=True)
            ws.merge_cells(f"B{r}:{derniere}{r}")
            b = ws.cell(r, 2, t)
            b.font, b.alignment = Font(name=POLICE, size=9, color=NOIR), Alignment(vertical="top", wrap_text=True)
            ws.row_dimensions[r].height = max(15, 12.5 * (len(t) // 120 + 1))
            r += 1
        if python:
            a = ws.cell(r, 1, "Calcul Python")
            a.font, a.fill = Font(name=POLICE, size=9, bold=True, color=NOIR), PatternFill("solid", fgColor=F_PY)
            a.alignment = Alignment(vertical="top", wrap_text=True)
            ws.merge_cells(f"B{r}:{derniere}{r}")
            b = ws.cell(r, 2, python)
            b.font, b.fill = Font(name=POLICE, size=9, color=NOIR), PatternFill("solid", fgColor=F_PY)
            b.alignment = Alignment(vertical="top", wrap_text=True)
            ws.row_dimensions[r].height = max(15, 12.5 * (len(python) // 120 + 1))
            r += 1
        self.r = r + 1
        self.largeurs = {}

    def tableau(self, entetes, lignes, formats=None, py_cols=(), orange_cols=(), source=None, largeurs=None, total=None):
        """lignes : listes de valeurs ou de formules (chaînes « = » ; « {r} » = ligne courante).
        py_cols : colonnes calculées par Python (fond bleu clair). Renvoie (ligne d'en-tête, première, dernière ligne de données)."""
        ws, h = self.ws, self.r
        for j, t in enumerate(entetes, 1):
            c = ws.cell(h, j, str(t))
            c.font, c.fill = Font(name=POLICE, size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="000000")
            c.alignment, c.border = Alignment(horizontal="center", vertical="center", wrap_text=True), BORD
        ws.row_dimensions[h].height = 30
        d = h + 1
        for i, lig in enumerate(lignes):
            r = d + i
            for j, v in enumerate(lig, 1):
                if isinstance(v, str) and v.startswith("="):
                    v = v.replace("{r}", str(r))
                if isinstance(v, float) and math.isnan(v):
                    v = None
                c = ws.cell(r, j, v)
                c.border = BORD
                fond = F_COL1 if j == 1 else (F_ALT if i % 2 else None)
                if (j - 1) in py_cols:
                    fond = F_PY
                if (j - 1) in orange_cols:
                    fond = ORANGE
                if fond:
                    c.fill = PatternFill("solid", fgColor=fond)
                c.font = Font(name=POLICE, size=10, bold=(j == 1), color=NOIR)
                num = isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("="))
                c.alignment = Alignment(horizontal="right" if num else "left", vertical="top")
                if formats and j - 1 < len(formats) and formats[j - 1]:
                    c.number_format = formats[j - 1]
        f = d + len(lignes) - 1
        if total:
            r = f + 1
            for j, v in enumerate(total, 1):
                if isinstance(v, str) and v.startswith("="):
                    v = v.replace("{d}", str(d)).replace("{f}", str(f)).replace("{r}", str(r))
                c = ws.cell(r, j, v)
                c.font, c.fill, c.border = Font(name=POLICE, size=10, bold=True), PatternFill("solid", fgColor=F_COL1), BORD
                if formats and j - 1 < len(formats) and formats[j - 1]:
                    c.number_format = formats[j - 1]
            f_tot = r
        else:
            f_tot = f
        if source:
            s = ws.cell(f_tot + 2, 1, source)
            s.font = Font(name=POLICE, size=8, italic=True, color=GRIS_TXT)
        for j, t in enumerate(entetes, 1):
            w = (largeurs[j - 1] if largeurs and j - 1 < len(largeurs) else min(40, max(12, len(str(t)) * 1.1 + 2)))
            self.largeurs[j] = max(self.largeurs.get(j, 0), w)
            ws.column_dimensions[L(j)].width = self.largeurs[j]
        if self.ws.column_dimensions["A"].width < 24:
            self.ws.column_dimensions["A"].width = 24
        self.r = f_tot + 4
        return h, d, f

    def note(self, texte):
        c = self.ws.cell(self.r, 1, texte)
        c.font, c.alignment = Font(name=POLICE, size=9, italic=True, color=GRIS_TXT), Alignment(wrap_text=False)
        self.r += 1


def controler(onglet, h, valeurs, cols=None):
    """Enregistre la comparaison cellule par cellule entre le tableau reconstruit et les valeurs publiées.
    valeurs : DataFrame (même ordre de lignes et de colonnes que le tableau) ; cols : colonnes à contrôler (indices)."""
    for i in range(len(valeurs)):
        for j in (cols if cols is not None else range(valeurs.shape[1])):
            v = valeurs.iat[i, j]
            if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and not (isinstance(v, float) and math.isnan(v)):
                CTRL.append((onglet, f"{L(j + 1)}{h + 1 + i}", float(v)))


# ---------------------------------------------------------------------------
# Agrégats comptables : définitions (préfixes de comptes) et formules
# ---------------------------------------------------------------------------
# mode : S solde (débit - crédit), D somme des soldes débiteurs, C somme des soldes créditeurs (en positif),
#        MD mouvements débit, MC mouvements crédit ; signe -1 : produits et ressources présentés en positif.
# Un préfixe « =40110000 » désigne le compte exact.
AGG = [
    # -- cotisations et prestations
    ("cot_maladie_7010", "Cotisations maladie", [(-1, "S", ["7010"])]),
    ("cot_deces_7021", "Cotisations décès", [(-1, "S", ["7021"])]),
    ("prest_maladie", "Prestations maladie", [(1, "S", ["601", "60800007"])]),
    ("prest_deces", "Prestations décès", [(1, "S", ["6021", "6024"])]),
    ("prime_6518", "Prime d'assurance décès (6518)", [(1, "S", ["6518"])]),
    ("charges_61_66", "Charges 61 à 66", [(1, "S", ["61", "62", "63", "64", "65", "66"])]),
    ("frais_gestion", "Frais de gestion (61 à 66 hors 6518)", "{charges_61_66}-{prime_6518}"),
    ("produits_7", "Produits (classe 7)", [(-1, "S", ["7"])]),
    ("charges_6", "Charges (classe 6)", [(1, "S", ["6"])]),
    ("resultat_hao", "Résultat hors activités ordinaires (classe 8)", [(-1, "S", ["8"])]),
    ("resultat_net", "Résultat net", "{produits_7}-{charges_6}+{resultat_hao}"),
    ("nat_pharmacie", "Prestations : pharmacie", [(1, "S", ["60111"])]),
    ("nat_consultations", "Prestations : consultations", [(1, "S", ["60112"])]),
    ("nat_hospitalisations", "Prestations : hospitalisations", [(1, "S", ["60113"])]),
    ("nat_analyses", "Prestations : analyses médicales", [(1, "S", ["60114"])]),
    ("nat_radio", "Prestations : radio et imagerie", [(1, "S", ["60115"])]),
    ("nat_autres_6012", "Prestations : autres prestations maladie (6012)", [(1, "S", ["6012"])]),
    ("nat_autres", "Prestations : autres (transport, rééducation, bons)", [(1, "S", ["60116", "60118", "60119", "6013", "6018", "60800007"])]),
    # -- bilan
    ("prov_maladie", "Provision maladie", [(-1, "S", ["15801", "19011"])]),
    ("prov_obseques", "Provision obsèques", [(-1, "S", ["15802", "19012"])]),
    ("epargne_pep", "Épargne PEP (41195100)", [(-1, "S", ["411951"])]),
    ("epargne_efc", "Épargne EFC (41195200)", [(-1, "S", ["411952"])]),
    ("banque_efc", "Banques et DAT de l'EFC", [(1, "S", ["5211020", "5221020"])]),
    ("banque_pep", "Banques et DAT du PEP", [(1, "S", ["5211030", "5221030", "526120"])]),
    ("fcp", "FCP (5330)", [(1, "S", ["5330"])]),
    ("tresorerie_totale", "Trésorerie (51 à 57)", [(1, "S", ["51", "52", "53", "54", "55", "56", "57"])]),
    ("creance_hopital", "Créance sur l'hôpital de la Police (47111100)", [(1, "S", ["47111100"])]),
    ("cmp", "Investissements au CMP (47111101, 47111104, 47111105)", [(1, "S", ["47111101", "47111104", "47111105"])]),
    ("grenier", "Financement du Grenier (47111102)", [(1, "S", ["47111102"])]),
    ("cot_a_recevoir", "Cotisations à recevoir de l'État (4111100)", [(1, "S", ["4111100"])]),
    ("terrains", "Terrains (222)", [(1, "S", ["222"])]),
    ("immob_nettes", "Immobilisations nettes (classe 2)", [(1, "S", ["2"])]),
    ("fournisseurs_cliniques", "Dettes envers les cliniques (40110)", [(-1, "S", ["40110"])]),
    ("reserves_10_11", "Réserves (10 et 11)", [(-1, "S", ["10", "11"])]),
    ("ran_120_121", "Report à nouveau créditeur (120, 121)", [(-1, "S", ["120", "121"])]),
    ("pertes_129", "Report à nouveau débiteur (129)", [(-1, "S", ["129"])]),
    ("prets_pep_bilan", "Prêts PEP (2711, 27120)", [(1, "S", ["2711", "27120"])]),
    ("total_actif", "Total de l'actif (comptes débiteurs des classes 2 à 5)", [(1, "D", ["2", "3", "4", "5"])]),
    ("produits_fin_77", "Produits financiers (77)", [(-1, "S", ["77"])]),
    ("placements_52_53", "Placements (522 à 526, 533)", [(1, "S", ["522", "523", "524", "525", "526", "533"])]),
    # -- ratios (comptes cités par la Décision n° 05/2022)
    ("cotisations_acquises", "Cotisations acquises (70 + 7906 - 6906)", [(-1, "S", ["70"]), (-1, "S", ["7906"]), (1, "S", ["6906"])]),
    ("cot_maladie_701", "Cotisations maladie (701)", [(-1, "S", ["701"])]),
    ("cot_deces_702", "Cotisations décès (702)", [(-1, "S", ["702"])]),
    ("charges_prestations", "Charges de prestations (60 + 6901 + 7901)", [(1, "S", ["60"]), (1, "S", ["6901"]), (1, "S", ["7901"])]),
    ("charges_fonctionnement", "Charges de fonctionnement (612, 613, 618, 62 à 68)", [(1, "S", ["612", "613", "618", "62", "63", "64", "65", "66", "67", "68"])]),
    ("charges_fonctionnement_brut", "Charges de fonctionnement brut (61 à 68, 691, 697)", [(1, "S", ["61", "62", "63", "64", "65", "66", "67", "68", "691", "697"])]),
    ("charges_personnel", "Charges de personnel (66)", [(1, "S", ["66"])]),
    ("charges_personnel_641", "Impôts et taxes sur salaires (641)", [(1, "S", ["641"])]),
    ("cout_formation", "Coût de formation (6391)", [(1, "S", ["6391"])]),
    ("total_produits", "Total des produits (liste de l'annexe de la Décision)",
     [(-1, "S", ["70", "71", "72", "73", "75", "77", "781", "782", "791", "797", "82", "84", "86", "88"]), (-1, "S", ["7906"]), (1, "S", ["6906"])]),
    ("subventions", "Subventions (751 à 758, 88)", [(-1, "S", ["751", "752", "753", "754", "755", "756", "757", "758", "88"])]),
    ("total_charges", "Total des charges (classe 6, 81, 83, 85)", [(1, "S", ["6"]), (1, "S", ["81", "83", "85"])]),
    ("fonds_mutualistes_10", "Fonds mutualistes (10)", [(-1, "S", ["10"])]),
    ("reserve_disponible_106", "Réserve disponible (106)", [(-1, "S", ["106"])]),
    ("reserves_legales_111", "Réserves légales (111)", [(-1, "S", ["111"])]),
    ("reserves_facultatives_118", "Réserves facultatives (118)", [(-1, "S", ["118"])]),
    ("ran_crediteur_121", "Report à nouveau créditeur (121)", [(-1, "S", ["121"])]),
    ("ran_debiteur_129", "Report à nouveau débiteur (129)", [(-1, "S", ["129"])]),
    ("subventions_invest_14", "Subventions d'investissement (14)", [(-1, "S", ["14"])]),
    ("provisions_techniques", "Provisions techniques (190, 15801, 15802)", [(-1, "S", ["190", "15801", "15802"])]),
    ("provisions_reglementees_15", "Provisions réglementées (15)", [(-1, "S", ["15"])]),
    ("provisions_190", "Provisions techniques (190)", [(-1, "S", ["190"])]),
    ("provision_196", "Provision pour pensions (196)", [(-1, "S", ["196"])]),
    ("emprunts_16", "Emprunts (16)", [(-1, "S", ["16"])]),
    ("epargne", "Épargne PEP et EFC", "{epargne_pep}+{epargne_efc}"),
    ("sorties_epargne", "Sorties d'épargne (mouvements débit 411951 et 411952)", [(1, "MD", ["411951", "411952"])]),
    ("emplois_longs_tiers", "Emplois longs en comptes de tiers (hôpital, CMP, Grenier, 47111201, 485)",
     [(1, "S", ["47111100", "47111101", "47111102", "47111104", "47111105", "47111201", "485"])]),
    ("tresorerie_actif", "Trésorerie à l'actif (comptes débiteurs de la classe 5)", [(1, "D", ["5"])]),
    ("tresorerie_passif", "Trésorerie au passif (comptes créditeurs de la classe 5)", [(1, "C", ["5"])]),
    ("disponibilites", "Disponibilités (comptes débiteurs 521, 531, 532, 57)", [(1, "D", ["521", "531", "532", "57"])]),
    ("placements_tresorerie", "Placements de trésorerie (comptes débiteurs 522, 526, 533, 50)", [(1, "D", ["522", "526", "533", "50"])]),
    ("comptes_courants_521", "Comptes courants (521)", [(1, "S", ["521"])]),
    ("debiteurs_3_4", "Comptes débiteurs des classes 3 et 4", [(1, "D", ["3", "4"])]),
    ("creances_ct", "Créances à court terme", "{debiteurs_3_4}-{emplois_longs_tiers}"),
    ("crediteurs_4", "Comptes créditeurs de la classe 4", [(1, "C", ["4"])]),
    ("dettes_ct_pcms", "Dettes à court terme (classe 4 créditrice + trésorerie au passif)", "{crediteurs_4}+{tresorerie_passif}"),
    ("dettes_ct_hors_epargne", "Dettes à court terme hors épargne", "{dettes_ct_pcms}-{epargne}"),
    ("actifs_dedies_epargne", "Actifs réservés à l'épargne (banques et DAT dédiés, FCP)", "{banque_efc}+{banque_pep}+{fcp}"),
    ("prets_pep_ratios", "Prêts PEP (2711, 2712, 4111031)", [(1, "S", ["2711", "2712", "4111031"])]),
    ("solde_47111201_485", "Autres emplois longs (47111201, 485)", [(1, "S", ["47111201", "485"])]),
    ("fonds_propres", "Fonds propres (10, 111, 118, 121, 129, 14 et résultat)",
     "{fonds_mutualistes_10}+{reserves_legales_111}+{reserves_facultatives_118}+{ran_crediteur_121}+{ran_debiteur_129}+{subventions_invest_14}+{resultat_net}"),
    ("var_provisions_techniques", "Variation des provisions techniques", "VAR:provisions_techniques"),
    # -- provisions et variantes de sinistralité
    ("charges_60", "Charges de prestations (60)", [(1, "S", ["60"])]),
    ("dot_rep_6901_7901", "Dotations et reprises de provisions techniques (6901, 7901)", [(1, "S", ["6901", "7901"])]),
    ("cotisations_70", "Cotisations (70)", [(-1, "S", ["70"])]),
    ("dot_rep_7906_6906", "Variations de cotisations (7906, 6906)", [(-1, "S", ["7906"]), (1, "S", ["6906"])]),
    ("cot_medoc_7022", "Cotisations MEDOC (7022)", [(-1, "S", ["7022"])]),
    ("bons_60800007_8", "Bons de prise en charge (60800007, 60800008)", [(1, "S", ["60800007", "60800008"])]),
    ("dettes_cliniques_40110000", "Dettes envers les cliniques (compte 40110000)", [(-1, "S", ["=40110000"])]),
    ("gna_solde", "Compte GNA 47111106 : solde", [(1, "S", ["=47111106"])]),
    ("gna_debit", "Compte GNA 47111106 : mouvements débit", [(1, "MD", ["=47111106"])]),
    ("gna_credit", "Compte GNA 47111106 : mouvements crédit", [(1, "MC", ["=47111106"])]),
    ("pidr_196", "Provision pour départ à la retraite (196)", [(-1, "S", ["196"])]),
    ("dotation_6911", "Dotation (6911)", [(1, "S", ["6911"])]),
    ("reprise_7911", "Reprise (7911)", [(-1, "S", ["7911"])]),
]


def lib_prefixes(termes):
    morceaux = []
    for s, m, pre in termes:
        mode = {"S": "solde", "D": "comptes débiteurs", "C": "comptes créditeurs", "MD": "mouvements débit", "MC": "mouvements crédit"}[m]
        txt = ", ".join(p[1:] if p.startswith("=") else p + "…" for p in pre)
        morceaux.append(("moins " if s < 0 else "plus ") + f"{mode} de {txt}")
    t = " ; ".join(morceaux)
    return t[0].upper() + t[1:]


class Agregats:
    """Onglet 10_agregats : une ligne par agrégat, une colonne par exercice, formules SUMIFS sur D1_balances."""

    def __init__(self, wb, index, bal_dl, explic):
        self.o = Onglet(wb, "10_agregats", "Agrégats comptables utilisés par tous les calculs, 2020 au 31/08/2026", index, explic, ncol=11)
        self.row, self.col = {}, {}
        d1, f1 = bal_dl
        D = q("D1_balances")
        rng = {k: f"{D}!${c}${d1}:${c}${f1}" for k, c in (("compte", "A"), ("ex", "C"), ("ad", "F"), ("ac", "G"), ("md", "H"), ("mc", "I"), ("net", "L"))}
        ws = self.o.ws
        h = self.o.r
        ent = ["Agrégat (clé)", "Libellé", "Comptes et calcul"] + [str(a) for a in EX]
        lignes = []
        for i, (cle, lib, defi) in enumerate(AGG):
            self.row[cle] = h + 1 + i
        for j, a in enumerate(EX):
            self.col[a] = L(4 + j)
        for cle, lib, defi in AGG:
            ligne = [cle, lib, lib_prefixes(defi) if isinstance(defi, list) else ("Variation d'une année sur l'autre" if defi.startswith("VAR:") else "Calcul : " + re.sub(r"\{(\w+)\}", r"\1", defi))]
            for a in EX:
                yc = str(a)
                if isinstance(defi, list):
                    parts = []
                    for s, m, pre in defi:
                        for p in pre:
                            crit = f'"{p[1:]}"' if p.startswith("=") else f'"{p}*"'
                            if m == "S":
                                e = f"SUMIFS({rng['net']},{rng['compte']},{crit},{rng['ex']},{yc})"
                            elif m == "D":
                                e = f'SUMIFS({rng["net"]},{rng["compte"]},{crit},{rng["ex"]},{yc},{rng["net"]},">0")'
                            elif m == "C":
                                e = f'-SUMIFS({rng["net"]},{rng["compte"]},{crit},{rng["ex"]},{yc},{rng["net"]},"<0")'
                            elif m == "MD":
                                e = f"SUMIFS({rng['md']},{rng['compte']},{crit},{rng['ex']},{yc})"
                            else:
                                e = f"SUMIFS({rng['mc']},{rng['compte']},{crit},{rng['ex']},{yc})"
                            parts.append(("-" if s < 0 else "+") + e)
                    f = "=" + "".join(parts).lstrip("+")
                elif defi.startswith("VAR:"):
                    k = defi[4:]
                    f = "=0" if a == 2020 else f"={self.col[a]}{self.row[k]}-{self.col[a - 1]}{self.row[k]}"
                else:
                    f = "=" + re.sub(r"\{(\w+)\}", lambda m_: f"{self.col[a]}{self.row[m_.group(1)]}", defi)
                ligne.append(f)
            lignes.append(ligne)
        self.h, _, _ = self.o.tableau(ent, lignes, [None, None, None] + [FCFA] * len(EX), largeurs=[26, 44, 60] + [18] * len(EX),
                                      source="Source : onglet D1_balances (balances du FPPN) ; comptes cités par la Décision n° 05/2022/COM/UEMOA (annexe) et par le rapport")
        for j in range(1, 4):
            for i in range(len(AGG)):
                ws.cell(self.h + 1 + i, j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = None

    def ref(self, cle, a, absolu=True):
        if a not in self.col:
            return None
        return f"{q('10_agregats')}!${self.col[a]}${self.row[cle]}" if absolu else f"{self.col[a]}{self.row[cle]}"


# ---------------------------------------------------------------------------
# Onglets de données brutes
# ---------------------------------------------------------------------------
def feuille_donnees(wb, nom, titre, explic, entetes, lignes, formats=None, py_cols=(), largeurs=None, python=None):
    """Onglet de données : bloc d'explication puis une ligne par enregistrement (style léger, sans bordure)."""
    o = Onglet(wb, nom, titre, None, explic, python=python, ncol=len(entetes))
    ws, h = o.ws, o.r
    for j, t in enumerate(entetes, 1):
        c = ws.cell(h, j, t)
        c.font, c.fill = Font(name=POLICE, size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = (largeurs[j - 1] if largeurs else 14)
    ws.row_dimensions[h].height = 30
    fpy = PatternFill("solid", fgColor=F_PY)
    for i, lig in enumerate(lignes):
        r = h + 1 + i
        for j, v in enumerate(lig, 1):
            if isinstance(v, str) and v.startswith("="):
                v = v.replace("{r}", str(r))
            if isinstance(v, float) and math.isnan(v):
                v = None
            c = ws.cell(r, j, v)
            if formats and formats[j - 1]:
                c.number_format = formats[j - 1]
            if (j - 1) in py_cols:
                c.fill = fpy
    ws.sheet_properties.tabColor = "808080"
    return h, h + 1, h + len(lignes)


def plage(nom, col, d, f):
    return f"{q(nom)}!${col}${d}:${col}${f}"


def donnees(wb, base, C, tables_dir):
    P = {}
    # D1 balances
    bal = pd.read_parquet(os.path.join(base, "compta_balances.parquet"))
    bal = bal[bal.nature == "comptes"].sort_values(["exercice", "compte"])
    lig = [[str(r.compte), str(r.intitule).strip(), int(r.exercice), r.periode_du.to_pydatetime(), r.periode_au.to_pydatetime(),
            float(r.ouv_d), float(r.ouv_c), float(r.mvt_d), float(r.mvt_c), float(r.sol_d), float(r.sol_c), "=J{r}-K{r}", r.fichier]
           for r in bal.itertuples()]
    h, d, f = feuille_donnees(wb, "D1_balances", "Balances des comptes du FPPN, une ligne par compte et par exercice", {
        "À quoi sert cet onglet": "C'est la source de tous les montants comptables du classeur : cotisations, prestations, frais, provisions, épargne, trésorerie, fonds propres. Les onglets de calcul additionnent ces lignes par numéro de compte (formules SUMIFS de l'onglet 10_agregats).",
        "D'où viennent les chiffres": "Balances exportées du logiciel comptable du FPPN (Sage 100), reçues le 15/10/2026 : « BALANCE FORMAT EXCEL DE 09-2020 AU 31-12-2020.xlsx », « BALANCE FPPN FORMAT EXCEL DU 01-01-2021 AU 31-12-2026.xlsx » (exercice 2021), « BALANCE FPPN 2022.xlsx » à « BALANCE FPPN 2025.xlsx », « BALANCE FPPN DU 01-01-2026 AU 31-08-2026.xlsx » (provisoire). La colonne « Fichier source » donne le fichier de chaque ligne. Lecture par le script Python « compta_fppn.py », sans modification des montants.",
        "Comment le calcul est fait": "Solde net = solde débit - solde crédit (colonne L, formule). Un solde net positif est un actif ou une charge ; négatif, un passif ou un produit. Pour l'exercice 2020, la balance commence au 01/09/2020 : les à-nouveaux contiennent les mouvements de janvier à août 2020, et le solde couvre donc toute l'année.",
        "Comment lire le résultat": "Filtrer la colonne « Compte » (par exemple « 7010* ») et la colonne « Exercice » pour retrouver à la main n'importe quel agrégat du classeur.",
        "Valable pour": "Exercices 2020 à 2025 et du 01/01 au 31/08/2026 (provisoire), comptes généraux (les balances des tiers ne sont pas reprises ici).",
        "Limites": "Montants tels que comptabilisés par le FPPN, non audités."},
        ["Compte", "Intitulé", "Exercice", "Du", "Au", "À-nouveaux débit", "À-nouveaux crédit", "Mouvements débit", "Mouvements crédit", "Solde débit", "Solde crédit", "Solde net (débit - crédit)", "Fichier source"],
        lig, [None, None, ENT, DATE, DATE] + [FCFA] * 7 + [None], largeurs=[12, 40, 10, 12, 12, 17, 17, 17, 17, 17, 17, 19, 52])
    P["bal"] = (d, f)
    wb["D1_balances"].freeze_panes = None
    # D2 fichiers de précomptes
    inv = pd.read_parquet(os.path.join(base, "sources_mois.parquet")).sort_values(["poste", "sens", "annee", "mois"])
    lig = [[r.src, r.dossier, r.fichier, r.feuille, str(r.poste), r.sens, int(r.annee), int(r.mois), int(r.lignes), float(r.montant),
            "Oui" if r.retenu else "Non", '=IF(AND(G{r}>=2020,OR(G{r}<2026,H{r}<=8)),"Oui","Non")'] for r in inv.itertuples()]
    h, d, f = feuille_donnees(wb, "D2_fichiers_precomptes", "Fichiers de précomptes reçus : une ligne par fichier et par mois", {
        "À quoi sert cet onglet": "Lister chaque fichier de précompte reçu (Solde, CGRAE, GNA), le poste, le sens (aller : demandé par le FPPN ; retour : précompté), le mois, le nombre de lignes et le montant. Sert aux onglets 1_couverture et 1_aller_retour.",
        "D'où viennent les chiffres": "Fichiers reçus le 15/10/2026 dans les dossiers « FPPN-SOLDE », « FPPN-CGRAE » et « FPPN-GNA » (colonnes Dossier et Fichier). Lecture et contrôle par le script Python " + SCRIPT_BASE + " : lignes de total et montants doublés écartés.",
        "Comment le calcul est fait": "« Retenu » = Non pour un fichier illisible, en double ou hors format (liste et motif dans l'onglet D6_rejets). « Dans la période » = Oui de janvier 2020 à août 2026 (formule).",
        "Comment lire le résultat": "Filtrer un poste et un sens pour voir les mois reçus et ceux qui manquent.",
        "Valable pour": "Fichiers reçus le 15/10/2026.", "Limites": "Montant et lignes du fichier entier : pas de détail agent par agent ici (voir D3_precomptes_mensuels)."},
        ["Identifiant", "Dossier", "Fichier", "Feuille", "Poste", "Sens", "Année", "Mois", "Lignes", "Montant", "Retenu", "Dans la période"],
        lig, [None, None, None, None, None, None, ENT, ENT, FCFA, FCFA, None, None], largeurs=[10, 36, 46, 22, 8, 8, 8, 7, 11, 16, 8, 10])
    P["inv"] = (d, f)
    # D3 précomptes mensuels (lignes individuelles agrégées par Python)
    pre = pd.read_parquet(os.path.join(base, "precomptes.parquet"), columns=["poste", "sens", "annee", "mois", "matricule", "mat6", "montant", "retenu"])
    pre = pre[pre.retenu]
    P["n_lignes"] = int(len(pre))
    g = pre.groupby(["poste", "sens", "annee", "mois"]).agg(lignes=("montant", "size"), montant=("montant", "sum"),
                                                            matricules=("matricule", "nunique"), mat6=("mat6", "nunique")).reset_index()
    lig = [[str(r.poste), r.sens, int(r.annee), int(r.mois), int(r.lignes), float(r.montant), int(r.matricules), int(r.mat6),
            '=IF(AND(C{r}>=2020,OR(C{r}<2026,D{r}<=8)),"Oui","Non")', '=IF(OR(C{r}<2026,D{r}<=8),"Oui","Non")'] for r in g.itertuples()]
    nl = f"{P['n_lignes']:,}".replace(",", " ")
    txt_py = (f"Les {nl} lignes individuelles retenues des fichiers de précomptes (une ligne par agent et par mois) dépassent la limite d'Excel "
              "(1 048 576 lignes par feuille). Le script Python « " + SCRIPT_ANA + " » (après lecture par « " + SCRIPT_BASE + " ») les additionne par poste, sens et mois ; "
              "les colonnes « Agents distincts » comptent les matricules différents du mois et ne peuvent pas être recalculées à partir des totaux. "
              "Méthode et résultats : rapport, partie 3 « Données et qualité » et partie 4 « Effectifs et cotisations ».").replace(",", " ")
    h, d, f = feuille_donnees(wb, "D3_precomptes_mensuels", "Précomptes par poste, sens et mois (totaux des lignes individuelles)", {
        "À quoi sert cet onglet": "Donner, pour chaque poste (936 maladie, 8Z6 décès, 998 PEP, 870 EFC, 9X4 Grenier, 9W3 GNA, 8Z7 MEDOC, CGRAE), chaque sens et chaque mois, le nombre de lignes, le montant précompté et le nombre d'agents distincts. Sert aux onglets 2, 4, 5 et 6.",
        "D'où viennent les chiffres": "Lignes retenues des fichiers de précomptes reçus le 15/10/2026 (onglet D2_fichiers_precomptes).",
        "Comment le calcul est fait": "Lignes = nombre de lignes du mois ; Montant = somme des montants du mois ; Agents distincts = nombre de matricules différents. « Dans la période » : janvier 2020 à août 2026 ; « Jusqu'au 31/08/2026 » : tous les mois jusqu'à août 2026 (utile pour le PEP, suivi depuis 2001).",
        "Comment lire le résultat": "Les montants s'additionnent d'un mois à l'autre ; les agents distincts ne s'additionnent pas (un agent présent douze mois compte une fois par mois).",
        "Valable pour": "Fichiers retenus après les contrôles de la partie 3 du rapport.", "Limites": "Les mois non reçus n'apparaissent pas (voir 1_couverture)."},
        ["Poste", "Sens", "Année", "Mois", "Lignes", "Montant", "Agents distincts (matricule)", "Agents distincts (mécano)", "Dans la période", "Jusqu'au 31/08/2026"],
        lig, [None, None, ENT, ENT, FCFA, FCFA, FCFA, FCFA, None, None], py_cols=(6, 7), largeurs=[8, 8, 8, 7, 11, 17, 14, 14, 10, 12], python=txt_py)
    P["pre"] = (d, f)
    # D3b agents distincts par année
    pa = pre[(pre.annee < 2026) | (pre.mois <= 8)].groupby(["poste", "sens", "annee"]).agg(matricules=("matricule", "nunique")).reset_index()
    lig = [[str(r.poste), r.sens, int(r.annee), int(r.matricules)] for r in pa.itertuples()]
    h, d, f = feuille_donnees(wb, "D3b_agents_par_annee", "Agents distincts par poste, sens et année (calcul Python)", {
        "À quoi sert cet onglet": "Donner le nombre d'agents différents précomptés dans l'année, par poste et par sens (épargnants du PEP et de l'EFC, adhérents des produits annexes).",
        "D'où viennent les chiffres": "Lignes individuelles retenues des fichiers de précomptes reçus le 15/10/2026, jusqu'au 31/08/2026.",
        "Comment le calcul est fait": "Nombre de matricules différents présents au moins un mois dans l'année.",
        "Comment lire le résultat": "Ce nombre n'est pas la somme des mois : un agent précompté douze mois compte une fois.",
        "Valable pour": "Années reçues, jusqu'au 31/08/2026.", "Limites": "Calcul impossible dans Excel (lignes individuelles trop nombreuses)."},
        ["Poste", "Sens", "Année", "Agents distincts dans l'année"], lig, [None, None, ENT, FCFA], py_cols=(3,), largeurs=[8, 8, 8, 18], python=txt_py)
    P["pa"] = (d, f)
    # D4 prêts du Grenier
    gb = pd.read_parquet(os.path.join(base, "grenier_balances.parquet"))
    gg = gb[gb.portefeuille == "global"].sort_values(["arrete", "compte"])
    lig = [[r.arrete.to_pydatetime(), str(r.compte), float(r.montant_debloque or 0), float(r.encours or 0), float(r.retard or 0),
            float(r.montant_echeance or 0), r.duree_mois, r.taux, r.fichier] for r in gg.itertuples()]
    h, d, f = feuille_donnees(wb, "D4_grenier_prets", "Prêts du Grenier : une ligne par prêt et par arrêté", {
        "À quoi sert cet onglet": "Source des encours, des retards et des échéances du Grenier (onglets 5_, 8_sensibilite_pertes_grenier et 12_grenier_tranches).",
        "D'où viennent les chiffres": "Balances âgées « BalanceAgeeDetaille … .xls » exportées du logiciel de prêts du Grenier, reçues le 15/10/2026 (portefeuille global ; colonne Fichier source). Noms des emprunteurs retirés.",
        "Comment le calcul est fait": "Retard = nombre de jours de retard de l'échéance impayée la plus ancienne, tel qu'exporté. Les onglets de calcul comptent les prêts et additionnent les encours par arrêté et par tranche de retard (COUNTIFS, SUMIFS).",
        "Comment lire le résultat": "Un encours est le capital restant dû à l'arrêté.",
        "Valable pour": "Arrêtés annuels du 31/12/2021 au 31/12/2025 et mensuels de 2026.", "Limites": "Exports non rapprochés des contrats de prêt."},
        ["Arrêté", "Compte du prêt", "Montant débloqué", "Encours", "Jours de retard", "Montant de l'échéance", "Durée (mois)", "Taux", "Fichier source"],
        lig, [DATE, None, FCFA, FCFA, ENT, FCFA, ENT, NB2, None], largeurs=[12, 14, 16, 16, 10, 16, 9, 8, 40])
    P["gre"] = (d, f)
    # D5 production du Grenier
    gp = pd.read_parquet(os.path.join(base, "grenier_production.parquet"))
    gp["annee"] = gp.periode_au.dt.year
    ret = (gp.portefeuille == "global") | ((gp.annee == 2026) & (gp.portefeuille != "global") & (gp.periode_au <= "2026-08-31"))
    lig = [[r.periode_du.to_pydatetime(), r.periode_au.to_pydatetime(), int(r.annee), r.portefeuille, r.type_credit, float(r.montant_debloque or 0),
            "Oui" if k else "Non"] for r, k in zip(gp.itertuples(), ret)]
    h, d, f = feuille_donnees(wb, "D5_grenier_production", "Prêts du Grenier débloqués : une ligne par prêt", {
        "À quoi sert cet onglet": "Source de la production annuelle de prêts par type de crédit (onglet 5_grenier_production).",
        "D'où viennent les chiffres": "Exports « production_de_la_periode … .xls » du logiciel de prêts du Grenier, reçus le 15/10/2026 : fichiers annuels de 2021 à 2025 (portefeuille global) et mensuels de 2026 (portefeuilles CRF, CRO, CRS).",
        "Comment le calcul est fait": "« Retenu » = Oui pour les fichiers annuels globaux et pour les fichiers mensuels 2026 par portefeuille jusqu'au 31/08/2026 ; les autres lignes feraient double compte.",
        "Comment lire le résultat": "La production est le montant prêté dans l'année, non l'encours.", "Valable pour": "2021 au 31/08/2026.", "Limites": "Aucune."},
        ["Période du", "Période au", "Année", "Portefeuille", "Type de crédit", "Montant débloqué", "Retenu"], lig,
        [DATE, DATE, ENT, None, None, FCFA, None], largeurs=[12, 12, 8, 12, 34, 16, 8])
    P["prod"] = (d, f)
    # D6 rejets
    rej = pd.read_parquet(os.path.join(base, "rejets.parquet"))
    lig = [[r.src, r.fichier, r.motif, re.sub(r" :.*", "", str(r.motif)), float(r.n or 0), (None if pd.isna(r.montant) else float(r.montant))] for r in rej.itertuples()]
    h, d, f = feuille_donnees(wb, "D6_rejets", "Lignes et fichiers écartés des calculs, avec leur motif", {
        "À quoi sert cet onglet": "Lister ce qui a été écarté et pourquoi (lignes de total, montants doublés, colonnes illisibles). Source de l'onglet 1_rejets.",
        "D'où viennent les chiffres": "Journal de contrôle produit par le script Python " + SCRIPT_BASE + " à la lecture des fichiers reçus le 15/10/2026.",
        "Comment le calcul est fait": "Motif court = début du motif, avant les deux-points.",
        "Comment lire le résultat": "Une ligne écartée n'est pas une perte : elle aurait faussé les totaux.", "Valable pour": "Fichiers reçus le 15/10/2026.", "Limites": "Aucune."},
        ["Identifiant", "Fichier", "Motif", "Motif court", "Lignes", "Montant"], lig, [None, None, None, None, FCFA, FCFA], largeurs=[10, 44, 60, 30, 10, 16])
    P["rej"] = (d, f)
    # D7 données externes
    infl = {int(k): v for k, v in C["inflation"]["valeurs"].items()}
    lig = [[a, infl[a], "Banque mondiale, Indicateurs du développement dans le monde, FP.CPI.TOTL.ZG, Côte d'Ivoire, mise à jour du " + str(C["inflation"]["mise_a_jour"])] for a in sorted(infl) if a >= 2020]
    h, d, f = feuille_donnees(wb, "D7_externe", "Données externes : inflation de la Côte d'Ivoire", {
        "À quoi sert cet onglet": "Inflation annuelle utilisée pour exprimer les cotisations et les prestations en francs constants (onglet 8_inflation_reel).",
        "D'où viennent les chiffres": "Banque mondiale, indicateur FP.CPI.TOTL.ZG (hausse annuelle des prix à la consommation, en %), téléchargé pour la mission.",
        "Comment le calcul est fait": "Valeur reprise telle quelle, en %.", "Comment lire le résultat": "4,2 veut dire que les prix ont augmenté de 4,2 % dans l'année.",
        "Valable pour": "2020 à 2025.", "Limites": "Inflation générale, non spécifique à la santé (voir l'onglet 13g)."},
        ["Année", "Inflation (%)", "Source"], lig, [ENT, NB2, None], largeurs=[8, 12, 90])
    P["infl"] = {a: f"{q('D7_externe')}!$B${d + i}" for i, a in enumerate(sorted(x for x in infl if x >= 2020))}
    # D8 tables de mortalité
    from python_calamine import CalamineWorkbook
    tabs = {}
    for nom, fic in (("CIMA H", "CIMA_H.xlsx"), ("CIMA F", "CIMA_F.xlsx"), ("TD CIMA", "TD_CIMA.xlsx"), ("TV CIMA", "TV_CIMA.xlsx")):
        w = CalamineWorkbook.from_path(os.path.join(tables_dir, fic))
        sh = w.get_sheet_by_name(w.sheet_names[0]).to_python()
        tabs[nom] = ({int(r[0]): float(r[1]) for r in sh[1:] if r[0] != ""}, w.sheet_names[0])
    lig = [[a] + [tabs[n][0].get(a, 0.0) for n in tabs] for a in range(0, 111)]
    h, d, f = feuille_donnees(wb, "D8_tables_mortalite", "Tables de mortalité CIMA : nombre de survivants lx par âge", {
        "À quoi sert cet onglet": "Source des espérances de vie, des rentes viagères et des probabilités de décès (onglet 8_tables_mortalite).",
        "D'où viennent les chiffres": "Fichiers « CIMA_H.xlsx », « CIMA_F.xlsx », « TD_CIMA.xlsx » et « TV_CIMA.xlsx » fournis le 18/10/2026 (colonne Lx).",
        "Comment le calcul est fait": "lx = nombre de personnes encore en vie à l'âge x, sur 1 000 000 de naissances.",
        "Comment lire le résultat": "lx(60) / lx(0) = probabilité d'atteindre 60 ans.", "Valable pour": "Population de référence des tables CIMA, non le FPPN.", "Limites": "Âges des mutualistes non reçus."},
        ["Âge"] + list(tabs), lig, [ENT] + ["#,##0.00"] * 4, largeurs=[8, 16, 16, 16, 16])
    P["mort"] = (d, f, list(tabs), {n: tabs[n][1] for n in tabs})
    return P


# ---------------------------------------------------------------------------
# Reconstruction des onglets 1 à 12
# ---------------------------------------------------------------------------
def py_txt(quoi, slides):
    return ("Les colonnes sur fond bleu clair (" + quoi + ") sont calculées par Python à partir des lignes individuelles des fichiers de précomptes "
            "ou des exports du Grenier, trop nombreuses pour Excel (plus de 1 048 576 lignes). Script : « " + SCRIPT_ANA + " », après lecture par « " + SCRIPT_BASE + " ». "
            "Méthode et résultats : rapport, " + slides + ".")


def reconstruire(wb, base, raw_dir, C, tables_dir, explic, lisible, index_rapport=None):
    raw = {}
    for f in ("FPPN_Calculs_rapport.xlsx", "FPPN_Calculs_complements.xlsx", "FPPN_Calculs_ratios.xlsx", "FPPN_Calculs_provisions.xlsx"):
        for k, v in pd.read_excel(os.path.join(raw_dir, f), sheet_name=None).items():
            raw[k] = v
    P = donnees(wb, base, C, tables_dir)
    ROW = {}

    def place(nom):
        if nom in wb.sheetnames:
            i = wb.sheetnames.index(nom)
            del wb[nom]
            return i
        return None

    def E(nom):
        return explic.get(nom, {})

    def ent(df):
        return [lisible(c) for c in df.columns]

    D2, D3, D3b, D4, D5, D6 = "D2_fichiers_precomptes", "D3_precomptes_mensuels", "D3b_agents_par_annee", "D4_grenier_prets", "D5_grenier_production", "D6_rejets"
    i2, j2 = P["inv"]
    i3, j3 = P["pre"]
    ib, jb = P["pa"]
    i4, j4 = P["gre"]
    i5, j5 = P["prod"]
    i6, j6 = P["rej"]

    def c2(col):
        return plage(D2, col, i2, j2)

    def c3(col):
        return plage(D3, col, i3, j3)

    def c4(col):
        return plage(D4, col, i4, j4)

    # ---------------- agrégats (onglet 10_agregats, recréé en premier car tout s'y réfère)
    pos = place("10_agregats")
    AGR = Agregats(wb, pos, P["bal"], E("10_agregats"))

    def A(cle, a):
        return AGR.ref(cle, a)

    df = raw["10_agregats"]
    mapk = {"cot_maladie": "cot_maladie_701", "cot_deces": "cot_deces_702", "prime_assurance_6518": "prime_6518", "provision_pensions_196": "provision_196",
            "actif_immobilise_net": "immob_nettes", "dettes_fournisseurs_cliniques": "fournisseurs_cliniques", "prets_pep": "prets_pep_ratios",
            "interets_dat_fcp": "produits_fin_77", "prest_maladie": "prest_maladie", "prest_deces": "prest_deces"}
    for _, r in df.iterrows():
        a = int(r["exercice"])
        for k in df.columns[1:]:
            cle = mapk.get(k, k)
            if cle in AGR.row and pd.notna(r[k]):
                CTRL.append(("10_agregats", f"{AGR.col[a]}{AGR.row[cle]}", float(r[k])))

    # ---------------- 1_couverture
    df = raw["1_couverture"]
    o = Onglet(wb, "1_couverture", "Couverture des fichiers de précomptes reçus, par poste et par sens", place("1_couverture"), E("1_couverture"))
    lig = [[str(r.poste), r.sens,
            f'=COUNTIFS({c2("E")},A{{r}},{c2("F")},B{{r}},{c2("K")},"Oui",{c2("L")},"Oui")',
            f'=SUMIFS({c2("I")},{c2("E")},A{{r}},{c2("F")},B{{r}},{c2("K")},"Oui",{c2("L")},"Oui")',
            f'=SUMIFS({c2("J")},{c2("E")},A{{r}},{c2("F")},B{{r}},{c2("K")},"Oui",{c2("L")},"Oui")'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [None, None, ENT, FCFA, FCFA], source="Source : onglet D2_fichiers_precomptes (fichiers retenus, janvier 2020 à août 2026)")
    controler("1_couverture", h, df)

    # ---------------- 1_aller_retour (avec onglet de détail mensuel)
    df = raw["1_aller_retour"]
    inv = pd.read_parquet(os.path.join(base, "sources_mois.parquet"))
    inv = inv[inv.retenu & (inv.annee >= 2020) & ((inv.annee < 2026) | (inv.mois <= 8))]
    pm = sorted({(str(p), int(a), int(m)) for p, a, m in zip(inv.poste, inv.annee, inv.mois) if str(p) in set(df.poste.astype(str))})
    om = Onglet(wb, "1_aller_retour_mois", "Rapprochement aller et retour, mois par mois", place("1_aller_retour") , {
        "À quoi sert cet onglet": "Comparer, pour chaque poste et chaque mois, le montant demandé par le FPPN (fichier aller) et le montant précompté (fichier retour). Sert à l'onglet 1_aller_retour.",
        "D'où viennent les chiffres": "Onglet D2_fichiers_precomptes (fichiers retenus, janvier 2020 à août 2026).",
        "Comment le calcul est fait": "Montant aller = somme des fichiers aller du mois ; montant retour = somme des fichiers retour du mois ; « Mois commun » = Oui si les deux fichiers existent ; taux = retour / aller.",
        "Exemple chiffré": "Un taux de 97,58 % veut dire que 97,58 FCFA ont été précomptés pour 100 FCFA demandés.",
        "Comment lire le résultat": "Un mois sans fichier aller ou sans fichier retour ne peut pas être rapproché.",
        "Valable pour": "Mois où les fichiers sont reçus.", "Limites": "Totaux du mois, pas de rapprochement agent par agent."})
    crit = '{c}("{s}")'
    lig = []
    for p, a, m in pm:
        lig.append([p, a, m,
                    f'=SUMIFS({c2("J")},{c2("E")},A{{r}},{c2("F")},"aller",{c2("G")},B{{r}},{c2("H")},C{{r}},{c2("K")},"Oui")',
                    f'=COUNTIFS({c2("E")},A{{r}},{c2("F")},"aller",{c2("G")},B{{r}},{c2("H")},C{{r}},{c2("K")},"Oui")',
                    f'=SUMIFS({c2("J")},{c2("E")},A{{r}},{c2("F")},"retour",{c2("G")},B{{r}},{c2("H")},C{{r}},{c2("K")},"Oui")',
                    f'=COUNTIFS({c2("E")},A{{r}},{c2("F")},"retour",{c2("G")},B{{r}},{c2("H")},C{{r}},{c2("K")},"Oui")',
                    '=IF(AND(E{r}>0,G{r}>0),"Oui","Non")', '=IF(H{r}="Oui",F{r}/D{r},"")'])
    hm, dm, fm = om.tableau(["Poste", "Année", "Mois", "Montant aller", "Fichiers aller", "Montant retour", "Fichiers retour", "Mois commun", "Taux retour / aller"],
                            lig, [None, ENT, ENT, FCFA, ENT, FCFA, ENT, None, PCT], source="Source : onglet D2_fichiers_precomptes")
    MO = "1_aller_retour_mois"

    def cm(col):
        return plage(MO, col, dm, fm)

    o = Onglet(wb, "1_aller_retour", "Rapprochement des fichiers aller et retour, par poste", wb.sheetnames.index(MO) + 1, E("1_aller_retour"))
    lig = [[str(r.poste), f'=COUNTIFS({cm("A")},A{{r}},{cm("H")},"Oui")', f'=SUMIFS({cm("D")},{cm("A")},A{{r}},{cm("H")},"Oui")',
            f'=SUMIFS({cm("F")},{cm("A")},A{{r}},{cm("H")},"Oui")', "=D{r}/C{r}",
            f'=_xlfn.MINIFS({cm("I")},{cm("A")},A{{r}},{cm("H")},"Oui")', f'=_xlfn.MAXIFS({cm("I")},{cm("A")},A{{r}},{cm("H")},"Oui")'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [None, ENT, FCFA, FCFA, PCT, PCT, PCT], source="Source : onglet 1_aller_retour_mois")
    controler("1_aller_retour", h, df)

    # ---------------- 1_rejets
    df = raw["1_rejets"]
    o = Onglet(wb, "1_rejets", "Lignes écartées des traitements, par motif", place("1_rejets"), E("1_rejets"))
    lig = [[r.motif_court, f'=COUNTIFS({plage(D6, "D", i6, j6)},A{{r}})', f'=SUMIFS({plage(D6, "E", i6, j6)},{plage(D6, "D", i6, j6)},A{{r}})'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [None, ENT, FCFA], source="Source : onglet D6_rejets")
    controler("1_rejets", h, df)

    # ---------------- 2_effectifs_mensuels
    df = raw["2_effectifs_mensuels"]
    o = Onglet(wb, "2_effectifs_mensuels", "Cotisants et retraités précomptés, par mois", place("2_effectifs_mensuels"), E("2_effectifs_mensuels"),
               python=py_txt("nombre de cotisants et de retraités : matricules différents du mois, repris de l'onglet D3_precomptes_mensuels", "partie 4 « Effectifs et cotisations », diapositive « Effectifs cotisants »"))
    lig = []
    ROW["2_eff"] = {}
    for i, r in enumerate(df.itertuples()):
        a, m = int(r.annee), int(r.mois)
        l_ = [a, m]
        l_.append(f'=SUMIFS({c3("G")},{c3("A")},"936",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")' if pd.notna(r.cotisants) else None)
        l_.append(f'=SUMIFS({c3("F")},{c3("A")},"936",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")' if pd.notna(r.montant) else None)
        l_.append(f'=SUMIFS({c3("G")},{c3("A")},"CGRAE",{c3("B")},"retour",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")' if pd.notna(r.retraites) else None)
        l_.append(f'=SUMIFS({c3("F")},{c3("A")},"CGRAE",{c3("B")},"retour",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")' if pd.notna(r.montant_cgrae) else None)
        l_.append(r.libelle)
        lig.append(l_)
    h, d, f = o.tableau(ent(df), lig, [ENT, ENT, FCFA, FCFA, FCFA, FCFA, None], py_cols=(2, 4), source="Source : onglet D3_precomptes_mensuels (postes 936 et CGRAE)")
    for i, r in enumerate(df.itertuples()):
        ROW["2_eff"][(int(r.annee), int(r.mois))] = d + i
    controler("2_effectifs_mensuels", h, df)

    # ---------------- onglets entièrement calculés par Python (lignes individuelles)
    PYS = {"2_entrees_sorties": ("Entrées et sorties mensuelles de cotisants", "entrées, sorties : comparaison des matricules d'un mois à l'autre", "partie 4, diapositive « Entrées et sorties »"),
           "2_taux_936": ("Taux du précompte maladie (poste 936) rapporté au salaire de base", "nombre de lignes par tranche de taux, calculé ligne par ligne sur le salaire de base", "partie 4, diapositives « Calcul pas à pas : taux de cotisation » et « Taux de cotisation »"),
           "2_rapport_8Z6_936": ("Rapport des précomptes décès (8Z6) et maladie (936)", "rapport médian 8Z6 / 936 par agent et par mois", "partie 3, diapositive « Contrôle croisé : effectifs du FPPN et fichiers de la Solde »"),
           "2_anciennete": ("Ancienneté des agents cotisants", "nombre d'agents par tranche d'ancienneté, calculé agent par agent", "partie 4, diapositive « Structure des effectifs »"),
           "4_poste_998_references": ("Références du poste 998 (Plan Épargne Plus)", "lignes, montants et mutualistes par référence de produit", "partie 6, diapositive « Poste 998 : autres références »")}
    for nom, (titre, quoi, sl) in PYS.items():
        df = raw[nom]
        o = Onglet(wb, nom, titre, place(nom), E(nom), python=py_txt(quoi, sl))
        fm_ = [None if df[c].dtype == object else (PCT if "taux" in c.lower() or c == "r" else FCFA) for c in df.columns]
        lig = [[(None if (isinstance(v, float) and math.isnan(v)) else v) for v in row] for row in df.itertuples(index=False)]
        cols_py = tuple(j for j, c in enumerate(df.columns) if df[c].dtype != object)
        if nom == "2_entrees_sorties":
            lig = [l_[:5] + [f'=SUMIFS({c3("G")},{c3("A")},"936",{c3("C")},B{{r}},{c3("D")},C{{r}},{c3("I")},"Oui")'] for l_ in lig]
            cols_py = (3, 4)
            fm_ = [None, ENT, ENT, FCFA, FCFA, FCFA]
        if nom == "2_rapport_8Z6_936":
            fm_ = [ENT, "0.0000"]
        h, d, f = o.tableau(ent(df), lig, fm_, py_cols=cols_py, source="Source : calcul Python « " + SCRIPT_ANA + " » sur les lignes individuelles des fichiers de précomptes reçus le 15/10/2026")
        if nom == "2_entrees_sorties":
            controler(nom, h, df, cols=[5])

    # ---------------- 2_cotisations
    df = raw["2_cotisations"]
    o = Onglet(wb, "2_cotisations", "Cotisations comptabilisées et précomptées, par exercice", place("2_cotisations"), E("2_cotisations"))
    lig = []
    for r in df.itertuples():
        a = int(r.exercice)
        l_ = [a, "=" + A("cot_maladie_7010", a), "=" + A("cot_deces_7021", a)]
        for po, vm, vn in (("936", r.prec_936, r.prec_936_mois), ("8Z6", r.prec_8Z6, r.prec_8Z6_mois), ("CGRAE", r.prec_cgrae, r.prec_cgrae_mois)):
            l_.append(f'=SUMIFS({c3("F")},{c3("A")},"{po}",{c3("B")},"retour",{c3("C")},A{{r}},{c3("I")},"Oui")' if pd.notna(vm) else None)
            l_.append(f'=COUNTIFS({c3("A")},"{po}",{c3("B")},"retour",{c3("C")},A{{r}},{c3("I")},"Oui")' if pd.notna(vn) else None)
        lig.append(l_)
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, ENT, FCFA, ENT, FCFA, ENT], source="Source : onglets 10_agregats (comptes 7010 et 7021) et D3_precomptes_mensuels (fichiers retour)")
    controler("2_cotisations", h, df)

    # ---------------- 3_sinistralite_resultat
    df = raw["3_sinistralite_resultat"]
    o = Onglet(wb, "3_sinistralite_resultat", "Sinistralité et résultat, par exercice", place("3_sinistralite_resultat"), E("3_sinistralite_resultat"))
    lig = []
    for a in EX:
        lig.append([a, "=" + A("cot_maladie_7010", a), "=" + A("prest_maladie", a), "=C{r}/B{r}", "=" + A("cot_deces_7021", a), "=" + A("prest_deces", a),
                    "=" + A("prime_6518", a), "=F{r}/E{r}", "=(F{r}+G{r})/E{r}", "=" + A("frais_gestion", a), "=J{r}/(B{r}+E{r})",
                    "=" + A("produits_7", a), "=" + A("charges_6", a), "=" + A("resultat_hao", a), "=L{r}-M{r}+N{r}"])
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, PCT, FCFA, FCFA, FCFA, PCT, PCT, FCFA, PCT, FCFA, FCFA, FCFA, FCFA], source="Source : onglet 10_agregats")
    ROW["3"] = {a: d + i for i, a in enumerate(EX)}
    controler("3_sinistralite_resultat", h, df)
    S3 = q("3_sinistralite_resultat")

    # ---------------- 3_prestations_maladie_nature
    df = raw["3_prestations_maladie_nature"]
    o = Onglet(wb, "3_prestations_maladie_nature", "Prestations maladie par nature (FCFA)", place("3_prestations_maladie_nature"), E("3_prestations_maladie_nature"))
    cles = ["nat_pharmacie", "nat_consultations", "nat_hospitalisations", "nat_analyses", "nat_radio", "nat_autres_6012", "nat_autres"]
    lig = [[r.nature] + ["=" + A(k, a) for a in EX] for r, k in zip(df.itertuples(), cles)]
    h, d, f = o.tableau(ent(df), lig, [None] + [FCFA] * 7, source="Source : onglet 10_agregats")
    controler("3_prestations_maladie_nature", h, df)

    # ---------------- 4_bilan
    df = raw["4_bilan"]
    o = Onglet(wb, "4_bilan", "Postes du bilan utilisés par le rapport (FCFA)", place("4_bilan"), E("4_bilan"))
    cles = ["prov_maladie", "prov_obseques", "epargne_pep", "epargne_efc", "banque_efc", "banque_pep", "fcp", "tresorerie_totale", "creance_hopital", "cmp", "grenier",
            "cot_a_recevoir", "terrains", "immob_nettes", "fournisseurs_cliniques", "reserves_10_11", "ran_120_121", "pertes_129", "prets_pep_bilan", "total_actif"]
    lig = []
    for a in EX:
        l_ = [a] + ["=" + A(k, a) for k in cles]
        l_ += ["=D{r}+E{r}", "=F{r}+G{r}+H{r}", "=W{r}/V{r}", "=B{r}+C{r}", "=Q{r}+R{r}+S{r}", f"={S3}!$O${ROW['3'][a]}", "=Z{r}+AA{r}"]
        lig.append(l_)
    h, d, f = o.tableau(ent(df), lig, [ENT] + [FCFA] * 22 + [PCT] + [FCFA] * 4, source="Source : onglets 10_agregats et 3_sinistralite_resultat")
    ROW["4"] = {a: d + i for i, a in enumerate(EX)}
    controler("4_bilan", h, df)
    S4 = q("4_bilan")

    def B4(col, a):
        return f"{S4}!${col}${ROW['4'][a]}"

    # ---------------- 4_pep_versements
    df = raw["4_pep_versements"]
    o = Onglet(wb, "4_pep_versements", "Versements au Plan Épargne Plus", place("4_pep_versements"), E("4_pep_versements"),
               python=py_txt("épargnants : agents différents précomptés dans l'année, repris de l'onglet D3b_agents_par_annee", "partie 6, diapositive « Versements d'épargne »"))
    lig = []
    for i, r in enumerate(df.itertuples()):
        a = int(r.annee)
        lig.append([a, f'=SUMIFS({c3("F")},{c3("A")},"998",{c3("B")},"retour",{c3("C")},A{{r}},{c3("J")},"Oui")',
                    f'=SUMIFS({plage(D3b, "D", ib, jb)},{plage(D3b, "A", ib, jb)},"998",{plage(D3b, "B", ib, jb)},"retour",{plage(D3b, "C", ib, jb)},A{{r}})',
                    "=B{r}" if i == 0 else "=D" + "{p}+B{r}", ("=" + B4("D", a)) if a in ROW["4"] else None])
    for i, l_ in enumerate(lig):
        if i:
            l_[3] = l_[3].replace("{p}", str(o.r + i))
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, FCFA], py_cols=(2,), source="Source : onglets D3_precomptes_mensuels, D3b_agents_par_annee et 4_bilan")
    controler("4_pep_versements", h, df)

    # ---------------- 4_efc_versements
    df = raw["4_efc_versements"]
    o = Onglet(wb, "4_efc_versements", "Versements à l'Entraide de Fin de Carrière", place("4_efc_versements"), E("4_efc_versements"),
               python=py_txt("épargnants : agents différents précomptés dans l'année, repris de l'onglet D3b_agents_par_annee", "partie 6, diapositive « Versements d'épargne »"))
    lig = []
    for i, r in enumerate(df.itertuples()):
        a = int(r.annee)
        lig.append([a, f'=SUMIFS({c3("F")},{c3("A")},"870",{c3("B")},"retour",{c3("C")},A{{r}},{c3("I")},"Oui")',
                    f'=SUMIFS({plage(D3b, "D", ib, jb)},{plage(D3b, "A", ib, jb)},"870",{plage(D3b, "B", ib, jb)},"retour",{plage(D3b, "C", ib, jb)},A{{r}})',
                    f'=COUNTIFS({c3("A")},"870",{c3("B")},"retour",{c3("C")},A{{r}},{c3("I")},"Oui")', "=" + B4("E", a), None if i == 0 else "=E{r}-E" + str(o.r + i)])
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, ENT, FCFA, FCFA], py_cols=(2,), source="Source : onglets D3_precomptes_mensuels, D3b_agents_par_annee et 4_bilan")
    controler("4_efc_versements", h, df)

    # ---------------- 5_grenier_encours
    df = raw["5_grenier_encours"]
    o = Onglet(wb, "5_grenier_encours", "Encours des prêts du Grenier", place("5_grenier_encours"), E("5_grenier_encours"))
    lig = []
    ROW["5enc"] = {}
    for r in df.itertuples():
        dt_ = pd.Timestamp(r.arrete)
        fin = None
        if pd.notna(r.financement_fppn_47111102):
            fin = "=" + A("grenier", dt_.year)
        lig.append([dt_.to_pydatetime(), f"=COUNTIFS({c4('A')},A{{r}})", f"=SUMIFS({c4('D')},{c4('A')},A{{r}})", f"=SUMIFS({c4('C')},{c4('A')},A{{r}})", fin])
    h, d, f = o.tableau(ent(df), lig, [DATE, ENT, FCFA, FCFA, FCFA], source="Source : onglets D4_grenier_prets et 10_agregats (compte 47111102)")
    for i, r in enumerate(df.itertuples()):
        ROW["5enc"][pd.Timestamp(r.arrete).strftime("%d/%m/%Y")] = d + i
    controler("5_grenier_encours", h, df, cols=[1, 2, 3, 4])

    TR = [(-1, 0), (0, 30), (30, 90), (90, 180), (180, 365), (365, 10000)]

    def tranche_f(date_cell, lo, hi, quoi):
        if quoi == "n":
            return f'=COUNTIFS({c4("A")},{date_cell},{c4("E")},">{lo}",{c4("E")},"<={hi}")'
        return f'=SUMIFS({c4("D")},{c4("A")},{date_cell},{c4("E")},">{lo}",{c4("E")},"<={hi}")'

    # ---------------- 5_grenier_retards (31/08/2026)
    df = raw["5_grenier_retards"]
    o = Onglet(wb, "5_grenier_retards", "Encours du Grenier par tranche de retard au 31/08/2026", place("5_grenier_retards"), E("5_grenier_retards"))
    lig = [[r.tranche, tranche_f("DATE(2026,8,31)", lo, hi, "n"), tranche_f("DATE(2026,8,31)", lo, hi, "e"), "=C{r}/SUM(C$" + str(o.r + 1) + ":C$" + str(o.r + 6) + ")"]
           for r, (lo, hi) in zip(df.itertuples(), TR)]
    h, d, f = o.tableau(ent(df), lig, [None, ENT, FCFA, PCT], source="Source : onglet D4_grenier_prets (arrêté du 31/08/2026)")
    controler("5_grenier_retards", h, df)

    # ---------------- 5_grenier_par_historique
    df = raw["5_grenier_par_historique"]
    o = Onglet(wb, "5_grenier_par_historique", "Portefeuille en retard du Grenier (plus de 30 et de 90 jours)", place("5_grenier_par_historique"), E("5_grenier_par_historique"))
    lig = [[pd.to_datetime(r.arrete, dayfirst=True).to_pydatetime(),
            f'=SUMIFS({c4("D")},{c4("A")},A{{r}},{c4("E")},">30")/SUMIFS({c4("D")},{c4("A")},A{{r}})',
            f'=SUMIFS({c4("D")},{c4("A")},A{{r}},{c4("E")},">90")/SUMIFS({c4("D")},{c4("A")},A{{r}})'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [DATE, PCT, PCT], source="Source : onglet D4_grenier_prets")
    controler("5_grenier_par_historique", h, df, cols=[1, 2])

    # ---------------- 5_grenier_production
    df = raw["5_grenier_production"]
    o = Onglet(wb, "5_grenier_production", "Production du Grenier par type de prêt (FCFA)", place("5_grenier_production"), E("5_grenier_production"))
    hdr = o.r
    lig = [[int(r[0])] + [f'=SUMIFS({plage(D5, "F", i5, j5)},{plage(D5, "C", i5, j5)},$A{{r}},{plage(D5, "E", i5, j5)},{L(j + 2)}${hdr},{plage(D5, "G", i5, j5)},"Oui")'
                          for j in range(len(df.columns) - 1)] for r in df.itertuples(index=False)]
    h, d, f = o.tableau([df.columns[0].capitalize()] + list(df.columns[1:]), lig, [ENT] + [FCFA] * (len(df.columns) - 1),
                        total=["Total"] + [f"=SUM({L(j + 2)}{{d}}:{L(j + 2)}{{f}})" for j in range(len(df.columns) - 1)], source="Source : onglet D5_grenier_production (lignes retenues)")
    ROW["5prod"] = {int(r[0]): d + i for i, r in enumerate(df.itertuples(index=False))}
    ROW["5prod_ncol"] = len(df.columns)
    controler("5_grenier_production", h, df)

    # ---------------- 5_grenier_echeances_precomptes
    df = raw["5_grenier_echeances_precomptes"]
    o = Onglet(wb, "5_grenier_echeances_precomptes", "Échéances du Grenier et précomptes du poste 9X4", place("5_grenier_echeances_precomptes"), E("5_grenier_echeances_precomptes"),
               python=py_txt("emprunteurs de la balance âgée et emprunteurs précomptés : nombres de mécanos différents", "partie 7, diapositive « Échéances et précomptes »"))
    gb = pd.read_parquet(os.path.join(base, "grenier_balances.parquet"), columns=["arrete", "portefeuille"])
    ar26 = sorted({d_ for d_ in gb[gb.portefeuille == "global"].arrete if d_.year == 2026})
    lig = []
    for r, d_ in zip(df.itertuples(), ar26):
        dc = f"DATE({d_.year},{d_.month},{d_.day})"
        lig.append([r.mois, f'=SUMIFS({c4("F")},{c4("A")},{dc},{c4("D")},">0")',
                    (f'=SUMIFS({c3("F")},{c3("A")},"9X4",{c3("B")},"retour",{c3("C")},{d_.year},{c3("D")},{d_.month})' if pd.notna(r.precompte_9x4) else None),
                    r.emprunteurs_balance, (f'=SUMIFS({c3("H")},{c3("A")},"9X4",{c3("B")},"retour",{c3("C")},{d_.year},{c3("D")},{d_.month})' if pd.notna(r.precompte_9x4) else r.precomptes_9x4)])
    h, d, f = o.tableau(ent(df), lig, [None, FCFA, FCFA, FCFA, FCFA], py_cols=(3, 4), source="Source : onglets D4_grenier_prets et D3_precomptes_mensuels (poste 9X4)")
    controler("5_grenier_echeances_precomptes", h, df, cols=[1, 2, 4])

    # ---------------- 6_cgrae
    df = raw["6_cgrae"]
    o = Onglet(wb, "6_cgrae", "Retraités précomptés par la CGRAE, par mois", place("6_cgrae"), E("6_cgrae"),
               python=py_txt("retraités : matricules différents du mois, repris de l'onglet D3_precomptes_mensuels", "partie 8, diapositive « Retraités CGRAE »"))
    lig = [[int(r.annee), int(r.mois), f'=SUMIFS({c3("G")},{c3("A")},"CGRAE",{c3("B")},"retour",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")',
            f'=SUMIFS({c3("F")},{c3("A")},"CGRAE",{c3("B")},"retour",{c3("C")},A{{r}},{c3("D")},B{{r}},{c3("I")},"Oui")', "=D{r}/C{r}", r.libelle] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [ENT, ENT, FCFA, FCFA, FCFA, None], py_cols=(2,), source="Source : onglet D3_precomptes_mensuels (CGRAE, fichiers retour)")
    controler("6_cgrae", h, df)

    # ---------------- 6_produits_annexes
    df = raw["6_produits_annexes"]
    o = Onglet(wb, "6_produits_annexes", "Produits annexes : flotte automobile (GNA) et MEDOC", place("6_produits_annexes"), E("6_produits_annexes"),
               python=py_txt("adhérents : agents différents précomptés dans l'année, repris de l'onglet D3b_agents_par_annee", "partie 8, diapositive « Produits annexes »"))
    lig = [[str(r.poste), int(r.annee), f'=SUMIFS({plage(D3b, "D", ib, jb)},{plage(D3b, "A", ib, jb)},A{{r}},{plage(D3b, "B", ib, jb)},"retour",{plage(D3b, "C", ib, jb)},B{{r}})',
            f'=SUMIFS({c3("F")},{c3("A")},A{{r}},{c3("B")},"retour",{c3("C")},B{{r}},{c3("I")},"Oui")',
            f'=COUNTIFS({c3("A")},A{{r}},{c3("B")},"retour",{c3("C")},B{{r}},{c3("I")},"Oui")'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [None, ENT, FCFA, FCFA, ENT], py_cols=(2,), source="Source : onglets D3_precomptes_mensuels et D3b_agents_par_annee")
    controler("6_produits_annexes", h, df)

    # ---------------- 7_rendement_placements
    df = raw["7_rendement_placements"]
    o = Onglet(wb, "7_rendement_placements", "Rendement apparent des placements", place("7_rendement_placements"), E("7_rendement_placements"))
    lig = [[a, "=" + A("produits_fin_77", a), f"=({A('placements_52_53', a)}+{A('placements_52_53', a - 1)})/2", "=B{r}/C{r}"] for a in range(2021, 2026)]
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, PCT], source="Source : onglet 10_agregats (comptes 77, 522 à 526, 533)")
    ROW["7"] = {a: d + i for i, a in enumerate(range(2021, 2026))}
    controler("7_rendement_placements", h, df)

    # ---------------- 8_inflation_reel
    df = raw["8_inflation_reel"]
    o = Onglet(wb, "8_inflation_reel", "Cotisations et prestations en valeur réelle (base 2020)", place("8_inflation_reel"), E("8_inflation_reel"))
    lig = []
    for i, a in enumerate(range(2020, 2026)):
        r3 = ROW["3"][a]
        lig.append([a, f"={S3}!$B${r3}+{S3}!$E${r3}", "=B{r}*100/F{r}", f"={S3}!$C${r3}", "=D{r}*100/F{r}",
                    "=100" if i == 0 else "=F" + str(o.r + i) + "*(1+G{r}/100)", "=" + P["infl"][a]])
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, FCFA, NB2, NB2], source="Source : onglets 3_sinistralite_resultat et D7_externe (Banque mondiale)")
    controler("8_inflation_reel", h, df)

    # ---------------- 8_tcam
    df = raw["8_tcam"]
    o = Onglet(wb, "8_tcam", "Taux de croissance annuel moyen des principaux indicateurs", place("8_tcam"), E("8_tcam"))
    S2, S5e, S5p = q("2_effectifs_mensuels"), q("5_grenier_encours"), q("5_grenier_production")
    tot_col = L(ROW["5prod_ncol"] + 1) if False else None
    refs = [(f"={S3}!$B${ROW['3'][2020]}+{S3}!$E${ROW['3'][2020]}", f"={S3}!$B${ROW['3'][2025]}+{S3}!$E${ROW['3'][2025]}"),
            (f"={S3}!$C${ROW['3'][2021]}", f"={S3}!$C${ROW['3'][2025]}"),
            ("=" + B4("V", 2020), "=" + B4("V", 2025)),
            (f"={S5e}!$C${ROW['5enc']['31/12/2022']}", f"={S5e}!$C${ROW['5enc']['31/12/2025']}"),
            (f"=SUM({S5p}!$B${ROW['5prod'][2022]}:${L(ROW['5prod_ncol'])}${ROW['5prod'][2022]})", f"=SUM({S5p}!$B${ROW['5prod'][2025]}:${L(ROW['5prod_ncol'])}${ROW['5prod'][2025]})"),
            (f"={S2}!$C${ROW['2_eff'][(2024, 1)]}", f"={S2}!$C${ROW['2_eff'][(2026, 8)]}"),
            ("=" + B4("K", 2020), "=" + B4("K", 2025))]
    lig = [[r.indicateur, str(r.debut), str(r.fin), v0, v1, ("=31/12" if abs(r.annees - 31 / 12) < 1e-9 else float(r.annees)), "=(E{r}/D{r})^(1/F{r})-1", "=E{r}/D{r}"]
           for r, (v0, v1) in zip(df.itertuples(), refs)]
    h, d, f = o.tableau(ent(df), lig, [None, None, None, FCFA, FCFA, NB2, PCT, NB2], source="Source : onglets 3_sinistralite_resultat, 4_bilan, 5_grenier_encours, 5_grenier_production, 2_effectifs_mensuels")
    controler("8_tcam", h, df)

    # ---------------- 8_tendances (séries dans un onglet de détail)
    df = raw["8_tendances"]
    os_ = Onglet(wb, "8_tendances_series", "Séries utilisées pour les tendances", place("8_tendances"), {
        "À quoi sert cet onglet": "Ranger, colonne par colonne, les valeurs dont l'onglet 8_tendances mesure la tendance (x = rang du mois ou de l'année, y = valeur).",
        "D'où viennent les chiffres": "Onglets 2_effectifs_mensuels, 3_sinistralite_resultat et D3_precomptes_mensuels (formules).",
        "Comment le calcul est fait": "Mois : x = (année - année de départ) × 12 + mois - 1. Prestations maladie : y = logarithme népérien, pour mesurer une croissance en %.",
        "Comment lire le résultat": "Chaque paire de colonnes (x, y) est une série.", "Valable pour": "Périodes indiquées dans l'onglet 8_tendances.", "Limites": "Aucune."})
    pre = pd.read_parquet(os.path.join(base, "precomptes.parquet"), columns=["poste", "sens", "annee", "mois", "retenu"])
    pre = pre[pre.retenu & (pre.annee >= 2020) & ((pre.annee < 2026) | (pre.mois <= 8))]
    k9 = sorted({(int(a), int(m)) for a, m in zip(pre[(pre.poste == "9X4") & (pre.sens == "retour")].annee, pre[(pre.poste == "9X4") & (pre.sens == "retour")].mois)})
    k8 = sorted({(int(a), int(m)) for a, m in zip(pre[(pre.poste == "998") & (pre.sens == "retour")].annee, pre[(pre.poste == "998") & (pre.sens == "retour")].mois)})
    keff = [k for k in sorted(ROW["2_eff"]) if pd.notna(raw["2_effectifs_mensuels"].set_index(["annee", "mois"]).loc[k, "cotisants"])]
    series = [
        [((a - 2024) * 12 + m - 1, f"={S2}!$C${ROW['2_eff'][(a, m)]}") for a, m in keff],
        [(a, f"={S3}!$D${ROW['3'][a]}*100") for a in range(2020, 2026)],
        [(a, f"=LN({S3}!$C${ROW['3'][a]})") for a in range(2021, 2026)],
        [((a - 2023) * 12 + m - 1, f'=SUMIFS({c3("F")},{c3("A")},"9X4",{c3("B")},"retour",{c3("C")},{a},{c3("D")},{m},{c3("I")},"Oui")') for a, m in k9],
        [((a - 2020) * 12 + m - 1, f'=SUMIFS({c3("F")},{c3("A")},"998",{c3("B")},"retour",{c3("C")},{a},{c3("D")},{m},{c3("I")},"Oui")') for a, m in k8]]
    nmax = max(len(s_) for s_ in series)
    lig = [[(s_[i][k] if i < len(s_) else None) for s_ in series for k in (0, 1)] for i in range(nmax)]
    entS = []
    for j in range(1, 6):
        entS += [f"x série {j}", f"y série {j}"]
    hs, ds, fs = os_.tableau(entS, lig, [NB2] * 10, source="Source : onglets 2_effectifs_mensuels, 3_sinistralite_resultat, D3_precomptes_mensuels")
    o = Onglet(wb, "8_tendances", "Tendances statistiques (régressions linéaires)", wb.sheetnames.index("8_tendances_series") + 1, E("8_tendances"))
    SS = q("8_tendances_series")
    lig = []
    for j, (r, s_) in enumerate(zip(df.itertuples(), series)):
        xr = f"{SS}!${L(2 * j + 1)}${ds}:${L(2 * j + 1)}${ds + len(s_) - 1}"
        yr = f"{SS}!${L(2 * j + 2)}${ds}:${L(2 * j + 2)}${ds + len(s_) - 1}"
        se = f"(STEYX({yr},{xr})/SQRT(DEVSQ({xr})))"
        l_ = [r.serie, f"=COUNT({yr})", f"=SLOPE({yr},{xr})", f"=C{{r}}-_xlfn.T.INV.2T(0.05,B{{r}}-2)*{se}", f"=C{{r}}+_xlfn.T.INV.2T(0.05,B{{r}}-2)*{se}",
              f"=_xlfn.T.DIST.2T(ABS(C{{r}}/{se}),B{{r}}-2)", f"=RSQ({yr},{xr})"]
        l_ += (["=EXP(C{r})-1", "=EXP(D{r})-1", "=EXP(E{r})-1"] if pd.notna(r.croissance_annuelle) else [None, None, None])
        lig.append(l_)
    h, d, f = o.tableau(ent(df), lig, [None, ENT, "0.0000", "0.0000", "0.0000", "0.0000", "0.0000", PCT, PCT, PCT], source="Source : onglet 8_tendances_series")
    controler("8_tendances", h, df)

    # ---------------- 8_sensibilite_retraits
    df = raw["8_sensibilite_retraits"]
    o = Onglet(wb, "8_sensibilite_retraits", "Sensibilité de la trésorerie à des retraits d'épargne", place("8_sensibilite_retraits"), E("8_sensibilite_retraits"))
    lig = [[float(r.taux_retrait), "=" + B4("V", 2026) + "*A{r}", "=" + B4("W", 2026), "=" + B4("I", 2026), "=C{r}-B{r}", "=D{r}-B{r}"] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [PCT, FCFA, FCFA, FCFA, FCFA, FCFA], orange_cols=(0,), source="Source : onglet 4_bilan (31/08/2026) ; cellules orange modifiables")
    controler("8_sensibilite_retraits", h, df)

    # ---------------- 8_sensibilite_pertes_grenier
    df = raw["8_sensibilite_pertes_grenier"]
    o = Onglet(wb, "8_sensibilite_pertes_grenier", "Sensibilité aux pertes sur les prêts du Grenier", place("8_sensibilite_pertes_grenier"), E("8_sensibilite_pertes_grenier"))
    lig = [[float(r.perte), f'=SUMIFS({c4("D")},{c4("A")},DATE(2026,8,31),{c4("E")},">180")*A{{r}}', f'=SUMIFS({c4("D")},{c4("A")},DATE(2026,8,31),{c4("E")},">90")*A{{r}}'] for r in df.itertuples()]
    h, d, f = o.tableau(ent(df), lig, [PCT, FCFA, FCFA], orange_cols=(0,), source="Source : onglet D4_grenier_prets (31/08/2026) ; cellules orange modifiables")
    controler("8_sensibilite_pertes_grenier", h, df)

    # ---------------- 8_provisions_mois
    df = raw["8_provisions_mois"]
    o = Onglet(wb, "8_provisions_mois", "Provisions techniques en mois de prestations", place("8_provisions_mois"), E("8_provisions_mois"))
    lig = [[a, "=" + B4("Y", a), f"=({S3}!$C${ROW['3'][a]}+{S3}!$F${ROW['3'][a]})/{8 if a == 2026 else 12}", "=B{r}/C{r}"] for a in EX]
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, NB2], source="Source : onglets 4_bilan et 3_sinistralite_resultat")
    controler("8_provisions_mois", h, df)

    # ---------------- 8_tables_mortalite
    df = raw["8_tables_mortalite"]
    md, mf, noms, feuilles = P["mort"]
    o = Onglet(wb, "8_tables_mortalite", "Indicateurs des tables de mortalité CIMA", place("8_tables_mortalite"), E("8_tables_mortalite"))
    DM = q("D8_tables_mortalite")
    age = f"{DM}!$A${md}:$A${mf}"
    lig = []
    for k, nom in enumerate(noms):
        col = L(2 + k)
        lx = f"{DM}!${col}${md}:${col}${mf}"

        def LX(x):
            return f"INDEX({lx},{x + 1})"

        def e_(x):
            return f'=SUMIFS({lx},{age},">{x}")/{LX(x)}+0.5'

        def a_(x):
            return f"=SUMPRODUCT(({age}>={x})*{lx}*(1/1.035)^({age}-{x}))/{LX(x)}"
        lig.append([nom, feuilles[nom], e_(0), e_(55), e_(60), a_(55), a_(60), a_(65), f"=1-{LX(41)}/{LX(40)}", f"=1-{LX(56)}/{LX(55)}",
                    f"=({LX(20)}-{LX(60)})/{LX(20)}", f"={LX(60)}/{LX(0)}"])
    h, d, f = o.tableau(ent(df), lig, [None, None] + [NB2] * 6 + ["0.0000"] * 4, source="Source : onglet D8_tables_mortalite")
    controler("8_tables_mortalite", h, df)
    return P, ROW, AGR, raw


# ---------------------------------------------------------------------------
# Onglets 10 et 12
# ---------------------------------------------------------------------------
def test_ok(test, v):
    t = test[0]
    if t == "entre":
        return f"AND({v}>={test[1]},{v}<={test[2]})"
    if t == "min":
        return f"{v}>={test[1]}"
    if t == "min_strict":
        return f"{v}>{test[1]}"
    return f"{v}<={test[1]}"


def statut_f(test, v, a):
    """Formule du statut d'un ratio (règle écrite de ratios_airms_fppn.py)."""
    ind = '" (indicatif)"' if (a < 2022 or a == 2026) else '""'
    if test is None:
        return '="À apprécier"'
    t = test[0]
    if t == "entre":
        core = f'IF({test_ok(test, v)},"Conforme",IF({v}>{test[2]},"Au-dessus de la fourchette","Non conforme"))'
    elif t == "max":
        core = f'IF({v}<0,"Non conforme (ressources négatives)",IF({test_ok(test, v)},"Conforme","Non conforme"))'
    else:
        core = f'IF({test_ok(test, v)},"Conforme","Non conforme")'
    return f'=IF({v}="","En attente",{core}&{ind})'


def reconstruire_10_12(wb, base, raw, P, ROW, AGR, Q, explic, lisible):
    def place(nom):
        if nom in wb.sheetnames:
            i = wb.sheetnames.index(nom)
            del wb[nom]
            return i
        return None

    def E(nom):
        return explic.get(nom, {})

    def ent(df):
        return [lisible(c) for c in df.columns]

    def A(cle, a):
        return AGR.ref(cle, a)

    S3, S4, S7 = q("3_sinistralite_resultat"), q("4_bilan"), q("7_rendement_placements")
    i4, j4 = P["gre"]

    def c4(col):
        return plage("D4_grenier_prets", col, i4, j4)

    # ---------------- 10_marge_solvabilite
    df = raw["10_marge_solvabilite"]
    o = Onglet(wb, "10_marge_solvabilite", "Marge de solvabilité exigée et disponible", place("10_marge_solvabilite"), E("10_marge_solvabilite"))
    lig = []
    for a in EX:
        e18 = f"=0.18*{A('cotisations_acquises', a - 1)}" if a > 2020 else None
        e26 = f"=0.26*AVERAGE({A('prest_maladie', a)},{A('prest_maladie', a - 1)},{A('prest_maladie', a - 2)})" if a >= 2022 else None
        ex_ = "=MAX(B{r},C{r})" if a > 2020 else None
        lig.append([a, e18, e26, ex_, "=" + A("reserves_legales_111", a), "=" + A("reserves_facultatives_118", a), "=" + A("reserve_disponible_106", a), 0,
                    "=E{r}+F{r}+G{r}", "=" + A("ran_debiteur_129", a), "=" + A("fonds_propres", a),
                    '=IF(D{r}="","",I{r}/D{r})' if a > 2020 else None, '=IF(D{r}="","",K{r}/D{r})' if a > 2020 else None, "=D{r}/3" if a > 2020 else None])
    h, d, f = o.tableau(ent(df), lig, [ENT] + [FCFA] * 10 + [NB2, NB2, FCFA], source="Source : onglet 10_agregats ; Décision n° 05/2022, annexe, § 1.1.1 ; règlement d'exécution n° 004/2022, art. 23")
    ROW["10ms"] = {a: d + i for i, a in enumerate(EX)}
    controler("10_marge_solvabilite", h, df)
    SM = q("10_marge_solvabilite")

    # ---------------- 10_ratios
    df = raw["10_ratios"]
    RT = {x["code"]: x for x in Q["ratios"]}
    mois = {a: (8 if a == 2026 else 12) for a in EX}

    def N(cle, a):
        return A(cle, a)

    num_den = {
        "SIN": lambda a: (N("charges_prestations", a), N("cotisations_acquises", a)),
        "ROB": lambda a: (f"{N('reserves_legales_111', a)}*{mois[a]}", N("charges_prestations", a)),
        "FEP": lambda a: (f"({N('fonds_propres', a)}+{N('provisions_190', a)}+{N('provisions_reglementees_15', a)})",
                          f"({N('dettes_ct_pcms', a)}+{N('emprunts_16', a)}+{N('provisions_190', a)}+{N('provisions_reglementees_15', a)}+{N('provision_196', a)})"),
        "LIQ": lambda a: (N("tresorerie_actif", a), f"({N('dettes_ct_pcms', a)}+{N('var_provisions_techniques', a)})"),
        "EQS": lambda a: (N("immob_nettes", a), f"({N('fonds_propres', a)}+{N('provisions_techniques', a)}+{N('provision_196', a)}+{N('emprunts_16', a)})"),
        "AUT": lambda a: (N("cotisations_acquises", a), N("total_charges", a)),
        "EXN": lambda a: (N("resultat_net", a), N("cotisations_acquises", a)),
        "TFP": lambda a: (f"({N('total_produits', a)}-{N('subventions', a)})", N("total_charges", a)),
        "RL1": lambda a: (N("actifs_dedies_epargne", a), N("epargne", a)),
        "RL2": lambda a: (f"({N('creances_ct', a)}+{N('tresorerie_actif', a)})", N("dettes_ct_pcms", a)),
        "RL3": lambda a: (N("tresorerie_actif", a), N("dettes_ct_pcms", a)),
        "RL4": lambda a: ("0", f"(({N('sorties_epargne', a)}+{N('charges_fonctionnement', a)}*{N('epargne', a)}/({N('epargne', a)}+{N('cotisations_acquises', a)}))/4)"),
        "RL5": lambda a: (N("fonds_propres", a), f"({N('sorties_epargne', a)}+{N('sorties_epargne', a - 1)})"),
        "CHF": lambda a: (N("charges_fonctionnement", a), N("cotisations_acquises", a)),
        "CHB": lambda a: (N("charges_fonctionnement_brut", a), N("total_produits", a)),
        "CUP": lambda a: (N("charges_fonctionnement", a), N("charges_prestations", a)),
        "CHP": lambda a: (N("charges_personnel", a), N("charges_fonctionnement_brut", a)),
        "DPP": lambda a: (f"{N('fournisseurs_cliniques', a)}*{365 if a < 2026 else 243}", N("charges_prestations", a)),
        "FOR": lambda a: (N("cout_formation", a), N("total_produits", a)),
    }
    retr = {
        "LIQ": lambda a: f"{N('tresorerie_actif', a)}/({N('dettes_ct_hors_epargne', a)}+{N('var_provisions_techniques', a)})",
        "RL2": lambda a: f"({N('creances_ct', a)}+{N('tresorerie_actif', a)})/{N('dettes_ct_hors_epargne', a)}",
        "RL3": lambda a: f"{N('tresorerie_actif', a)}/{N('dettes_ct_hors_epargne', a)}",
        "EQS": lambda a: f"({N('immob_nettes', a)}+{N('emplois_longs_tiers', a)})/({N('fonds_propres', a)}+{N('provisions_techniques', a)}+{N('provision_196', a)}+{N('emprunts_16', a)}+{N('epargne', a)})",
    }
    cols = list(df.columns)
    ci = {c: j for j, c in enumerate(cols)}
    o = Onglet(wb, "10_ratios", "Ratios prudentiels de la Décision n° 05/2022, 2020 au 31/08/2026", place("10_ratios"), E("10_ratios"), ncol=12,
               python=("Seuls les taux de fidélisation (FID) et de croissance des bénéficiaires (CRB) sont calculés par Python (« ratios_airms_fppn.py ») : ils comparent les listes de matricules d'une année à l'autre "
                       "(lignes individuelles des fichiers de précomptes, trop nombreuses pour Excel). Rapport, partie 10, « Tableau de bord (3/4) : gestion ». Toutes les autres valeurs et tous les statuts sont des formules."))
    h0 = o.r
    lig, pyc = [], set()
    ROW["10r"] = {}
    for i, r in enumerate(df.itertuples(index=False)):
        rr = h0 + 1 + i
        code = r[ci["Code"]]
        x = RT[code]
        test = tuple(x["test"]) if x.get("test") else None
        l_ = list(r[:8])
        for a in EX:
            v_orig = r[ci[str(a)]]
            if pd.isna(v_orig):
                l_ += [None, r[ci[f"Statut {a}"]]]
                continue
            if code in num_den:
                n_, d_ = num_den[code](a)
                f_ = f"={n_}/{d_}" if code != "RL4" else "=0"
            elif code == "MSE":
                f_ = f"={SM}!$L${ROW['10ms'][a]}"
            elif code == "RL6":
                f_ = f"={S7}!$D${ROW['7'][a]}"
            elif code == "RL7":
                f_ = "=" + N("comptes_courants_521", a)
            else:
                f_ = float(v_orig)
                pyc.add(ci[str(a)])
            vcell = f"{L(ci[str(a)] + 1)}{rr}"
            l_ += [f_, statut_f(test, vcell, a)]
        # statut global, deux exercices consécutifs, observation
        sc = [f"{L(ci[f'Statut {a}'] + 1)}{rr}" for a in (2022, 2023, 2024, 2025)]
        has = any(pd.notna(r[ci[str(a)]]) for a in (2022, 2023, 2024, 2025))
        nc = [f'OR(LEFT({c},3)="Non",LEFT({c},9)="Au-dessus")' for c in sc]
        if test is None or not has:
            l_.append(r[ci["Statut global (2022 à 2025)"]])
        else:
            rng = f"{sc[0]}:{sc[-1]}"
            l_.append(f'=IF(OR({",".join(nc)}),"Non conforme",IF(COUNTIF({rng},"Conforme*")=4,"Conforme","En attente"))')
            # les statuts sont espacés d'une colonne : la plage contient aussi les valeurs, d'où le test explicite
            l_[-1] = f'=IF(OR({",".join(nc)}),"Non conforme",IF(AND({",".join(f"LEFT({c},8)=" + chr(34) + "Conforme" + chr(34) for c in sc)}),"Conforme","En attente"))'
        l_.append(f'=IF(OR(AND({nc[0]},{nc[1]}),AND({nc[1]},{nc[2]}),AND({nc[2]},{nc[3]})),"Oui","Non")' if (test is not None and has) else r[ci["Non conforme deux exercices consécutifs"]])
        l_.append(r[ci["Observation"]] if pd.notna(r[ci["Observation"]]) else "")
        for a in EX:
            v = r[ci[f"Retraité {a}"]]
            if pd.isna(v):
                l_.append(None)
            elif code in retr:
                l_.append("=" + retr[code](a))
            elif code == "MSE":
                l_.append(f"={SM}!$M${ROW['10ms'][a]}")
            else:
                l_.append(float(v))
        sr = r[ci["Statut retraité"]]
        if pd.notna(sr) and test is not None:
            conds = []
            for a in (2022, 2023, 2024, 2025):
                if pd.notna(r[ci[f"Retraité {a}"]]):
                    c = f"{L(ci[f'Retraité {a}'] + 1)}{rr}"
                    conds.append(f"NOT({test_ok(test, c)})")
            l_.append(f'=IF(OR({",".join(conds)}),"Non conforme","Conforme")' if conds else sr)
        else:
            l_.append(sr if pd.notna(sr) else None)
        lig.append(l_)
        ROW["10r"][code] = rr
    fmts = [ENT, None, None, None, None, None, None, None]
    for a in EX:
        fmts += ["0.0000", None]
    fmts += [None, None, None] + ["0.0000"] * 7 + [None]
    h, d, f = o.tableau([str(c) for c in cols], lig, fmts, py_cols=tuple(sorted(pyc)), largeurs=[8, 8, 14, 30, 40, 30, 22, 30] + [11, 16] * 7 + [16, 14, 30] + [11] * 7 + [14],
                        source="Source : onglets 10_agregats, 10_marge_solvabilite et 7_rendement_placements ; Décision n° 05/2022/COM/UEMOA, annexe")
    for i, r in enumerate(df.itertuples(index=False)):
        pct = "%" in str(r[ci["Norme"]]) and "mois" not in str(r[ci["Norme"]])
        for a in EX:
            o.ws.cell(d + i, ci[str(a)] + 1).number_format = PCT if pct else ("#,##0" if r[ci["Code"]] in ("RL7",) else NB2)
            o.ws.cell(d + i, ci[f"Retraité {a}"] + 1).number_format = PCT if pct else NB2
    num_cols = [ci[str(a)] for a in EX] + [ci[f"Retraité {a}"] for a in EX]
    controler("10_ratios", h, df, cols=num_cols)

    # ---------------- 10_ratios_detail
    df = raw["10_ratios_detail"]
    o = Onglet(wb, "10_ratios_detail", "Numérateurs, dénominateurs, valeurs et statuts des ratios", place("10_ratios_detail"), E("10_ratios_detail"))
    SR = q("10_ratios")
    lig = []
    for r in df.itertuples():
        code, a = r.Code, int(r.Exercice)
        if code in num_den:
            n_, d_ = num_den[code](a)
            if code == "ROB":
                n_ = N("reserves_legales_111", a)
            n_f, d_f = ("=" + n_ if pd.notna(r.Numérateur) else None), ("=" + d_ if pd.notna(r.Dénominateur) else None)
        elif code == "MSE":
            n_f, d_f = f"={SM}!$I${ROW['10ms'][a]}", (f"={SM}!$D${ROW['10ms'][a]}" if pd.notna(r.Dénominateur) else None)
        else:
            n_f, d_f = (float(r.Numérateur) if pd.notna(r.Numérateur) else None), (float(r.Dénominateur) if pd.notna(r.Dénominateur) else None)
        vc = f"{SR}!${L(ci[str(a)] + 1)}${ROW['10r'][code]}"
        scl = f"{SR}!${L(ci[f'Statut {a}'] + 1)}${ROW['10r'][code]}"
        lig.append([code, a, n_f, d_f, (f'=IF({vc}="","",{vc})' if pd.notna(r.Valeur) else None), f"={scl}"])
    h, d, f = o.tableau(ent(df), lig, [None, ENT, FCFA, FCFA, "0.0000", None], source="Source : onglets 10_agregats et 10_ratios")
    controler("10_ratios_detail", h, df, cols=[2, 3, 4])

    # ---------------- 10_affectation_excedents
    df = raw["10_affectation_excedents"]
    o = Onglet(wb, "10_affectation_excedents", "Affectation des excédents et dotation à la réserve obligatoire", place("10_affectation_excedents"), E("10_affectation_excedents"))
    SE = q("10_excedent_par_risque")
    hE = None
    lig = []
    for a in (2023, 2024, 2025):
        lig.append([a, "=" + A("resultat_net", a - 1), "=" + A("ran_crediteur_121", a - 1), f"={A('reserves_legales_111', a)}-{A('reserves_legales_111', a - 1)}",
                    f"={A('reserves_facultatives_118', a)}-{A('reserves_facultatives_118', a - 1)}", f"={A('ran_debiteur_129', a)}-{A('ran_debiteur_129', a - 1)}",
                    "=D{r}/B{r}", f"__EXR_{a - 1}__", "=D{r}-H{r}"])
    h_aff, d_aff, f_aff = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, FCFA, FCFA, PCT, FCFA, FCFA], source="Source : onglets 10_agregats et 10_excedent_par_risque ; statuts du FPPN, art. 81 et 90")
    controler("10_affectation_excedents", h_aff, df)
    o_aff = o

    # ---------------- 10_excedent_par_risque
    df = raw["10_excedent_par_risque"]
    o = Onglet(wb, "10_excedent_par_risque", "Excédent par risque (maladie, décès)", place("10_excedent_par_risque"), E("10_excedent_par_risque"))
    lig = []
    for a in EX:
        lig.append([a, "=" + A("cot_maladie_701", a), "=" + A("prest_maladie", a), f"=({A('charges_fonctionnement', a)}-{A('prime_6518', a)})*B{{r}}/(B{{r}}+F{{r}})", "=B{r}-C{r}-D{r}",
                    "=" + A("cot_deces_702", a), f"={A('prest_deces', a)}+{A('prime_6518', a)}", f"=({A('charges_fonctionnement', a)}-{A('prime_6518', a)})*F{{r}}/(B{{r}}+F{{r}})", "=F{r}-G{r}-H{r}",
                    "=K{r}-E{r}-I{r}", "=" + A("resultat_net", a), "=B{r}/(B{r}+F{r})", "=0.3*MAX(E{r},0)+0.5*MAX(I{r},0)"])
    h, d, f = o.tableau(ent(df), lig, [ENT] + [FCFA] * 10 + [PCT, FCFA], source="Source : onglet 10_agregats ; frais répartis au prorata des cotisations")
    ROW["10ex"] = {a: d + i for i, a in enumerate(EX)}
    controler("10_excedent_par_risque", h, df)
    for i, a in enumerate((2023, 2024, 2025)):
        o_aff.ws.cell(d_aff + i, 8).value = f"={SE}!$M${ROW['10ex'][a - 1]}"

    # ---------------- 10_actif_mobilisation
    df = raw["10_actif_mobilisation"]
    o = Onglet(wb, "10_actif_mobilisation", "Actif par délai de mobilisation et passif à moins d'un an", place("10_actif_mobilisation"), E("10_actif_mobilisation"))
    lig = []
    for a, lib in ((2025, "31/12/2025"), (2026, "31/08/2026")):
        lig.append([lib, "=" + A("disponibilites", a), f"={A('placements_tresorerie', a)}+{A('debiteurs_3_4', a)}-{A('emplois_longs_tiers', a)}",
                    f"={A('prets_pep_ratios', a)}+{A('grenier', a)}",
                    f"={A('immob_nettes', a)}-{A('prets_pep_ratios', a)}+{A('cmp', a)}+{A('creance_hopital', a)}+{A('solde_47111201_485', a)}",
                    f"={A('dettes_ct_hors_epargne', a)}+{A('provisions_techniques', a)}", "=" + A("epargne", a), "=B{r}+C{r}+D{r}+E{r}"])
    h, d, f = o.tableau(ent(df), lig, [None] + [FCFA] * 7, source="Source : onglet 10_agregats")
    ROW["10mob"] = {"31/12/2025": d, "31/08/2026": d + 1}
    controler("10_actif_mobilisation", h, df)

    # ---------------- 10_classement_comptes
    df = raw["10_classement_comptes"]
    o = Onglet(wb, "10_classement_comptes", "Classement des comptes selon le PCMS et selon la nature des risques", place("10_classement_comptes"), E("10_classement_comptes"), ncol=10,
               python="Les colonnes de rubrique, de tiroir et de classement appliquent la règle écrite du script « ratios_airms_fppn.py » (fonctions rubrique et tiroir), d'après le numéro du compte et le plan comptable des mutuelles sociales (PCMS, mai 2012). Les soldes sont des formules sur l'onglet D1_balances.")
    D1 = P["bal"]
    cpt, exr, net = (plage("D1_balances", c, D1[0], D1[1]) for c in ("A", "C", "L"))
    lig = []
    for r in df.itertuples(index=False):
        l_ = [str(r[0])] + list(r[1:7])
        l_ += [f'=SUMIFS({net},{cpt},$A{{r}},{exr},{a})' for a in EX]
        lig.append(l_)
    h, d, f = o.tableau(ent(df), lig, [None] * 7 + [FCFA] * 7, py_cols=(2, 3, 4, 5), largeurs=[12, 36, 30, 24, 24, 30, 22] + [16] * 7, source="Source : onglet D1_balances ; PCMS (mai 2012)")
    controler("10_classement_comptes", h, df, cols=list(range(7, 14)))

    # ---------------- 10_tableau_de_bord : valeurs reprises de 10_ratios par formules
    if "10_tableau_de_bord" in wb.sheetnames:
        tb = wb["10_tableau_de_bord"]
        cols_tb = {}
        for j in range(1, tb.max_column + 1):
            cols_tb[str(tb.cell(5, j).value)] = j
        for i, code in enumerate(df_codes := [x for x in raw["10_ratios"]["Code"]]):
            rr = 6 + i
            for a in (2022, 2023, 2024, 2025, 2026):
                if str(a) in cols_tb and tb.cell(rr, cols_tb[str(a)]).value is not None:
                    tb.cell(rr, cols_tb[str(a)]).value = f'={SR}!${L(ci[str(a)] + 1)}${ROW["10r"][code]}'
            if "Statut 2022-2025" in cols_tb:
                tb.cell(rr, cols_tb["Statut 2022-2025"]).value = f'={SR}!${L(ci["Statut global (2022 à 2025)"] + 1)}${ROW["10r"][code]}'
            if "Statut retraité" in cols_tb and tb.cell(rr, cols_tb["Statut retraité"]).value is not None:
                tb.cell(rr, cols_tb["Statut retraité"]).value = f'={SR}!${L(ci["Statut retraité"] + 1)}${ROW["10r"][code]}'

    # ---------------- 12_provisions_inventaire
    df = raw["12_provisions_inventaire"]
    o = Onglet(wb, "12_provisions_inventaire", "Inventaire des provisions et dépréciations (FCFA)", place("12_provisions_inventaire"), E("12_provisions_inventaire"))
    lig = [[str(r[0]), r[1]] + [f'=-SUMIFS({net},{cpt},$A{{r}},{exr},{a})' for a in EX] for r in df.itertuples(index=False)]
    h, d, f = o.tableau(ent(df), lig, [None, None] + [FCFA] * 7, source="Source : onglet D1_balances (comptes 15, 19, 29, 39, 49, 59)")
    controler("12_provisions_inventaire", h, df, cols=list(range(2, 9)))

    # ---------------- 12_psap_rapprochement
    df = raw["12_psap_rapprochement"]
    o = Onglet(wb, "12_psap_rapprochement", "Provision maladie, dettes envers les cliniques et prestations", place("12_psap_rapprochement"), E("12_psap_rapprochement"))
    lig = [[a, "=" + A("prov_maladie", a), "=" + A("dettes_cliniques_40110000", a), "=" + A("prest_maladie", a), mois[a], "=B{r}/(D{r}/E{r})", "=C{r}/(D{r}/E{r})", "=B{r}/C{r}",
            "=" + A("prest_deces", a), f"={A('prov_obseques', a)}/(I{{r}}/E{{r}})"] for a in EX]
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, ENT, NB2, NB2, NB2, FCFA, NB2], source="Source : onglet 10_agregats")
    controler("12_psap_rapprochement", h, df)

    # ---------------- 12_grenier_tranches
    df = raw["12_grenier_tranches"]
    o = Onglet(wb, "12_grenier_tranches", "Encours du Grenier par tranche de retard", place("12_grenier_tranches"), E("12_grenier_tranches"))
    TR = [(-1, 0), (0, 30), (30, 90), (90, 180), (180, 365), (365, 100000)]
    lig = []
    for k, r in enumerate(df.itertuples()):
        dt_ = pd.Timestamp(r.arrete)
        lo, hi = TR[k % 6]
        dc = f"DATE({dt_.year},{dt_.month},{dt_.day})"
        base_r = o.r + 1 + (k // 6) * 6
        lig.append([r.arrete, r.tranche, f'=COUNTIFS({c4("A")},{dc},{c4("E")},">{lo}",{c4("E")},"<={hi}")',
                    f'=SUMIFS({c4("D")},{c4("A")},{dc},{c4("E")},">{lo}",{c4("E")},"<={hi}")', f"=D{{r}}/SUM(D${base_r}:D${base_r + 5})"])
    h, d, f = o.tableau(ent(df), lig, [None, None, ENT, FCFA, PCT], source="Source : onglet D4_grenier_prets")
    controler("12_grenier_tranches", h, df, cols=[2, 3, 4])

    # ---------------- 12_gna_47111106, 12_pidr_196
    for nom, titre, cles, src in (("12_gna_47111106", "Compte de tiers GNA (47111106)", ["gna_solde", "gna_debit", "gna_credit"], "compte 47111106"),
                                  ("12_pidr_196", "Provision pour départ à la retraite du personnel (compte 196)", ["pidr_196", "dotation_6911", "reprise_7911"], "comptes 196, 6911 et 7911")):
        df = raw[nom]
        o = Onglet(wb, nom, titre, place(nom), E(nom))
        lig = [[a] + ["=" + A(k, a) for k in cles] for a in EX]
        h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA], source="Source : onglet 10_agregats (" + src + ")")
        controler(nom, h, df)

    # ---------------- 12_sinistralite_variantes
    df = raw["12_sinistralite_variantes"]
    o = Onglet(wb, "12_sinistralite_variantes", "Sinistralité : formule du texte et variantes", place("12_sinistralite_variantes"), E("12_sinistralite_variantes"))
    lig = []
    for a in EX:
        lig.append([a, "=" + A("charges_60", a), "=" + A("dot_rep_6901_7901", a), "=" + A("cotisations_70", a), "=" + A("dot_rep_7906_6906", a), "=B{r}+C{r}", "=D{r}+E{r}", "=F{r}/G{r}",
                    "=" + A("cot_medoc_7022", a), "=F{r}/(G{r}-I{r})", "=" + A("bons_60800007_8", a), "=(F{r}-K{r})/G{r}",
                    f"=B{{r}}-{A('prest_maladie', a)}-{A('prest_deces', a)}", "=" + A("cot_maladie_7010", a), "=" + A("prest_maladie", a), "=O{r}/N{r}",
                    "=" + A("cot_deces_7021", a), "=" + A("prest_deces", a), "=R{r}/Q{r}"])
    h, d, f = o.tableau(ent(df), lig, [ENT, FCFA, FCFA, FCFA, FCFA, FCFA, FCFA, PCT, FCFA, PCT, FCFA, PCT, FCFA, FCFA, FCFA, PCT, FCFA, FCFA, PCT], source="Source : onglet 10_agregats")
    controler("12_sinistralite_variantes", h, df)
    return ROW


# ---------------------------------------------------------------------------
# Onglet 0_controle : chaque valeur recalculée face à la valeur publiée
# ---------------------------------------------------------------------------
def onglet_controle(wb):
    o = Onglet(wb, "0_controle", "Contrôle : chaque formule redonne la valeur publiée dans le rapport", 2, {
        "À quoi sert cet onglet": "Prouver que les formules du classeur redonnent exactement les chiffres du rapport. Chaque ligne compare une cellule recalculée à la valeur publiée.",
        "D'où viennent les chiffres": "Valeur publiée : résultat des scripts Python de la mission, repris dans le rapport. Valeur recalculée : formule de la cellule indiquée.",
        "Comment le calcul est fait": "Écart = valeur recalculée - valeur publiée. « Identique » si l'écart est au plus de 0,5 FCFA pour un montant, ou de 0,000001 pour un taux ou un petit nombre.",
        "Comment lire le résultat": "Toutes les lignes doivent être « Identique ». Le total en tête compte les écarts.",
        "Valable pour": "Onglets 1 à 12 reconstruits en formules.", "Limites": "Les colonnes calculées par Python (fond bleu clair) sont reprises telles quelles et ne sont pas contrôlées ici."}, ncol=6)
    ws = o.ws
    ws.cell(o.r, 1, "Écarts trouvés").font = Font(name=POLICE, bold=True)
    tot = ws.cell(o.r, 2)
    o.r += 2
    lig = []
    for onglet, cell, v in CTRL:
        lig.append([onglet, cell, v, f"={q(onglet)}!{cell}", '=IF(ISNUMBER(D{r}),D{r}-C{r},"")',
                    '=IF(ISNUMBER(D{r}),IF(ABS(E{r})<=IF(ABS(C{r})>=100,0.5,0.000001),"Identique","Écart"),"Vide")'])
    h, d, f = o.tableau(["Onglet", "Cellule", "Valeur publiée", "Valeur recalculée", "Écart", "Contrôle"], lig, [None, None, "#,##0.000000", "#,##0.000000", "#,##0.000000", None],
                        largeurs=[30, 10, 22, 22, 18, 12], source="Source : rapport (valeurs publiées) et formules du classeur")
    tot.value = f'=COUNTIF(F{d}:F{f},"Écart")+COUNTIF(F{d}:F{f},"Vide")'
    return d, f


def verifier(chemin_recalcule, d, f):
    """Lit le classeur recalculé par LibreOffice et renvoie les lignes de contrôle en écart."""
    w = openpyxl.load_workbook(chemin_recalcule, data_only=True, read_only=True)["0_controle"]
    mauvais = []
    for row in w.iter_rows(min_row=d, max_row=f, values_only=True):
        if row[5] != "Identique":
            mauvais.append(row[:6])
    return mauvais


# ---------------------------------------------------------------------------
# Onglets 13 à 17 : montants comptables recopiés remplacés par des renvois aux agrégats
# ---------------------------------------------------------------------------
def relier_existants(wb, AGR, P):
    def A(cle, a):
        return AGR.ref(cle, a)

    def poser(ws, cell, formule):
        v = ws[cell].value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            CTRL.append((ws.title, cell, float(v)))
            ws[cell].value = formule
    i3, j3 = P["pre"]

    def c3(col):
        return plage("D3_precomptes_mensuels", col, i3, j3)
    if "13c_donnees" in wb.sheetnames:
        ws = wb["13c_donnees"]
        F = {
            "COT": lambda a: A("cotisations_acquises", a), "PRE": lambda a: A("charges_prestations", a), "CF": lambda a: A("charges_fonctionnement", a),
            "P6518": lambda a: A("prime_6518", a), "CFB": lambda a: A("charges_fonctionnement_brut", a), "PROD": lambda a: A("total_produits", a),
            "RN": lambda a: A("resultat_net", a), "RES": lambda a: A("reserves_legales_111", a),
            "FP": lambda a: f"{A('fonds_propres', a)}+{A('provisions_190', a)}+{A('provisions_reglementees_15', a)}",
            "DET": lambda a: f"{A('dettes_ct_pcms', a)}+{A('emprunts_16', a)}+{A('provisions_190', a)}+{A('provisions_reglementees_15', a)}+{A('provision_196', a)}",
            "TRE": lambda a: A("tresorerie_actif", a), "DCT": lambda a: f"{A('dettes_ct_pcms', a)}+{A('var_provisions_techniques', a)}",
            "RL2N": lambda a: f"{A('creances_ct', a)}+{A('tresorerie_actif', a)}", "EMP": lambda a: A("immob_nettes", a),
            "RESL": lambda a: f"{A('fonds_propres', a)}+{A('provisions_techniques', a)}+{A('provision_196', a)}+{A('emprunts_16', a)}",
            "ACTEP": lambda a: A("actifs_dedies_epargne", a), "EPA": lambda a: A("epargne", a),
            "RL4D": lambda a: f"({A('sorties_epargne', a)}+{A('charges_fonctionnement', a)}*{A('epargne', a)}/({A('epargne', a)}+{A('cotisations_acquises', a)}))/4",
            "FP2": lambda a: A("fonds_propres", a), "SORT2": lambda a: (f"{A('sorties_epargne', a)}+{A('sorties_epargne', a - 1)}" if a > 2020 else None),
            "P129": lambda a: "-" + A("ran_debiteur_129", a), "COTM": lambda a: A("cot_maladie_7010", a), "PREM": lambda a: A("prest_maladie", a),
            "LIQr": lambda a: f"{A('tresorerie_actif', a)}/({A('dettes_ct_hors_epargne', a)}+{A('var_provisions_techniques', a)})",
            "RL2r": lambda a: f"({A('creances_ct', a)}+{A('tresorerie_actif', a)})/{A('dettes_ct_hors_epargne', a)}",
            "RL3r": lambda a: f"{A('tresorerie_actif', a)}/{A('dettes_ct_hors_epargne', a)}",
            "EQSr": lambda a: f"({A('immob_nettes', a)}+{A('emplois_longs_tiers', a)})/({A('fonds_propres', a)}+{A('provisions_techniques', a)}+{A('provision_196', a)}+{A('emprunts_16', a)}+{A('epargne', a)})",
        }
        hdr = next(r for r in range(1, 12) if ws.cell(r, 1).value == "Code")
        cols = {int(ws.cell(hdr, c).value): c for c in range(3, 10) if str(ws.cell(hdr, c).value).isdigit()}
        for r in range(hdr + 1, ws.max_row + 1):
            code = ws.cell(r, 1).value
            if code in F:
                for a, c in cols.items():
                    f_ = F[code](a)
                    if f_ is not None:
                        poser(ws, f"{L(c)}{r}", "=" + f_)
    if "16_efc_taux_servi" in wb.sheetnames:
        ws = wb["16_efc_taux_servi"]
        for r in range(5, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, int) and a in EX:
                poser(ws, f"B{r}", "=" + A("epargne_efc", a))
                if 2021 <= a <= 2025:
                    poser(ws, f"I{r}", "=" + A("produits_fin_77", a))
    if "17_controles_croises" in wb.sheetnames:
        ws = wb["17_controles_croises"]
        for r in range(5, 12):
            a = ws.cell(r, 1).value
            if isinstance(a, int) and a in EX:
                poser(ws, f"C{r}", "=" + A("total_actif", a))
                poser(ws, f"F{r}", "=" + A("resultat_net", a))
                poser(ws, f"I{r}", "=" + A("tresorerie_actif", a))
        for r in range(14, 22):
            a, m = ws.cell(r, 1).value, ws.cell(r, 2).value
            if isinstance(a, int) and isinstance(m, int):
                poser(ws, f"D{r}", f'=SUMIFS({c3("E")},{c3("A")},"8Z6",{c3("B")},"retour",{c3("C")},{a},{c3("D")},{m})')
    if "14c_pidr_synthese" in wb.sheetnames:
        ws = wb["14c_pidr_synthese"]
        for r in range(5, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, int) and a in EX:
                poser(ws, f"H{r}", "=" + A("pidr_196", a))
    if "15_mortalite_sinistralite" in wb.sheetnames:
        ws = wb["15_mortalite_sinistralite"]
        for r in range(5, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, int) and 2020 <= a <= 2025:
                poser(ws, f"I{r}", "=" + A("cot_deces_7021", a))
                poser(ws, f"J{r}", "=" + A("prest_deces", a))
    if "14t_table_cima_f" in wb.sheetnames and "mort" in P:
        ws = wb["14t_table_cima_f"]
        md, mf, noms, _ = P["mort"]
        col = L(2 + noms.index("CIMA F"))
        for r in range(1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, int) and 0 <= a <= 110:
                poser(ws, f"B{r}", f"=INDEX({q('D8_tables_mortalite')}!${col}${md}:${col}${mf},{a + 1})")
