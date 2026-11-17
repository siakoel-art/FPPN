# -*- coding: utf-8 -*-
"""Classeur des deux anomalies (dépenses hors soins, provision de la clinique), commenté en mots simples.
Utilisation : python excel.py <dossier de travail> <sortie.xlsx>"""
import json
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

W, OUT = sys.argv[1], sys.argv[2]
if os.path.exists(OUT):
    raise SystemExit("le fichier existe déjà, choisir une nouvelle version : " + OUT)
R = json.load(open(os.path.join(W, "resultats.json"), encoding="utf-8"))
T = pd.read_parquet(os.path.join(W, "t1_avant_apres.parquet"))
DET = pd.read_parquet(os.path.join(W, "t1_detail.parquet"))
OD = pd.read_parquet(os.path.join(W, "t2_od.parquet"))
SIN = pd.read_parquet(os.path.join(W, "t2_sin.parquet"))
C26 = pd.read_parquet(os.path.join(W, "t2_cmp_2026.parquet"))
AUT = pd.read_parquet(os.path.join(W, "t2_autres.parquet"))
AUTM = pd.read_parquet(os.path.join(W, "t2_autres_mixtes.parquet"))
GL = R["sources"]["grands_livres"]
BAL = R["sources"]["balances"]
BAL["2021"] = BAL["2021"]
GL26 = GL["2026"]
GL25 = GL["2025"]
BAL26 = BAL["2026"]

NOIR, ORANGE, GRIS, BLANC, JAUNE = "060606", "EFAC62", "BBBCBC", "FFFFFF", "FFFF00"
F = "Helvetica"
fn = lambda **k: Font(name=F, size=k.pop("size", 11), **k)
TH = Side(style="thin", color=GRIS)
BORD = Border(left=TH, right=TH, top=TH, bottom=TH)
FCFA = '#,##0 "FCFA";-#,##0 "FCFA"'
PCT = '0.00 %'
wb = Workbook()


def memo(cell, texte, l=360, h=150):
    c = Comment("Mémo : " + texte, "Adenka")
    c.width, c.height = l, h
    cell.comment = c


def titre(ws, t, st):
    ws["A1"] = t
    ws["A1"].font = fn(size=16, bold=True, color=NOIR)
    ws["A2"] = st
    ws["A2"].font = fn(size=12, color="B07A3C", bold=True)
    ws.sheet_view.showGridLines = False


def entete(ws, r, cols, larg=None):
    for j, c in enumerate(cols, 1):
        x = ws.cell(r, j, c)
        x.font = fn(bold=True, color=BLANC)
        x.fill = PatternFill("solid", fgColor=NOIR)
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        x.border = BORD
    ws.row_dimensions[r].height = 45
    if larg:
        for j, w in enumerate(larg, 1):
            ws.column_dimensions[get_column_letter(j)].width = w


def val(ws, r, c, v, fmt=None, bold=False, fill=None, wrap=False):
    x = ws.cell(r, c, v)
    x.font = fn(bold=bold)
    x.border = BORD
    if fmt:
        x.number_format = fmt
    if fill:
        x.fill = PatternFill("solid", fgColor=fill)
    x.alignment = Alignment(wrap_text=wrap, vertical="top")
    return x


def note(ws, r, texte, h=None, col=1, ncol=10, couleur=None):
    ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=ncol)
    x = ws.cell(r, col, texte)
    x.font = fn(size=11, italic=couleur is None)
    x.alignment = Alignment(wrap_text=True, vertical="top")
    if couleur:
        x.fill = PatternFill("solid", fgColor=couleur)
    ws.row_dimensions[r].height = h or 18 * (1 + len(texte) // 140)


AN = lambda a: f"{a} (8 mois)" if a == 2026 else str(a)

# ====================================================================== 0 Lis-moi
ws = wb.active
ws.title = "0_Lis-moi"
titre(ws, "FPPN : deux anomalies dans les dépenses de soins", "Classeur de travail pour relecture interne, préparé le 17/11/2026")
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
L = [("De quoi parle ce classeur ?", "Le Fonds paie des soins pour les policiers malades. On compare ce qu'il paie à ce qu'il reçoit en cotisations : c'est la sinistralité. Nous avons trouvé deux erreurs qui gonflent ce que le Fonds dit payer pour les soins."),
     ("Anomalie 1 (onglets 1_...)", "Dans les comptes de soins, il y a aussi des dons, des cadeaux, des imprimés et des honoraires d'un cabinet. Ce ne sont pas des soins. On les enlève et on refait les calculs, de 2021 au 31/08/2026."),
     ("Anomalie 2 (onglets 2_...)", "Fin 2025, le Fonds a mis de côté 255 000 000 FCFA pour les factures de sa clinique d'octobre, novembre et décembre 2025 (une provision : de l'argent réservé pour une facture qu'on attend). Les vraies factures sont arrivées en janvier 2026. Le Fonds les a comptées en 2026 sans retirer l'argent réservé. Les mêmes soins sont donc comptés deux fois."),
     ("Onglet 3_Ensemble_2026", "Ce que deviennent les chiffres de 2026 quand on corrige les deux anomalies en même temps."),
     ("Onglet Sources", "La liste des fichiers du Fonds utilisés et la façon dont chaque chiffre a été obtenu."),
     ("Les petits triangles rouges", "Passez la souris sur une cellule qui a un petit triangle rouge dans son coin : un mémo explique le chiffre avec des mots simples."),
     ("Les couleurs", "En-tête noir : titre des colonnes. Fond orange clair : chiffre corrigé par Adenka. Fond jaune : point à confirmer par le Fonds. Les ratios sont des formules : cliquez dessus pour voir le calcul."),
     ("Le mot « ratio »", "Un ratio de 73,53 % veut dire : pour 100 FCFA de cotisations reçues, le Fonds a payé 73,53 FCFA de soins."),
     ("Ce qui reste à confirmer", "Le lettrage du compte 40110000 (il relie chaque facture à son paiement), les pièces des honoraires du cabinet, et une facture mensuelle de la clinique qui semble manquer en 2026.")]
for i, (a, b) in enumerate(L, 4):
    ws.cell(i, 1, a).font = fn(bold=True)
    x = ws.cell(i, 2, b)
    x.font = fn()
    x.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(i, 1).alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[i].height = 18 * (1 + len(b) // 105)
ws.cell(12, 2).fill = PatternFill("solid", fgColor=JAUNE)

# ====================================================================== 1 Avant / après
ws = wb.create_sheet("1_Avant_apres")
titre(ws, "Anomalie 1 : sans les dépenses qui ne sont pas des soins, la maladie consomme 53,30 % de ses cotisations en 2025, et non 73,53 %",
      "Prestations et sinistralité par année, comptes tels quels puis sans les dépenses hors soins, 2021 au 31/08/2026")
cols = ["Année", "Cotisations maladie (comptes 7010)", "Prestations maladie telles quelles (601 + 60800007)", "Dont hors soins, règles du rapport",
        "Dont hors soins, variante élargie (en plus)", "Soins seuls, règles du rapport", "Ratio maladie avant", "Ratio maladie après (règles du rapport)",
        "Ratio maladie après (variante)", "Toutes cotisations (comptes 70)", "Toutes prestations (comptes 60)", "Hors soins, toutes garanties",
        "Ratio global avant", "Ratio global après (règles du rapport)", "Ratio global après (variante)", "Norme du ratio global"]
entete(ws, 4, cols, [14, 20, 22, 20, 20, 20, 13, 15, 13, 20, 20, 20, 13, 15, 13, 14])
r0 = 5
for i, (an, t) in enumerate(T.iterrows()):
    r = r0 + i
    val(ws, r, 1, AN(an), bold=True)
    val(ws, r, 2, round(t.cot_maladie), FCFA)
    val(ws, r, 3, round(t.prest_maladie), FCFA)
    val(ws, r, 4, round(t.hors_soins_dans_maladie), FCFA)
    val(ws, r, 5, round(t.variante_dans_maladie), FCFA)
    val(ws, r, 6, f"=C{r}-D{r}", FCFA, fill="FBE5CB")
    val(ws, r, 7, f"=C{r}/B{r}", PCT)
    val(ws, r, 8, f"=F{r}/B{r}", PCT, bold=True, fill="FBE5CB")
    val(ws, r, 9, f"=(F{r}-E{r})/B{r}", PCT, fill="FBE5CB")
    val(ws, r, 10, round(t.cot_toutes), FCFA)
    val(ws, r, 11, round(t.prest_toutes), FCFA)
    val(ws, r, 12, round(t.hors_soins), FCFA)
    val(ws, r, 13, f"=K{r}/J{r}", PCT)
    val(ws, r, 14, f"=(K{r}-L{r})/J{r}", PCT, bold=True, fill="FBE5CB")
    val(ws, r, 15, f"=(K{r}-L{r}-E{r})/J{r}", PCT, fill="FBE5CB")
    val(ws, r, 16, "70 % à 80 %")
r25 = r0 + list(T.index).index(2025)
r26 = r0 + list(T.index).index(2026)
memo(ws.cell(4, 2), "C'est l'argent que les policiers ont versé au Fonds pour la garantie maladie pendant l'année. On le lit dans la balance de chaque année, comptes qui commencent par 7010.")
memo(ws.cell(4, 3), "C'est tout ce que le Fonds a rangé dans ses tiroirs « soins maladie » (comptes 601... et 60800007). C'est le chiffre que le Fonds utilise aujourd'hui. On l'a recompté ligne par ligne dans le grand livre, et le total est le même que dans la balance.")
memo(ws.cell(4, 4), "Ce sont les lignes rangées dans le tiroir « soins » alors que ce ne sont pas des soins : dons à des autorités, paniers cadeaux, gadgets, imprimés, honoraires d'un cabinet. On les repère grâce aux mots écrits dans le libellé de chaque ligne. Le détail ligne par ligne est dans l'onglet 1_Detail.")
memo(ws.cell(4, 5), "Ce sont d'autres lignes qui ressemblent aussi à des dépenses hors soins (imprimeur GIS PLUS, sacs, hadj, cérémonies, lobbying, locations) mais que le rapport n'a pas retirées. On les montre à part, pour voir si cela change la conclusion. Cela ne la change pas.")
memo(ws.cell(4, 6), "Soins seuls = prestations telles quelles moins les dépenses hors soins. Exemple 2025 : 1 771 635 361 moins 487 492 003 = 1 284 143 358 FCFA.")
memo(ws.cell(r25, 7), "En 2025, pour 100 FCFA de cotisations maladie, le Fonds dit avoir payé 73,53 FCFA de soins. C'est le chiffre « avant ».")
memo(ws.cell(r25, 8), "Quand on enlève les dons, les cadeaux, les imprimés et les honoraires, il ne reste que 53,30 FCFA de vrais soins pour 100 FCFA de cotisations. La garantie maladie va donc mieux que ce que disent les comptes.")
memo(ws.cell(r26, 7), "Attention : 2026 n'a que 8 mois (janvier à août). Et ce chiffre contient aussi la deuxième anomalie (factures de 2025 comptées en 2026). Voir l'onglet 3_Ensemble_2026.")
memo(ws.cell(4, 13), "Ratio global = toutes les prestations (maladie, décès, frais funéraires) divisées par toutes les cotisations. C'est la formule de la Décision n° 05/2022/COM/UEMOA. Le texte dit qu'il doit rester entre 70 % et 80 %.")
memo(ws.cell(4, 14), "Même calcul que le ratio global avant, mais sans les dépenses hors soins. Elles ne disparaissent pas : le Fonds les a bien payées. Elles doivent juste être rangées dans le bon tiroir (frais de fonctionnement), pas dans les soins.")
r = r26 + 2
note(ws, r, "Comment lire : en 2025, les dépenses hors soins font 487 492 003 FCFA, soit 27,52 FCFA pour 100 FCFA de prestations maladie, soit 27,52 % des prestations maladie. Sans elles, le ratio maladie passe de 73,53 % à 53,30 %.", ncol=12)
note(ws, r + 1, "Ces dépenses ne sont pas effacées. Elles restent des dépenses du Fonds. Elles changent la sinistralité, pas le résultat de l'année.", ncol=12)
note(ws, r + 2, "Avant 2023, le classement ne trouve presque rien : les dons passaient surtout par un autre compte (6583 : 52 270 446 FCFA en 2021 et 61 894 670 FCFA en 2022). Ce compte baisse ensuite et il est vide en 2025. Les montants négatifs de 2021 et 2022 sont des annulations de provisions de compte passées au même compte.", ncol=12)
note(ws, r + 3, f"Sources : prestations lues dans les grands livres « {GL['2021']} » à « {GL26} ». Cotisations lues dans les balances « {BAL['2021']} » à « {BAL26} ». Détail dans l'onglet Sources.", ncol=12)
ws.freeze_panes = "B5"

# ====================================================================== 1 Par type
ws = wb.create_sheet("1_Par_type")
titre(ws, "Les dons sont la plus grosse dépense hors soins : 203 689 140 FCFA en 2025", "Dépenses hors soins rangées dans les comptes de soins, par type et par année (FCFA)")
ty = DET[DET.categorie != "Soins"].pivot_table(index="categorie", columns="an", values="montant", aggfunc="sum").fillna(0)
va = DET[DET.variante != ""].groupby("an").montant.sum()
ans = list(T.index)
entete(ws, 4, ["Type de dépense", "Règle utilisée pour la repérer"] + [AN(a) for a in ans], [36, 50] + [18] * len(ans))
REG = {"Dons (autorités, élus, délégués)": "Le libellé contient DON, K'DO, CADEAU ou PANIER. Les dons pour des soins (dialyse, médicaments, accouchement) restent dans les soins.",
       "Gadgets, imprimés et cartes": "Le libellé contient GADGET, IMPRIME, CARTE, SAC KABA, ou le nom d'un imprimeur (FUTURE VISION, NEW PRINT).",
       "Honoraires d'un cabinet": "Le libellé porte une facture n° 00444..., payée tous les trois mois au même cabinet.",
       "Régularisations comptables": "Le libellé contient RECLASSEMENT FRAIS DE RECOUVREMENT ou PROV CPTE."}
r = 5
for c in REG:
    val(ws, r, 1, c, bold=True, wrap=True)
    val(ws, r, 2, REG[c], wrap=True)
    for j, a in enumerate(ans):
        val(ws, r, 3 + j, round(ty.loc[c, a]) if c in ty.index else 0, FCFA)
    ws.row_dimensions[r].height = 45
    r += 1
val(ws, r, 1, "Total, règles du rapport", bold=True)
val(ws, r, 2, "")
for j in range(len(ans)):
    L_ = get_column_letter(3 + j)
    val(ws, r, 3 + j, f"=SUM({L_}5:{L_}{r - 1})", FCFA, bold=True, fill="FBE5CB")
rt = r
r += 1
val(ws, r, 1, "Variante élargie (en plus)", bold=True, wrap=True)
val(ws, r, 2, "Lignes hors factures de soins (journal autre que SIN) dont le libellé contient GIS PLUS, SACS KABA, HADJ, CEREMONIE, LOBBING, ARBRE DE NOEL, ECRAN GEANT ou LOCAT.", wrap=True)
ws.row_dimensions[r].height = 45
for j, a in enumerate(ans):
    val(ws, r, 3 + j, round(va.get(a, 0)), FCFA)
memo(ws.cell(5, 1), "Des dons d'argent à des autorités, des élus ou des délégués, et des paniers cadeaux. C'est gentil, mais ce n'est pas soigner un malade. Ils doivent aller dans un autre tiroir.")
memo(ws.cell(6, 1), "Des stylos, des polos, des sacs, des fiches et des carnets imprimés avec le nom du Fonds. Ce sont des fournitures, pas des soins.")
memo(ws.cell(7, 1), "Un cabinet envoie une facture tous les trois mois (environ 32 000 000 FCFA chaque fois en 2026). On ne sait pas encore ce qu'il fait. Ce ne sont pas des soins à un malade. Le Fonds doit montrer les pièces.", h=160)
memo(ws.cell(rt, 1), "On additionne les lignes au-dessus. C'est ce qu'on enlève des soins dans l'onglet 1_Avant_apres.")
note(ws, r + 2, "Source : chaque ligne du grand livre de l'année, comptes 6011, 6012, 6013, 6018 et 6080, classée par son libellé. Le détail est dans l'onglet 1_Detail.", ncol=8)
ws.freeze_panes = "C5"

# ====================================================================== 1 Détail
ws = wb.create_sheet("1_Detail")
titre(ws, "Chaque ligne retirée des soins, avec l'endroit exact où on l'a trouvée", "Une ligne par écriture du grand livre (FCFA)")
cols = ["Année", "Date", "Journal", "Pièce", "Compte", "Libellé écrit par le Fonds", "N° facture", "Montant", "Type", "Pourquoi on la retire", "Variante élargie", "Fichier du grand livre", "Ligne dans le fichier"]
entete(ws, 4, cols, [8, 12, 9, 9, 11, 42, 20, 16, 30, 40, 30, 52, 10])
memo(ws.cell(4, 13), "Le numéro de ligne dans le fichier Excel du Fonds. Ouvrez le fichier, allez à cette ligne, et vous verrez exactement la même écriture.")
memo(ws.cell(4, 6), "C'est le petit texte que le comptable du Fonds a écrit pour dire à quoi sert le paiement. C'est grâce à lui qu'on sait que ce n'est pas un soin.")
D_ = DET.sort_values(["an", "categorie", "date"])
for i, x in enumerate(D_.itertuples(), 5):
    vals = [x.an, x.date, x.journal, x.piece, x.compte, x.libelle, x.facture, round(x.montant), x.categorie if x.categorie != "Soins" else "",
            x.regle, x.variante + (f" (mot {x.regle_var})" if x.regle_var else ""), x.fichier, x.ligne]
    for j, v in enumerate(vals, 1):
        c = ws.cell(i, j, v)
        c.font = fn(size=10)
        if j == 8:
            c.number_format = FCFA
        if j == 2:
            c.number_format = "dd/mm/yyyy"
ws.auto_filter.ref = f"A4:M{4 + len(D_)}"
ws.freeze_panes = "A5"

# ====================================================================== 2 Clinique mois
ws = wb.create_sheet("2_Clinique_mois")
titre(ws, "Anomalie 2 : les factures de la clinique d'octobre à décembre 2025 sont comptées en 2025 et encore en 2026",
      "Provision de fin 2025 et vraies factures reçues en 2026, mois par mois (FCFA)")
entete(ws, 4, ["Mois de soins", "Écriture de provision (31/12/2025)", "Argent mis de côté fin 2025 (provision)", "Pièce de la vraie facture (31/01/2026)",
               "Vraie facture reçue en 2026", "Écart : mis de côté moins facture", "Provision annulée en 2026 ?"], [18, 22, 22, 22, 22, 22, 20])
for i, m in enumerate(R["clinique_mois"], 5):
    val(ws, i, 1, m["mois"], bold=True)
    val(ws, i, 2, f"OD {m['od']}")
    val(ws, i, 3, round(m["provision"]), FCFA)
    val(ws, i, 4, f"SIN {m['sin']}")
    val(ws, i, 5, round(m["facture"]), FCFA)
    val(ws, i, 6, f"=C{i}-E{i}", FCFA)
    val(ws, i, 7, "Non")
val(ws, 8, 1, "Total", bold=True)
val(ws, 8, 2, "")
val(ws, 8, 3, "=SUM(C5:C7)", FCFA, bold=True)
val(ws, 8, 4, "")
val(ws, 8, 5, "=SUM(E5:E7)", FCFA, bold=True)
val(ws, 8, 6, "=C8-E8", FCFA, bold=True)
val(ws, 8, 7, "Non")
memo(ws.cell(5, 3), "Le 31/12/2025, la facture d'octobre de la clinique n'était pas encore arrivée. Le Fonds a donc réservé 85 000 000 FCFA pour la payer plus tard. C'est comme mettre de l'argent dans une enveloppe avec écrit dessus « facture d'octobre ». C'est la bonne façon de faire.")
memo(ws.cell(5, 5), "Le 31/01/2026, la vraie facture d'octobre arrive : 77 478 133 FCFA. Le Fonds l'enregistre comme une nouvelle dépense de 2026. Mais il oublie de vider l'enveloppe de 85 000 000 FCFA. La facture d'octobre est donc comptée deux fois : une fois en 2025 (l'enveloppe) et une fois en 2026 (la facture).", h=170)
memo(ws.cell(8, 6), "L'enveloppe contenait un peu plus que les vraies factures : 14 528 881 FCFA de trop. C'est normal pour une estimation. Ce petit surplus devrait revenir dans les comptes de 2026.")
memo(ws.cell(5, 7), "On a cherché dans tout le grand livre 2026 une ligne qui vide l'enveloppe (une écriture qui annule la provision). Il n'y en a aucune : aucun crédit sur les comptes de soins 60 en 2026, et aucune ligne « PROV » sur le compte 40110000.")

r = 10
ws.cell(r, 1, "Ce que cela change en 2026 (au 31/08/2026)").font = fn(bold=True, size=12)
entete(ws, r + 1, ["Chiffre", "Comptes tels quels", "Correction", "Après correction", "Source", "", ""])
t = T.loc[2026]
lig = [("Prestations maladie (601 + 60800007)", round(t.prest_maladie), f"=-C8", "Grand livre 2026 et balance 2026"),
       ("Cotisations maladie (7010)", round(t.cot_maladie), 0, BAL26),
       ("Ratio maladie", "=B12/B13", None, "Calcul Adenka"),
       ("Toutes prestations (60)", round(t.prest_toutes), "=-C8", "Grand livre 2026"),
       ("Toutes cotisations (70)", round(t.cot_toutes), 0, BAL26),
       ("Ratio global (norme 70 % à 80 %)", "=B15/B16", None, "Calcul Adenka"),
       ("Factures des cliniques à payer, compte 40110000", 690234451, "=-C8", BAL26 + ", solde créditeur")]
for i, (a, b_, c, s) in enumerate(lig, r + 2):
    val(ws, i, 1, a, bold=True, wrap=True)
    pct_ = isinstance(b_, str) and b_.startswith("=B")
    val(ws, i, 2, b_, PCT if pct_ else FCFA)
    if pct_:
        val(ws, i, 3, "")
        val(ws, i, 4, f"=D{i - 2}/D{i - 1}", PCT, bold=True, fill="FBE5CB")
    else:
        val(ws, i, 3, c, FCFA)
        val(ws, i, 4, f"=B{i}+C{i}", FCFA, fill="FBE5CB")
    val(ws, i, 5, s, wrap=True)
    ws.row_dimensions[i].height = 32
ws.column_dimensions["E"].width = 60
memo(ws.cell(14, 4), "Sans les factures de 2025 comptées une deuxième fois, la maladie consomme 82,31 FCFA pour 100 FCFA de cotisations sur les 8 premiers mois de 2026, et non 97,40.")
memo(ws.cell(18, 1), "Ce compte dit combien le Fonds doit encore aux cliniques. Comme la facture et l'enveloppe y sont toutes les deux, la dette est comptée deux fois. Il faut enlever 255 000 000 FCFA.")
memo(ws.cell(12, 3), "La bonne écriture : en 2026, le Fonds doit vider l'enveloppe de fin 2025 (255 000 000 FCFA en moins sur les soins de 2026), puisque la vraie facture est déjà passée.")
r = 20
note(ws, r, "Correction à faire par le Fonds : annuler en 2026 la provision de 255 000 000 FCFA (une écriture qui fait l'inverse des OD 805, 806 et 807). Le résultat 2026 gagne alors 255 000 000 FCFA, dont 14 528 881 FCFA de provision faite en trop en 2025.", ncol=7)
note(ws, r + 1, "À confirmer : le lettrage du compte 40110000 (il relie chaque facture à son paiement) doit montrer que les factures SIN 462, 552 et 619 et les OD 805, 806 et 807 désignent bien les mêmes soins.", ncol=7, couleur=JAUNE)
r += 3
ws.cell(r, 1, "Factures mensuelles de la clinique passées en 2026 (hors élèves policiers)").font = fn(bold=True, size=12)
entete(ws, r + 1, ["Date", "Pièce", "Libellé écrit par le Fonds", "N° facture", "Montant", "Mois de soins lu dans le libellé", "Ligne dans le fichier"])
MO = {"462": "Octobre 2025", "552": "Novembre 2025", "619": "Décembre 2025", "962": "Janvier 2026", "1651": "Août 2026"}
for i, x in enumerate(C26.itertuples(), r + 2):
    val(ws, i, 1, x.date, "dd/mm/yyyy")
    val(ws, i, 2, f"SIN {x.piece}")
    val(ws, i, 3, x.libelle, wrap=True)
    val(ws, i, 4, x.facture)
    val(ws, i, 5, round(x.credit), FCFA)
    m = MO.get(x.piece, "non écrit")
    val(ws, i, 6, m, fill=None if m != "non écrit" else JAUNE)
    val(ws, i, 7, x.ligne)
fin = r + 2 + len(C26)
memo(ws.cell(r + 1, 6), "Pour 8 mois de soins en 2026 (janvier à août), on ne trouve que 7 factures mensuelles de la clinique. Cinq factures n'ont pas de mois écrit dans leur libellé. Il manque peut-être une facture (février ?). Si elle n'est pas arrivée, le Fonds devrait mettre de côté de l'argent pour elle au 31/08/2026.", h=170)
note(ws, fin + 1, "À confirmer : 7 factures mensuelles de la clinique pour 8 mois de 2026. Si une facture de 2026 manque, il faut une provision au 31/08/2026 (à demander au Fonds).", ncol=7, couleur=JAUNE)
note(ws, fin + 2, f"Sources : « {GL25} » (OD 805, 806 et 807 du 31/12/2025, comptes 60120000 et 40110000). « {GL26} » (pièces SIN 462, 552 et 619 du 31/01/2026, comptes 601... et 40110000). « {BAL26} » (compte 40110000).", ncol=7)

# ====================================================================== 2 Écritures
ws = wb.create_sheet("2_Ecritures")
titre(ws, "Les écritures du Fonds qui prouvent le double comptage", "Toutes les lignes des grands livres pour les OD 805 à 807 (2025) et les pièces SIN 462, 552 et 619 (2026)")
cols = ["Année", "Date", "Journal", "Pièce", "Compte", "Libellé", "N° facture", "Débit", "Crédit", "Fichier", "Ligne dans le fichier"]
entete(ws, 4, cols, [8, 12, 9, 9, 11, 40, 20, 16, 16, 60, 10])
E = pd.concat([OD, SIN]).sort_values(["an", "piece", "compte"])
for i, x in enumerate(E.itertuples(), 5):
    for j, v in enumerate([x.an, x.date, x.journal, x.piece, x.compte, x.libelle, x.facture, round(x.debit), round(x.credit), x.fichier, x.ligne], 1):
        c = ws.cell(i, j, v)
        c.font = fn(size=10)
        c.border = BORD
        if j in (8, 9):
            c.number_format = FCFA
        if j == 2:
            c.number_format = "dd/mm/yyyy"
memo(ws.cell(4, 8), "Débit sur un compte 60 : le Fonds dit « j'ai une dépense de soins ». Crédit sur le compte 40110000 : il dit « je dois cet argent à la clinique ». On voit les deux, une fois en 2025 et une fois en 2026, pour les mêmes mois.")
ws.freeze_panes = "A5"

# ====================================================================== 2 Autres factures 2025
ws = wb.create_sheet("2_Autres_2025")
titre(ws, "Les autres factures de 2025 passées en 2026 sont petites : au moins 595 512 FCFA",
      "Factures d'autres prestataires de soins enregistrées en 2026, dont le libellé ou le numéro montre l'année 2025 (FCFA)")
cols = ["Date", "Journal", "Pièce", "Compte", "Libellé", "N° facture", "Montant", "Lien avec 2025", "Ligne dans le fichier"]
entete(ws, 4, cols, [12, 9, 9, 11, 42, 22, 16, 44, 10])
r = 5
for df, lien in ((AUT, "Soins de 2025 seulement"), (AUTM, "Mélange 2025 et 2026 (part de 2025 inconnue)")):
    for x in df.itertuples():
        for j, v in enumerate([x.date, x.journal, x.piece, x.compte, x.libelle, x.facture, round(x.montant), lien, x.ligne], 1):
            c = val(ws, r, j, v, FCFA if j == 7 else ("dd/mm/yyyy" if j == 1 else None))
            c.font = fn(size=10)
        r += 1
val(ws, r, 5, "Total soins de 2025 seulement", bold=True)
val(ws, r, 7, f'=SUMIF(H5:H{r - 1},"Soins de 2025 seulement",G5:G{r - 1})', FCFA, bold=True)
val(ws, r + 1, 5, "Total mélange 2025 et 2026", bold=True)
val(ws, r + 1, 7, f'=SUMIF(H5:H{r - 1},"Mélange*",G5:G{r - 1})', FCFA, bold=True)
memo(ws.cell(r, 7), "Ce sont des soins faits en 2025 dont la facture est arrivée en 2026. C'est normal qu'il y en ait un peu. Ici c'est très peu : moins de 1 200 000 FCFA. Comme beaucoup de libellés n'écrivent pas l'année, c'est un minimum.")
note(ws, r + 3, "Comment on les a trouvées : dans le grand livre 2026, comptes de soins 60, on garde les lignes dont le libellé ou le numéro de facture contient 2025, /25 ou -25. On enlève les numéros de chèque (CH...) et les factures de la clinique déjà traitées (SIN 462, 552 et 619).", ncol=9)
note(ws, r + 4, f"Source : « {GL26} ».", ncol=9)
note(ws, r + 5, "Limite : beaucoup de factures n'écrivent pas l'année des soins. Le vrai montant des factures de 2025 arrivées en 2026 ne peut être connu qu'avec la base des prestations (date des soins de chaque facture), demandée au Fonds.", ncol=9, couleur=JAUNE)

# ====================================================================== 3 Ensemble 2026
ws = wb.create_sheet("3_Ensemble_2026")
titre(ws, "Avec les deux corrections, la maladie consomme 59,86 % de ses cotisations sur 8 mois de 2026, et non 97,40 %",
      "Ratios de 2026 (01/01 au 31/08/2026) étape par étape")
entete(ws, 4, ["Étape", "Prestations maladie", "Cotisations maladie", "Ratio maladie", "Toutes prestations", "Toutes cotisations", "Ratio global"], [46, 20, 20, 14, 20, 20, 14])
et = [("1. Comptes tels quels", round(t.prest_maladie), round(t.prest_toutes)),
      ("2. Sans les factures de 2025 comptées deux fois (anomalie 2)", "=B5-'2_Clinique_mois'!C8", "=E5-'2_Clinique_mois'!C8"),
      ("3. Et sans les dépenses hors soins (anomalie 1)", f"=B6-'1_Avant_apres'!D{r26}", f"=E6-'1_Avant_apres'!L{r26}")]
for i, (a, pm, pt) in enumerate(et, 5):
    val(ws, i, 1, a, bold=True, wrap=True)
    val(ws, i, 2, pm, FCFA)
    val(ws, i, 3, round(t.cot_maladie), FCFA)
    val(ws, i, 4, f"=B{i}/C{i}", PCT, bold=True, fill="FBE5CB" if i > 5 else None)
    val(ws, i, 5, pt, FCFA)
    val(ws, i, 6, round(t.cot_toutes), FCFA)
    val(ws, i, 7, f"=E{i}/F{i}", PCT, bold=True, fill="FBE5CB" if i > 5 else None)
    ws.row_dimensions[i].height = 32
memo(ws.cell(7, 4), "Après les deux corrections, il reste 59,86 FCFA de vrais soins pour 100 FCFA de cotisations maladie sur les 8 premiers mois de 2026. Les comptes tels quels disent 97,40 FCFA.")
memo(ws.cell(7, 7), "Le ratio global passe sous la fourchette de 70 % à 80 % du texte : 48,68 %. Les comptes tels quels disaient 75,20 %.")
note(ws, 9, "Les étapes se suivent : on enlève d'abord les 255 000 000 FCFA comptés deux fois, puis les dépenses hors soins de 2026 (379 440 128 FCFA). Les deux corrections ne touchent pas les mêmes lignes.", ncol=7)
note(ws, 10, "2026 n'a que 8 mois. Les cotisations et les soins d'une année entière ne sont pas connus. Ne pas comparer 2026 à 2025 sans le dire.", ncol=7)

# ====================================================================== Sources
ws = wb.create_sheet("Sources")
titre(ws, "D'où viennent les chiffres", "Fichiers du Fonds utilisés et façon d'obtenir chaque chiffre")
entete(ws, 4, ["Chiffre", "Fichier exact du Fonds", "Où dans le fichier", "Comment on l'a obtenu"], [34, 60, 34, 70])
S = [("Prestations par année (2021 à 2026)", "; ".join([]) or "", "", "")]
S = []
for a in sorted(GL):
    S.append((f"Prestations {AN(int(a))}", GL[a], "Comptes 60... (toutes les lignes, sauf les à-nouveaux RAN)",
              "Lecture de chaque ligne du grand livre. Le total de chaque compte est égal au mouvement de la balance de la même année."))
for a in sorted(BAL):
    S.append((f"Cotisations {AN(int(a))}", BAL[a], "Comptes 7010 (maladie) et 70 (toutes)", "Mouvement crédit moins mouvement débit de l'année."))
S += [("Dépenses hors soins", "Grands livres ci-dessus", "Comptes 6011, 6012, 6013, 6018 et 6080, colonne « Libellé écriture »",
       "Classement de chaque ligne par les mots de son libellé (règles du rapport, onglet 1_Par_type). Variante élargie avec d'autres mots."),
      ("Provision de la clinique fin 2025", GL25, "OD 805, 806 et 807 du 31/12/2025, comptes 60120000 et 40110000", "Somme des trois écritures : 3 fois 85 000 000 FCFA."),
      ("Vraies factures de la clinique (octobre à décembre 2025)", GL26, "Pièces SIN 462, 552 et 619 du 31/01/2026, comptes 601... et 40110000", "Somme des lignes de chaque pièce. Débit des comptes 60 égal au crédit du compte 40110000."),
      ("Absence d'annulation de la provision", GL26, "Comptes 60 et 40110000, tout 2026", "Aucun crédit sur un compte 60 et aucune ligne « PROV » sur le compte 40110000."),
      ("Dette envers les cliniques au 31/08/2026", BAL26, "Compte 40110000, solde créditeur", "690 234 451 FCFA lus tels quels."),
      ("Autres factures de 2025 passées en 2026", GL26, "Comptes 60, colonnes « Libellé écriture » et « N° facture »", "Lignes dont le libellé ou le numéro porte l'année 2025 (onglet 2_Autres_2025)."),
      ("Règles de sinistralité", "Décision n° 05/2022/COM/UEMOA du 1er juin 2022", "Annexe, § 1.2.1.2.5", "Ratio global = prestations / cotisations, norme 70 % à 80 %.")]
for i, row in enumerate(S, 5):
    for j, v in enumerate(row, 1):
        val(ws, i, j, v, wrap=True).font = fn(size=10)
    ws.row_dimensions[i].height = 42
note(ws, 6 + len(S), "Tous les calculs sont faits par Adenka en Python à partir de ces fichiers, puis repris ici avec des formules Excel. Le Fonds ne calcule pas ces ratios corrigés : ce sont des estimations Adenka.", ncol=4)

for s in wb.worksheets:
    s.sheet_properties.tabColor = ORANGE if s.title[0] in "123" else NOIR
wb.save(OUT)
print("écrit", OUT)
