# -*- coding: utf-8 -*-
"""
FPPN - Mise à jour de la PBC pointée, de la liste des demandes au FPPN et du plan de travail
après la dataroom du 11/11/2026. Tout ajout ou changement est écrit en bleu.
Utilisation : python maj_pbc_demandes_v14_fppn.py <dossier du dépôt>
"""
import datetime as dt
import os
import sys
from copy import copy

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font

REP = sys.argv[1]
BLEU = "0070C0"
J = dt.datetime(2026, 11, 11)


def bleu(c):
    f = copy(c.font)
    c.font = Font(name=f.name or "Helvetica", size=f.size or 9, bold=f.bold, italic=f.italic, color=BLEU)


def ajout_texte(c, texte):
    """Ajoute un texte en bleu à la fin d'une cellule, sans changer la couleur du texte existant."""
    avant = "" if c.value is None else str(c.value)
    f = c.font
    base = InlineFont(rFont=f.name or "Helvetica", sz=f.size or 9)
    neuf = InlineFont(rFont=f.name or "Helvetica", sz=f.size or 9, color=BLEU)
    sep = " " if avant else ""
    c.value = CellRichText([TextBlock(base, avant + sep), TextBlock(neuf, texte)]) if avant else CellRichText([TextBlock(neuf, texte)])


# ------------------------------------------------------------------ PBC pointée (v16, 11/11/2026)
MAJ = [  # début du texte de l'item, nouveau statut (None : inchangé), commentaire ajouté
    ("Prêts aux mutualistes (« Le Grenier du Policier »)", None, "Reçu le 11/11/2026 : balances et grands livres du Grenier 2021 au 31/08/2026 (encours de 1 044 302 107 FCFA, aucune dépréciation). Balance âgée par emprunteur toujours attendue."),
    ("Rapport de réévaluation du siège", "Reçu", "Reçu le 11/11/2026 : rapport d'expertise n° 005/AKKA/19 du 19/03/2019 (2 899 400 000 FCFA, égal au compte 10620000). Décision de comptabilisation et coût d'origine du terrain toujours attendus."),
    ("Dossier de la créance sur l'hôpital de la Police", None, "Mémo du service comptable reçu le 11/11/2026 : charges et investissements faits pour l'hôpital « avant leur différend ». Pièces toujours attendues."),
    ("Liste des comptes bancaires, actifs et inactifs", None, "Complété le 11/11/2026 : état des comptes ouverts et fermés d'août 2020 au 31/08/2026."),
    ("Fichier du personnel de la Direction Générale", None, "Reçu le 11/11/2026 : déclaration des salaires d'août 2026 (100 salariés, dont 33 sans numéro CNPS). Dates de naissance et d'embauche de ces 33 salariés toujours attendues."),
    ("Rapprochement de l'épargne PEP et EFC entre les livres du FPPN et ceux du Grenier", "Partiellement Reçu", "Balances du Grenier reçues le 11/11/2026 : 8 493 055 501 FCFA figés depuis 2023, contrepartie « fonds en attente de réception ». Rapprochement toujours attendu."),
]
NOUV = [  # texte, priorité
    ("Écriture d'origine du compte 10620000 (2 899 400 000 FCFA), décision des organes et coût d'acquisition du terrain du siège.", "Urgent"),
    ("Dates de naissance, d'embauche et numéros CNPS des 33 salariés déclarés avec le numéro 111111111111 (août 2026).", "Urgent"),
    ("Rapprochement du compte de liaison FPPN (47111102) et Grenier (37921400) : écart de 70 191 500 FCFA au 31/08/2026.", "Urgent"),
    ("Liste des prêts PEP de 2004 à 2015 (27110000 et 27120000, 379 204 261 FCFA) et décision de dépréciation.", "Moyen"),
    ("Factures et contrat du cabinet payé chaque trimestre sur le compte des soins (factures n° 00444…, 275 960 518 FCFA de 2024 au 31/08/2026).", "Urgent"),
    ("Liste et décisions d'octroi des dons imputés aux prestations maladie (203 689 140 FCFA en 2025).", "Moyen"),
    ("Règlement et valeur liquidative du fonds commun de placement ouvert le 04/11/2025 (500 000 000 FCFA).", "Moyen"),
]

src = os.path.join(REP, "FPPN_PBC_complete_pointee_20261110.xlsx")
wb = openpyxl.load_workbook(src)
ws = wb["PBC Actuariat FPPN"]
der = max(r for r in range(11, ws.max_row + 1) if ws.cell(r, 3).value)
faits = set()
NUM_MAJ = {}
_k = 0
for r in range(11, der + 1):
    if ws.cell(r, 2).value:
        _k += 1
    t = str(ws.cell(r, 3).value or "")
    for deb, st, com in MAJ:
        if t.startswith(deb) and deb not in faits:
            faits.add(deb)
            ajout_texte(ws.cell(r, 11), com)
            ws.cell(r, 9).value = J
            bleu(ws.cell(r, 9))
            if st and ws.cell(r, 13).value != st:
                ws.cell(r, 13).value = st
                bleu(ws.cell(r, 13))
            NUM_MAJ[_k] = (ws.cell(r, 13).value, com)
manq = [d for d, _, _ in MAJ if d not in faits]
assert not manq, manq
# section et nouveaux items
modele_sec, modele_it = der - 7, der
while ws.cell(modele_sec, 2).value:
    modele_sec -= 1
r = der + 1
for j in range(1, 15):
    ws.cell(r, j)._style = copy(ws.cell(modele_sec, j)._style)
ws.cell(r, 3, "Documents reçus le 11/11/2026 : nouvelles demandes")
bleu(ws.cell(r, 3))
prec = der
for txt, prio in NOUV:
    r += 1
    for j in range(1, 15):
        ws.cell(r, j)._style = copy(ws.cell(modele_it, j)._style)
    ws.cell(r, 2, f"=+B{prec}+1")
    ws.cell(r, 3, txt)
    ws.cell(r, 5, prio)
    ws.cell(r, 7, dt.datetime(2026, 11, 11))
    ws.cell(r, 8, dt.datetime(2026, 11, 18))
    ws.cell(r, 11, "Demande issue des documents reçus le 11/11/2026")
    ws.cell(r, 13, "Non Reçu")
    for j in (2, 3, 5, 7, 8, 11, 13):
        bleu(ws.cell(r, j))
    prec = r
nouveau_der = r
# formules de synthèse : étendre les plages jusqu'à la nouvelle dernière ligne
for row in ws.iter_rows(min_row=1, max_row=10):
    for c in row:
        if isinstance(c.value, str) and c.value.startswith("="):
            c.value = c.value.replace(f"${der}", f"${nouveau_der}").replace(f":M{der}", f":M{nouveau_der}").replace(f":C{der}", f":C{nouveau_der}")
for cf in list(ws.conditional_formatting):
    pass
ws.cell(5, 3).value = str(ws.cell(5, 3).value or "")
ajout_texte(ws.cell(6, 3), "Pointage complété le 11/11/2026 avec les documents de la dataroom du 11/11/2026 (en bleu).")
wb.save(os.path.join(REP, "FPPN_PBC_complete_pointee_20261111.xlsx"))
print("PBC :", len(MAJ), "items mis à jour,", len(NOUV), "nouveaux items, dernière ligne", nouveau_der)

# ------------------------------------------------------------------ demandes au FPPN
wb = openpyxl.load_workbook(os.path.join(REP, "FPPN_Demandes_au_FPPN_20261110.xlsx"))
wq = wb["1_Questions"]
derq = max(r for r in range(5, wq.max_row + 1) if str(wq.cell(r, 1).value or "").isdigit())
REPONSES = {
    "Le compte 10620000 enregistre-t-il": "Élément reçu le 11/11/2026 : le compte égale au franc près la valeur de l'expertise du siège du 19/03/2019 (2 899 400 000 FCFA). L'écriture d'origine reste attendue.",
    "Qui est propriétaire de l'épargne PEP et EFC": "Élément reçu le 11/11/2026 : au Grenier, cette épargne a pour contrepartie un compte « fonds en attente de réception » et ne bouge plus depuis 2023.",
}
for r in range(5, derq + 1):
    q = str(wq.cell(r, 3).value or "")
    for deb, rep in REPONSES.items():
        if q.startswith(deb):
            ajout_texte(wq.cell(r, 6), rep)
            wq.cell(r, 8).value = J
            bleu(wq.cell(r, 8))
QN = [("Actifs", "Sur quelle décision et pour quel coût d'origine du terrain le compte 10620000 (2 899 400 000 FCFA) a-t-il été passé ?", "Urgent"),
      ("Personnel", "Pourquoi 33 salariés sont-ils déclarés avec le numéro CNPS 111111111111 en août 2026, et sont-ils affiliés ?", "Urgent"),
      ("Grenier", "Comment s'explique l'écart de 70 191 500 FCFA entre les comptes de liaison du FPPN et du Grenier au 31/08/2026 ?", "Urgent"),
      ("Prestations", "Quelle prestation rémunèrent les factures trimestrielles du cabinet payées sur le compte des soins (série 00444…) ?", "Urgent"),
      ("Prestations", "Pourquoi les dons, gadgets et imprimés sont-ils imputés aux prestations maladie plutôt qu'aux frais de gestion ?", "Moyen"),
      ("Placements", "Quel fonds commun de placement a reçu 500 000 000 FCFA d'épargne PEP le 04/11/2025, et quelle est sa valeur liquidative ?", "Moyen")]
wq.insert_rows(derq + 1, len(QN))
for k, (th, q, pr) in enumerate(QN, 1):
    r = derq + k
    for j in range(1, 9):
        wq.cell(r, j)._style = copy(wq.cell(derq, j)._style)
    vals = [str(derq - 4 + k), th, q, "Nouveau", pr]
    for j, v in enumerate(vals, 1):
        wq.cell(r, j, v)
        bleu(wq.cell(r, j))
    wq.row_dimensions[r].height = wq.row_dimensions[derq].height
for t in wq.tables.values():
    a, b_ = t.ref.split(":")
    t.ref = f"{a}:{b_[0]}{int(b_[1:]) + len(QN)}"
# documents : statuts repris de la PBC
wd = wb["2_Documents"]
n_maj = 0
for r in range(5, wd.max_row + 1):
    n = str(wd.cell(r, 2).value or "")
    if n.isdigit() and int(n) in NUM_MAJ:
        st, com = NUM_MAJ[int(n)]
        if wd.cell(r, 5).value != st:
            wd.cell(r, 5).value = st
            bleu(wd.cell(r, 5))
        ajout_texte(wd.cell(r, 8), com)
        n_maj += 1
wb.save(os.path.join(REP, "FPPN_Demandes_au_FPPN_20261111.xlsx"))
print("Demandes :", len(QN), "nouvelles questions,", n_maj, "documents mis à jour")

# ------------------------------------------------------------------ plan de travail
wb = openpyxl.load_workbook(os.path.join(REP, "FPPN_Plan_de_travail_Actuariat_20261110.xlsx"))
wp = wb["PT_Actuariat"]
OBS = {
    "1.2": "11/11/2026 : PBC pointée avec la dataroom du 11/11/2026 (FPPN_PBC_complete_pointee_20261111.xlsx).",
    "2.1": "11/11/2026 : inventaire de la dataroom (263 fichiers comparés par empreinte), 32 fichiers nouveaux lus en entier.",
    "4.4": "Provision du personnel : 33 salariés sans numéro CNPS en août 2026, hors du fichier de provision.",
    "5.1": "Réévaluation du siège (10620000), mémo des comptes 47 et 27 et liaison avec le Grenier intégrés au rapport du 11/11/2026.",
    "8.1": "Rapport du 11/11/2026 : sinistralité lue dans les deux sens et documents du 11/11/2026, ajouts en bleu.",
}
for r in range(7, 57):
    n = str(wp.cell(r, 1).value or "")
    if n in OBS:
        ajout_texte(wp.cell(r, 6), OBS[n])
ajout_texte(wp.cell(2, 1), "Mise à jour du 11/11/2026 en bleu (sinistralité et documents du 11/11/2026).")
wb.save(os.path.join(REP, "FPPN_Plan_de_travail_Actuariat_20261111.xlsx"))
print("Plan de travail : observations ajoutées pour", len(OBS), "activités")
