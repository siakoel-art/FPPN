# -*- coding: utf-8 -*-
"""
FPPN - Mise à jour de la PBC pointée, de la liste des demandes au FPPN et du plan de travail
après la dataroom du 09/11/2026 (lots A et B). Tout ajout ou changement est écrit en bleu.
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
J = dt.datetime(2026, 11, 9)


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


# ------------------------------------------------------------------ PBC pointée
MAJ = [  # début du texte de l'item, nouveau statut (None : inchangé), commentaire ajouté
    ("Tout rapport actuariel, étude technique ou note de calcul", None, "Reçu le 09/11/2026 : fichiers de calcul de la provision pour congés et retraite 2020 à 2025 (droits acquis). Toujours attendu : étude actuarielle des risques longs."),
    ("Barèmes des capitaux décès, frais funéraires", "Reçu", "Reçu le 09/11/2026 : « OFFRES ET BAREMES DES ALLOCATIONS ; CAPITAUX DECES ET FRAIS FUNERAIRES.docx » (barème 2001 à 2026, frais funéraires de 300 000 FCFA depuis 2024)."),
    ("Rapport de contrôle de l'AIRMS et plan de redressement", None, "Reçus le 09/11/2026 : 28 courriers de l'AIRMS (2022 à 2026), dont la transmission du rapport définitif de la mission de contrôle d'août 2025. Le rapport lui-même n'est pas joint."),
    ("Rapports du commissaire aux comptes 2020 à 2025", None, "Reçus le 09/11/2026 : rapports du commissaire aux comptes sur les exercices 2022 à 2025 et rapport au CA sur l'exercice 2024. Manquent 2020 et 2021."),
    ("Procès-verbaux des réunions du conseil d'administration", None, "Reçus le 09/11/2026 : PV des AG et des CA de 2020 à 2026 (28 documents) et rapports du CA. Manquent les rapports de la Commission de surveillance et de contrôle."),
    ("Bulletin de souscription au FCP géré par United Capital", "Partiellement Reçu", "Reçus le 09/11/2026 : bulletin de souscription, convention, formulaires KYC et spécimen de signature. Manquent les valeurs liquidatives et les relevés."),
    ("DAT BOA (comptes 52210200 et 52210300)", None, "Reçus le 09/11/2026 : courrier d'ouverture d'un DAT de 60 000 000 FCFA (10/01/2026, 5 % demandé) et notification de crédit BOA de 500 000 000 FCFA (2021)."),
    ("Titres de propriété ou actes d'acquisition des terrains", None, "Reçus le 09/11/2026 : protocole SONEB (terrain de Songon remplacé par Jacqueville). Le terrain d'Attoban n'a pas de titre (rapports du commissaire aux comptes)."),
    ("Dossier de la créance sur l'hôpital de la Police", "Partiellement Reçu", "Reçus le 09/11/2026 : rapport du commissaire aux comptes (dépenses pouvant être des subventions), PV du CA du 06/05/2025, courrier de l'AIRMS du 05/01/2026 (créance de 3 932 124 730 FCFA sur l'État)."),
    ("Convention d'apporteur d'affaires avec GNA", "Partiellement Reçu", "Reçue le 09/11/2026 : « CONVENTION GNA- FPPN.pdf ». Manquent les décomptes annuels des commissions."),
    ("Version des comptes 2023 approuvée par l'Assemblée générale", "Reçu", "Reçu le 09/11/2026 : PV de l'AGO du 26/06/2024 (excédent 2023 de 930 629 102 FCFA, identique à la balance) et rapport du commissaire aux comptes sur 2023."),
    ("Décision de désignation ou de prolongation du directeur général", None, "Reçus le 09/11/2026 : PV de 2024 à 2026 montrant la directrice générale en fonction. La décision de désignation reste attendue."),
    ("Comptabilité séparée du Centre médical du FPPN", None, "Reçu le 09/11/2026 : compte de résultat 2024 du CMP (PV du CA du 02/06/2025 : apport du FPPN de 187 916 783 FCFA)."),
    ("Texte qui fonde l'exclusion du suicide", None, "Reçu le 09/11/2026 : barème officiel. Le texte sur l'exclusion du suicide reste attendu."),
    ("Liste des prestataires de santé conventionnés", None, "Reçues le 09/11/2026 : conventions MUGEF-CI, MAABS, AMAT-CI, GIEPHARM, Transvie, DEPA, ministère de la Santé, Service de santé, CDCMGM."),
    ("Fichier du personnel de la Direction Générale", None, "Reçus le 09/11/2026 : fichiers du personnel civil 2020 à 2025 (naissance, embauche, salaire) des calculs de provision."),
    ("Décision d'agrément du FPPN", None, "Reçu le 09/11/2026 : réponse de l'AIRMS à la demande de copie de l'agrément, avec l'arrêté."),
    ("Fichiers retour maladie (936) et décès (8Z6) manquants ou vides", None, "Envoi du 09/11/2026 : les fichiers de décembre 2024 restent illisibles."),
]
NOUV = [  # thème de section, texte, priorité
    ("Jugement et objet du procès perdu sur l'EFC (263 000 000 FCFA bloqués à la BOA, CA du 12/08/2025).", "Urgent"),
    ("Étude de faisabilité, plan de financement et décisions des organes sur le projet d'Azaguié (19 539 377 111 FCFA).", "Urgent"),
    ("Dossier de la créance de 3 932 124 730 FCFA sur l'État (policiers blessés en service) et lien avec la créance sur l'hôpital de la Police.", "Urgent"),
    ("Rapprochement de l'épargne PEP et EFC entre les livres du FPPN et ceux du Grenier (8 493 055 501 FCFA au passif du Grenier).", "Urgent"),
    ("Détail du calcul des ratios prudentiels par le FPPN : réserve obligatoire (9 mois en 2025) et retraitement des ressources stables.", "Moyen"),
    ("Rapport définitif de la mission de contrôle de l'AIRMS du 19 au 22/08/2025 et réponse du FPPN.", "Urgent"),
    ("Décision du CA sur le terrain d'Attoban (375 000 000 FCFA) et sur la créance sur l'hôpital de la Police (3 407 519 850 FCFA).", "Urgent"),
]

src = os.path.join(REP, "FPPN_PBC_complete_pointee_20261026.xlsx")
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
ws.cell(r, 3, "Documents reçus le 09/11/2026 : nouvelles demandes")
bleu(ws.cell(r, 3))
prec = der
for txt, prio in NOUV:
    r += 1
    for j in range(1, 15):
        ws.cell(r, j)._style = copy(ws.cell(modele_it, j)._style)
    ws.cell(r, 2, f"=+B{prec}+1")
    ws.cell(r, 3, txt)
    ws.cell(r, 5, prio)
    ws.cell(r, 7, dt.datetime(2026, 11, 10))
    ws.cell(r, 8, dt.datetime(2026, 11, 17))
    ws.cell(r, 11, "Demande issue des documents reçus le 09/11/2026")
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
ajout_texte(ws.cell(6, 3), "Pointage complété le 10/11/2026 avec les documents de la dataroom du 09/11/2026 (en bleu).")
wb.save(os.path.join(REP, "FPPN_PBC_complete_pointee_20261110.xlsx"))
print("PBC :", len(MAJ), "items mis à jour,", len(NOUV), "nouveaux items, dernière ligne", nouveau_der)

# ------------------------------------------------------------------ demandes au FPPN
wb = openpyxl.load_workbook(os.path.join(REP, "FPPN_Demandes_au_FPPN_20261023.xlsx"))
wq = wb["1_Questions"]
derq = max(r for r in range(5, wq.max_row + 1) if str(wq.cell(r, 1).value or "").isdigit())
REPONSES = {
    "Quelle version des comptes 2023": "Réponse reçue le 09/11/2026 : comptes 2023 approuvés par l'AGO du 26/06/2024 (excédent de 930 629 102 FCFA).",
    "Quelle est la nature de la créance": "Élément reçu le 09/11/2026 : selon le commissaire aux comptes, ce seraient des subventions ; le CA du 06/05/2025 a confié le dossier au PCA.",
    "À quel grade correspond l'allocation de 2 200 000": "Réponse reçue le 09/11/2026 : commissaire de police (barème officiel).",
    "Le directeur général a-t-il été désigné": "Élément reçu le 09/11/2026 : les PV de 2024 à 2026 la présentent en fonction ; la décision reste attendue.",
}
for r in range(5, derq + 1):
    q = str(wq.cell(r, 3).value or "")
    for deb, rep in REPONSES.items():
        if q.startswith(deb):
            ajout_texte(wq.cell(r, 6), rep)
            wq.cell(r, 8).value = J
            bleu(wq.cell(r, 8))
QN = [("Épargne", "Quel est l'objet du procès perdu sur l'EFC, et pourquoi 263 000 000 FCFA sont-ils bloqués à la BOA ?", "Urgent"),
      ("Placements", "Quels sont l'étude de faisabilité et le plan de financement du projet d'Azaguié (19 539 377 111 FCFA) ?", "Urgent"),
      ("Placements", "Quel est le dossier de la créance de 3 932 124 730 FCFA sur l'État, et quel est son lien avec la créance sur l'hôpital ?", "Urgent"),
      ("Grenier", "Qui est propriétaire de l'épargne PEP et EFC inscrite à la fois dans les livres du FPPN et dans ceux du Grenier ?", "Urgent"),
      ("Ratios", "Comment le FPPN calcule-t-il 9 mois de réserve obligatoire en 2025 et retraite-t-il les ressources stables ?", "Moyen"),
      ("Conformité", "Le rapport définitif de la mission de contrôle de l'AIRMS (19 au 22/08/2025) peut-il être transmis ?", "Urgent")]
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
wb.save(os.path.join(REP, "FPPN_Demandes_au_FPPN_20261110.xlsx"))
print("Demandes :", len(QN), "nouvelles questions,", n_maj, "documents mis à jour")

# ------------------------------------------------------------------ plan de travail
wb = openpyxl.load_workbook(os.path.join(REP, "FPPN_Plan_de_travail_Actuariat_20261023.xlsx"))
wp = wb["PT_Actuariat"]
OBS = {
    "1.2": "10/11/2026 : PBC pointée avec la dataroom du 09/11/2026 (FPPN_PBC_complete_pointee_20261110.xlsx).",
    "2.1": "09/11/2026 : inventaire de la dataroom (5 366 fichiers comparés par empreinte), 980 fichiers transmis et contrôlés.",
    "2.5": "Balances 2020 à 2025 retéléchargées le 09/11/2026 : soldes identiques ; balance 2026 refaite (quatre reclassements).",
    "4.4": "Provision du personnel : méthode du FPPN (droits acquis) confirmée par ses fichiers de calcul 2020 à 2025.",
    "5.1": "Actifs dont la valeur n'est pas établie (terrains sans titre, créances) et engagements hors bilan intégrés au rapport du 10/11/2026.",
    "6.1": "Ratios calculés par le FPPN (2022 à 2026) comparés au recalcul du rapport (rapport du 10/11/2026).",
    "8.1": "Rapport du 10/11/2026 : documents du 09/11/2026 et commentaires intégrés, ajouts en bleu.",
}
for r in range(7, 57):
    n = str(wp.cell(r, 1).value or "")
    if n in OBS:
        ajout_texte(wp.cell(r, 6), OBS[n])
ajout_texte(wp.cell(2, 1), "Mise à jour du 10/11/2026 en bleu (documents du 09/11/2026).")
wb.save(os.path.join(REP, "FPPN_Plan_de_travail_Actuariat_20261110.xlsx"))
print("Plan de travail : observations ajoutées pour", len(OBS), "activités")
