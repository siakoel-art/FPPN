# -*- coding: utf-8 -*-
"""
FPPN - Préparation de l'envoi des nouveaux fichiers reçus le 22/10/2026, par lots de priorité
=============================================================================================
Copie les fichiers tels quels (aucune conversion, aucune modification) dans des archives ZIP découpées en parties
de 20 Mo, une archive par lot, avec un journal (chemin, taille, empreinte SHA-256) pour contrôler la réception.
Lot 1 : documents actuariels prioritaires (environ 45 Mo). Lot 2 : états financiers, rapports de gestion,
conventions et titres (environ 85 Mo). Lot 3 : budgets, PV de caisse, organigrammes (environ 60 Mo).
Utilisation : python envoi_nouveaux_fppn.py   (choisir les lots dans LOTS_A_PREPARER)
"""
# ============================================================================= PARAMÈTRES
RACINE = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client"
SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses\envoi_20261022"
TAILLE_PARTIE_MO = 20
LOTS_A_PREPARER = [1]          # puis [2], puis [3] après accord
LOTS = {
    1: [  # (dossier relatif, motif contenu dans le nom du fichier) ; motif vide = tout le dossier
        (r"03-Nouveaux Juridiques", "CAHIER DE CHARGE"), (r"03-Nouveaux Juridiques", "MANUEL DE PRODUIT D'EPARGNE"),
        (r"03-Nouveaux Juridiques", "POLITIK DE CREDIT"), (r"03-Nouveaux Juridiques", "NUMERO D'IMMATRICULATION"),
        (r"03-Nouveaux Juridiques", "PV AG ORDINAIRE"), (r"03-Nouveaux Juridiques", "LISTE DES COMPTES BANCAIRES"),
        (r"03-Nouveaux Juridiques", "Habilitation de signatures"),
        (r"05-Nouveaux Prestations", ""),
        (r"04-Nouveaux Personnel et paie", ".xls"), (r"04-Nouveaux Personnel et paie", ".xlsx"),  # DISA, ETAT 301, DASC
        (r"02-Nouveaux_GL et Balances", ""),
        (r"07-Nouveaux Trésorerie et placements", ".xlsx"),  # états de rapprochement bancaire 2020 à 2026
    ],
    2: [
        (r"03-Nouveaux Juridiques", "ETATS FINANCIERS"), (r"03-Nouveaux Juridiques", "RAPPORT"),
        (r"03-Nouveaux Juridiques", "BILAN_DE_L'EXERCICE"), (r"03-Nouveaux Juridiques", "Plan d'action"),
        (r"03-Nouveaux Juridiques", "PLAN D'ACTION"), (r"03-Nouveaux Juridiques", "CONVENTION"),
        (r"03-Nouveaux Juridiques", "Certificat de Propriété"), (r"03-Nouveaux Juridiques", "Vente ferme"),
        (r"03-Nouveaux Juridiques", "CONTRAT DE BAIL"), (r"03-Nouveaux Juridiques", "LISTE DES ADMINISTRATEURS"),
    ],
    3: [
        (r"06-Nouveaux Revue Budgétaire", ""), (r"07-Nouveaux Trésorerie et placements", ".pdf"),
        (r"03-Nouveaux Juridiques", "Organigramme"), (".", "COURRIER DE NOMINATION"), (".", "ORGANIGRAMME FPPN 2021"),
        (r"04-Nouveaux Personnel et paie", ".pdf"),
    ],
}
# =============================================================================
import csv
import hashlib
import os
import zipfile


def empreinte(chemin):
    h = hashlib.sha256()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1 << 20), b""):
            h.update(bloc)
    return h.hexdigest()


def fichiers_du_lot(regles):
    vus = []
    for dossier, motif in regles:
        base = os.path.join(RACINE, dossier)
        if dossier == ".":
            candidats = [os.path.join(RACINE, n) for n in os.listdir(RACINE) if os.path.isfile(os.path.join(RACINE, n))]
        else:
            candidats = [os.path.join(d, n) for d, _, ns in os.walk(base) for n in ns]
        for c in sorted(candidats):
            nom = os.path.basename(c)
            ok = (not motif) or (motif.startswith(".") and nom.lower().endswith(motif)) or (motif.lower() in nom.lower())
            if ok and c not in vus:
                vus.append(c)
    return vus


os.makedirs(SORTIE, exist_ok=True)
for lot in LOTS_A_PREPARER:
    liste = fichiers_du_lot(LOTS[lot])
    archive = os.path.join(SORTIE, f"FPPN_nouveaux_lot{lot}_20261022.zip")
    with open(os.path.join(SORTIE, f"journal_lot{lot}.csv"), "w", newline="", encoding="utf-8-sig") as j:
        w = csv.writer(j, delimiter=";")
        w.writerow(["chemin_relatif", "taille_octets", "sha256"])
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
            for f in liste:
                rel = os.path.relpath(f, RACINE)
                z.write(f, rel)
                w.writerow([rel, os.path.getsize(f), empreinte(f)])
                print(f"[lot {lot}] {rel}")
        z_journal = os.path.join(SORTIE, f"journal_lot{lot}.csv")
    with zipfile.ZipFile(archive, "a") as z:
        z.write(z_journal, os.path.basename(z_journal))
    taille, n = os.path.getsize(archive), 0
    with open(archive, "rb") as f:
        while True:
            bloc = f.read(int(TAILLE_PARTIE_MO * 1024 * 1024))
            if not bloc:
                break
            n += 1
            with open(f"{archive}.{n:03d}", "wb") as p:
                p.write(bloc)
    os.remove(archive)
    print(f"\nLot {lot} : {len(liste)} fichiers, {taille / 1e6:.1f} Mo, {n} partie(s) dans {SORTIE}")
print("Envoyer toutes les parties .001, .002, ... de chaque lot préparé.")
