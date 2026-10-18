# -*- coding: utf-8 -*-
"""
FPPN - Extraction des données reçues du client, pour transmission sans perte
============================================================================

Ce script :
  1. vérifie et installe les bibliothèques Python nécessaires ;
  2. parcourt tout le dossier des documents reçus (sous-dossiers et fichiers ZIP compris) ;
  3. convertit chaque feuille de chaque fichier Excel (.xlsx, .xls) en un fichier CSV brut :
     toutes les lignes et toutes les cellules, en-têtes compris, sans tri ni filtre ;
  4. copie tels quels les fichiers texte (CSV, TXT, .xls au format page web, .doc) et les fichiers
     qu'il ne sait pas lire : rien n'est modifié, rien n'est perdu ;
  5. laisse de côté les fichiers hors analyse (scans de pièces individuelles, PDF en double
     d'un fichier Excel), en les listant dans le journal ;
  6. écrit un journal de contrôle : pour chaque fichier, empreinte SHA-256, action, nombre de
     feuilles, de lignes et de cellules, somme des nombres ;
  7. regroupe le tout dans une archive ZIP, découpée en parties de taille fixe pour l'envoi.

Utilisation : ouvrir ce fichier dans VS Code, vérifier la partie PARAMÈTRES, puis lancer
(bouton « Run Python File » en haut à droite, ou dans le terminal : python extraction_fppn.py).

Le script ne modifie ni ne déplace aucun fichier du dossier source. Il n'envoie rien sur Internet,
sauf pour installer les bibliothèques manquantes (pip).
"""

# =============================================================================
# PARAMÈTRES (à vérifier avant de lancer)
# =============================================================================
DOSSIER_SOURCE = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client\01-download"
DOSSIER_SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses"

TAILLE_PARTIE_MO = 20        # taille de chaque partie à envoyer (Mo)
TAILLE_COPIE_MAX_MO = 50     # au-delà, un fichier « à copier » est listé sans être copié
SIMULATION = False           # True : calcule seulement le volume, sans rien écrire

# Dossiers laissés de côté : scans de pièces individuelles, sans données chiffrées exploitables.
# Leur nombre de fichiers et leur volume figurent quand même dans le journal.
DOSSIERS_EXCLUS = [
    r"PIECES A FOURNIR",
    r"CONTRATS",
    r"EFC\TEST",
    r"DOUBLE PRECOMPTE",
    r"PRECOMPTES INDUS",
    r"RESILIATIONS",
    r"BORDEREAUX",
    r"TRANSMISSION",
]

# PDF transmis malgré tout : leur version Excel est endommagée, ou ils n'existent qu'en PDF.
PDF_A_TRANSMETTRE = [
    r"DAFC\2021\Balance_des_tiers FPPN 2021.pdf",
    r"DAFC\2023\Grand-livre_des_tiers FPPN 2023.pdf",
    r"DAFC\2025\Balance_des_tiers FPPN 2025.pdf",
    r"SERVICE INFORMATIQUE\FPPN_reponse_Audit_SI.pdf",
    r"SERVICE INFORMATIQUE\RAPPORT AUDIT INFORMATIQUE DU FPPN 2022.pdf",
]

# =============================================================================
# 1. INSTALLATION DES BIBLIOTHÈQUES MANQUANTES
# =============================================================================
import importlib
import subprocess
import sys

BIBLIOTHEQUES = {          # nom à importer : nom à installer avec pip
    "python_calamine": "python-calamine",
    "openpyxl": "openpyxl",
}


def installer_bibliotheques():
    manquantes = []
    for module, paquet in BIBLIOTHEQUES.items():
        try:
            importlib.import_module(module)
        except ImportError:
            manquantes.append(paquet)
    if manquantes:
        print(f"Installation des bibliothèques manquantes : {', '.join(manquantes)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *manquantes])
        importlib.invalidate_caches()


installer_bibliotheques()

# =============================================================================
# 2. IMPORTS
# =============================================================================
import csv
import hashlib
import io
import math
import os
import re
import shutil
import time
import unicodedata
import warnings
import zipfile
from datetime import date, datetime, time as dtime, timedelta

import openpyxl
from python_calamine import CalamineWorkbook

warnings.filterwarnings("ignore")
csv.field_size_limit(2**31 - 1)

# =============================================================================
# 3. OUTILS
# =============================================================================
EXT_TABLEUR = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}
EXT_A_COPIER = {".csv", ".txt", ".doc", ".docx", ".rtf", ".xml", ".json"}


def sha256(donnees):
    return hashlib.sha256(donnees).hexdigest()


def nom_sur(texte, longueur=60):
    """Nom de fichier sans caractères interdits sous Windows."""
    texte = unicodedata.normalize("NFKD", texte).encode("ascii", "ignore").decode()
    texte = re.sub(r"[^A-Za-z0-9._-]+", "_", texte).strip("_")
    return texte[:longueur] or "sans_nom"


def normaliser(chemin):
    return chemin.replace("/", "\\").upper()


def est_exclu(chemin_relatif):
    c = normaliser(chemin_relatif)
    return any(("\\" + d.upper() + "\\") in ("\\" + c) for d in DOSSIERS_EXCLUS)


def pdf_a_transmettre(chemin_relatif):
    c = normaliser(chemin_relatif)
    return any(c.endswith(p.upper()) for p in PDF_A_TRANSMETTRE)


def format_tableur(donnees):
    """Reconnaît le vrai format d'un fichier tableur d'après ses premiers octets."""
    debut = donnees[:2048]
    if debut.startswith(b"PK"):
        return "xlsx"
    if debut.startswith(b"\xd0\xcf\x11\xe0"):
        return "xls"
    texte = debut.decode("utf-8", "ignore").lower()
    if "<html" in texte or "<table" in texte or "<?xml" in texte:
        return "page web"
    return "inconnu"


def texte_cellule(v):
    """Écrit une valeur de cellule sans perte : nombres exacts, dates au format ISO."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "VRAI" if v else "FAUX"
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return ""
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, datetime):
        return v.isoformat(sep=" ")
    if isinstance(v, (date, dtime)):
        return v.isoformat()
    if isinstance(v, timedelta):
        return str(v)
    return str(v)


def somme_nombres(lignes):
    total, n = 0.0, 0
    for ligne in lignes:
        for v in ligne:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                    continue
                total += v
                n += 1
    return total, n


def lire_calamine(donnees):
    wb = CalamineWorkbook.from_filelike(io.BytesIO(donnees))
    for nom in wb.sheet_names:
        yield nom, wb.get_sheet_by_name(nom).to_python(skip_empty_area=False)


def lire_openpyxl(donnees):
    wb = openpyxl.load_workbook(io.BytesIO(donnees), read_only=True, data_only=True)
    for ws in wb.worksheets:
        yield ws.title, [list(r) for r in ws.iter_rows(values_only=True)]


# =============================================================================
# 4. PARCOURS DU DOSSIER SOURCE
# =============================================================================
def lister_fichiers(source):
    """Renvoie (chemin relatif, contenu dans le ZIP, fonction de lecture des octets, taille)."""
    fichiers, archives = [], []
    for racine, _, noms in os.walk(source):
        for nom in sorted(noms):
            chemin = os.path.join(racine, nom)
            relatif = os.path.relpath(chemin, source)
            if nom.lower().endswith(".zip"):
                archives.append((chemin, relatif))
            else:
                fichiers.append((relatif, "", (lambda p=chemin: open(p, "rb").read()), os.path.getsize(chemin)))
    # Les fichiers hors ZIP passent en premier : une copie dans un ZIP est alors reconnue comme doublon.
    for chemin, relatif in archives:
        try:
            z = zipfile.ZipFile(chemin)
            for info in z.infolist():
                if info.is_dir():
                    continue
                fichiers.append((relatif, info.filename,
                                 (lambda z=z, i=info: z.read(i)), info.file_size))
        except Exception as e:
            fichiers.append((relatif, "", (lambda p=chemin: open(p, "rb").read()), os.path.getsize(chemin)))
            print(f"  ZIP illisible, copié tel quel : {relatif} ({e})")
    return fichiers


# =============================================================================
# 5. TRAITEMENT
# =============================================================================
def main():
    debut = time.time()
    if not os.path.isdir(DOSSIER_SOURCE):
        sys.exit(f"ERREUR : le dossier source est introuvable : {DOSSIER_SOURCE}")

    horodatage = datetime.now().strftime("%Y%m%d_%H%M")
    racine = os.path.join(DOSSIER_SORTIE, f"Extraction_FPPN_{horodatage}")
    dossier_donnees = os.path.join(racine, "donnees")
    dossier_tableaux = os.path.join(dossier_donnees, "tableaux")
    dossier_bruts = os.path.join(dossier_donnees, "bruts")
    dossier_envoi = os.path.join(racine, "A_ENVOYER")
    if not SIMULATION:
        for d in (dossier_tableaux, dossier_bruts, dossier_envoi):
            os.makedirs(d, exist_ok=True)

    print("Recherche des fichiers...")
    fichiers = lister_fichiers(DOSSIER_SOURCE)
    print(f"{len(fichiers)} fichiers trouvés.\n")

    journal, feuilles_journal = [], []
    vus = {}                      # empreinte -> identifiant du premier fichier identique
    volume_ecrit = 0

    for i, (relatif, membre, lire, taille) in enumerate(fichiers, 1):
        ident = f"F{i:05d}"
        nom = os.path.basename(membre) if membre else os.path.basename(relatif)
        chemin_affiche = relatif + (f" > {membre}" if membre else "")
        chemin_test = os.path.join(os.path.dirname(relatif), membre) if membre else relatif
        ext = os.path.splitext(nom)[1].lower()
        ligne = {"id": ident, "dossier": os.path.dirname(relatif) if not membre else relatif,
                 "fichier": nom, "dans_zip": membre, "extension": ext, "taille_octets": taille,
                 "sha256": "", "format_reel": "", "action": "", "sortie": "", "nb_feuilles": "",
                 "nb_lignes": "", "nb_cellules_non_vides": "", "nb_nombres": "", "somme_nombres": "",
                 "remarque": ""}
        print(f"[{i}/{len(fichiers)}] {chemin_affiche}")

        # Fichiers hors analyse : listés, non lus.
        if est_exclu(chemin_test):
            ligne["action"] = "Non transmis : dossier hors analyse (pièces scannées)"
            journal.append(ligne)
            continue
        if ext == ".pdf" and not pdf_a_transmettre(chemin_test):
            ligne["action"] = "Non transmis : PDF (version imprimée ou scan)"
            journal.append(ligne)
            continue
        if ext in (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif"):
            ligne["action"] = "Non transmis : image"
            journal.append(ligne)
            continue

        try:
            donnees = lire()
        except Exception as e:
            ligne["action"] = "Échec de lecture du fichier"
            ligne["remarque"] = f"{type(e).__name__} : {e}"
            journal.append(ligne)
            continue
        empreinte = sha256(donnees)
        ligne["sha256"] = empreinte

        # Doublons exacts : un seul exemplaire transmis.
        if empreinte in vus:
            ligne["action"] = f"Non transmis : doublon exact de {vus[empreinte]}"
            journal.append(ligne)
            continue
        vus[empreinte] = ident

        base = f"{ident}_{nom_sur(os.path.splitext(nom)[0])}"

        # Tableurs : une feuille = un CSV brut.
        if ext in EXT_TABLEUR:
            fmt = format_tableur(donnees)
            ligne["format_reel"] = fmt
            if fmt == "page web":
                # Export logiciel au format page web : copié tel quel (texte, se compresse très bien).
                sortie = os.path.join(dossier_bruts, base + ext)
                if not SIMULATION:
                    open(sortie, "wb").write(donnees)
                volume_ecrit += len(donnees)
                ligne["action"] = "Copié tel quel (tableau au format page web)"
                ligne["sortie"] = os.path.relpath(sortie, dossier_donnees)
                journal.append(ligne)
                continue
            lecteurs = [("calamine", lire_calamine)]
            if fmt == "xlsx":
                lecteurs.append(("openpyxl", lire_openpyxl))
            erreurs, reussi = [], False
            for nom_lecteur, lecteur in lecteurs:
                try:
                    feuilles = list(lecteur(donnees))
                    reussi = True
                    break
                except Exception as e:
                    erreurs.append(f"{nom_lecteur} : {type(e).__name__} : {str(e)[:150]}")
            if not reussi:
                sortie = os.path.join(dossier_bruts, base + ext)
                if not SIMULATION:
                    open(sortie, "wb").write(donnees)
                volume_ecrit += len(donnees)
                ligne["action"] = "Copié tel quel (illisible par le script)"
                ligne["sortie"] = os.path.relpath(sortie, dossier_donnees)
                ligne["remarque"] = " | ".join(erreurs)
                journal.append(ligne)
                continue
            tot_lignes = tot_cellules = tot_nombres = 0
            tot_somme = 0.0
            sorties = []
            for k, (nom_feuille, lignes) in enumerate(feuilles, 1):
                s, n = somme_nombres(lignes)
                nb_cell = sum(1 for l in lignes for v in l if v not in (None, ""))
                fichier_csv = os.path.join(dossier_tableaux, f"{base}__{k:02d}_{nom_sur(nom_feuille, 30)}.csv")
                if not SIMULATION:
                    with open(fichier_csv, "w", newline="", encoding="utf-8") as f:
                        w = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
                        for l in lignes:
                            w.writerow([texte_cellule(v) for v in l])
                    volume_ecrit += os.path.getsize(fichier_csv)
                else:
                    volume_ecrit += sum(len(";".join(texte_cellule(v) for v in l)) + 2 for l in lignes)
                sorties.append(os.path.relpath(fichier_csv, dossier_donnees))
                feuilles_journal.append({"id": ident, "feuille_n": k, "feuille": nom_feuille,
                                         "csv": os.path.basename(fichier_csv), "nb_lignes": len(lignes),
                                         "nb_colonnes": max((len(l) for l in lignes), default=0),
                                         "nb_cellules_non_vides": nb_cell, "nb_nombres": n,
                                         "somme_nombres": texte_cellule(round(s, 6))})
                tot_lignes += len(lignes)
                tot_cellules += nb_cell
                tot_nombres += n
                tot_somme += s
            ligne.update({"action": "Converti en CSV (une feuille = un fichier)",
                          "sortie": " | ".join(sorties), "nb_feuilles": len(feuilles),
                          "nb_lignes": tot_lignes, "nb_cellules_non_vides": tot_cellules,
                          "nb_nombres": tot_nombres, "somme_nombres": texte_cellule(round(tot_somme, 6))})
            journal.append(ligne)
            continue

        # Fichiers texte, documents et PDF retenus : copiés tels quels.
        if ext in EXT_A_COPIER or ext == ".pdf" or taille <= TAILLE_COPIE_MAX_MO * 1024 * 1024:
            if taille > TAILLE_COPIE_MAX_MO * 1024 * 1024:
                ligne["action"] = f"Non transmis : plus de {TAILLE_COPIE_MAX_MO} Mo"
                journal.append(ligne)
                continue
            sortie = os.path.join(dossier_bruts, base + ext)
            if not SIMULATION:
                open(sortie, "wb").write(donnees)
            volume_ecrit += len(donnees)
            ligne["action"] = "Copié tel quel"
            ligne["sortie"] = os.path.relpath(sortie, dossier_donnees)
            if ext in (".csv", ".txt"):
                ligne["nb_lignes"] = donnees.count(b"\n") + (0 if donnees.endswith(b"\n") or not donnees else 1)
            journal.append(ligne)
            continue

        ligne["action"] = f"Non transmis : plus de {TAILLE_COPIE_MAX_MO} Mo"
        journal.append(ligne)

    # -------------------------------------------------------------------------
    # 6. JOURNAL DE CONTRÔLE
    # -------------------------------------------------------------------------
    print("\nÉcriture du journal de contrôle...")
    colonnes = list(journal[0].keys()) if journal else []
    if not SIMULATION:
        with open(os.path.join(dossier_donnees, "journal_fichiers.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=colonnes, delimiter=";")
            w.writeheader()
            w.writerows(journal)
        if feuilles_journal:
            with open(os.path.join(dossier_donnees, "journal_feuilles.csv"), "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=list(feuilles_journal[0].keys()), delimiter=";")
                w.writeheader()
                w.writerows(feuilles_journal)

    # Synthèse par action.
    synthese = {}
    for l in journal:
        cle = l["action"].split(" de F")[0]
        n, o = synthese.get(cle, (0, 0))
        synthese[cle] = (n + 1, o + (l["taille_octets"] or 0))
    lignes_synthese = [f"Extraction FPPN du {datetime.now():%d/%m/%Y %H:%M}",
                       f"Dossier source : {DOSSIER_SOURCE}",
                       f"Fichiers trouvés : {len(journal)}", ""]
    for cle, (n, o) in sorted(synthese.items(), key=lambda x: -x[1][0]):
        lignes_synthese.append(f"{n:6d} fichiers  {o / 1048576:9.1f} Mo  {cle}")
    lignes_synthese.append("")
    lignes_synthese.append(f"Volume écrit avant compression : {volume_ecrit / 1048576:.1f} Mo")
    texte_synthese = "\n".join(lignes_synthese)
    print("\n" + texte_synthese)

    if SIMULATION:
        print("\nSIMULATION : rien n'a été écrit. Mettre SIMULATION = False pour lancer l'extraction.")
        return

    open(os.path.join(dossier_donnees, "synthese.txt"), "w", encoding="utf-8").write(texte_synthese)

    # -------------------------------------------------------------------------
    # 7. ARCHIVE ZIP DÉCOUPÉE EN PARTIES
    # -------------------------------------------------------------------------
    print("\nCompression...")
    archive = os.path.join(racine, f"FPPN_extraction_{horodatage}.zip")
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for d, _, noms in os.walk(dossier_donnees):
            for n in noms:
                p = os.path.join(d, n)
                z.write(p, os.path.relpath(p, dossier_donnees))
    taille_archive = os.path.getsize(archive)
    empreinte_archive = sha256(open(archive, "rb").read())

    taille_partie = int(TAILLE_PARTIE_MO * 1024 * 1024)
    parties = []
    with open(archive, "rb") as f:
        k = 1
        while True:
            bloc = f.read(taille_partie)
            if not bloc:
                break
            nom_partie = os.path.join(dossier_envoi, f"FPPN_extraction_{horodatage}.zip.{k:03d}")
            open(nom_partie, "wb").write(bloc)
            parties.append((os.path.basename(nom_partie), len(bloc), sha256(bloc)))
            k += 1
    with open(os.path.join(dossier_envoi, "LISEZMOI_parties.txt"), "w", encoding="utf-8") as f:
        f.write(f"Archive : FPPN_extraction_{horodatage}.zip\n")
        f.write(f"Taille : {taille_archive} octets\nSHA-256 : {empreinte_archive}\n")
        f.write(f"Nombre de parties : {len(parties)}\n\n")
        for n, t, e in parties:
            f.write(f"{n}  {t} octets  SHA-256 {e}\n")
    os.remove(archive)

    duree = time.time() - debut
    print(f"\nTerminé en {duree / 60:.1f} minutes.")
    print(f"Archive compressée : {taille_archive / 1048576:.1f} Mo, en {len(parties)} partie(s).")
    print(f"Fichiers à envoyer (toutes les parties et LISEZMOI_parties.txt) : {dossier_envoi}")


if __name__ == "__main__":
    main()
