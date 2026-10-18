# -*- coding: utf-8 -*-
"""
FPPN - Emballer quelques fichiers précis pour l'envoi (archive ZIP découpée en parties)
======================================================================================

À utiliser quand un fichier ne passe pas à l'envoi direct (PDF refusé, fichier trop gros).
Le script copie les fichiers listés dans FICHIERS dans une archive ZIP, la découpe en parties
de taille fixe et écrit un LISEZMOI avec l'empreinte SHA-256 de chaque partie.

Utilisation : vérifier la partie PARAMÈTRES, puis lancer (« Run Python File » dans VS Code).
Le script ne modifie ni ne déplace aucun fichier source.
"""

# =============================================================================
# PARAMÈTRES
# =============================================================================
DOSSIER_SOURCE = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client\01-download\AUDIT ADENKA\DAFC"
DOSSIER_SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses"
TAILLE_PARTIE_MO = 10

FICHIERS = [   # chemins relatifs à DOSSIER_SOURCE
    r"2021\Grand-livre_des_comptes FPPN 2021.pdf",
    r"2022\Grand-livre_des_comptes 2022.pdf",
    r"2024\Grand-livre_des_comptes FPPN 2024.pdf",
    r"2025\Grand-livre_des_tiers FPPN 2025.pdf",
]

# =============================================================================
import hashlib
import os
import sys
import zipfile
from datetime import datetime


def sha256(donnees):
    return hashlib.sha256(donnees).hexdigest()


def chemin_source(f):
    return os.path.join(DOSSIER_SOURCE, *f.replace("/", "\\").split("\\"))


def main():
    horodatage = datetime.now().strftime("%Y%m%d_%H%M")
    nom = f"FPPN_envoi_{horodatage}"
    dossier_envoi = os.path.join(DOSSIER_SORTIE, nom)
    os.makedirs(dossier_envoi, exist_ok=True)
    archive = os.path.join(dossier_envoi, nom + ".zip")

    manquants = [f for f in FICHIERS if not os.path.isfile(chemin_source(f))]
    if manquants:
        sys.exit("ERREUR : fichiers introuvables :\n  " + "\n  ".join(manquants))

    lignes = []
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for f in FICHIERS:
            chemin = chemin_source(f)
            donnees = open(chemin, "rb").read()
            z.writestr(f.replace("\\", "/"), donnees)
            lignes.append(f"{f}  {len(donnees)} octets  SHA-256 {sha256(donnees)}")
            print(f"Ajouté : {f} ({len(donnees) / 1048576:.1f} Mo)")

    contenu = open(archive, "rb").read()
    taille_partie = int(TAILLE_PARTIE_MO * 1024 * 1024)
    parties = []
    for k in range(0, len(contenu), taille_partie):
        bloc = contenu[k:k + taille_partie]
        nom_partie = f"{nom}.zip.{len(parties) + 1:03d}"
        open(os.path.join(dossier_envoi, nom_partie), "wb").write(bloc)
        parties.append(f"{nom_partie}  {len(bloc)} octets  SHA-256 {sha256(bloc)}")
    os.remove(archive)

    with open(os.path.join(dossier_envoi, "LISEZMOI_parties.txt"), "w", encoding="utf-8") as f:
        f.write(f"Archive : {nom}.zip\nTaille : {len(contenu)} octets\nSHA-256 : {sha256(contenu)}\n")
        f.write(f"Nombre de parties : {len(parties)}\n\nFichiers contenus :\n")
        f.write("\n".join(lignes) + "\n\nParties :\n" + "\n".join(parties) + "\n")

    print(f"\nTerminé : {len(parties)} partie(s) dans {dossier_envoi}")


if __name__ == "__main__":
    main()
