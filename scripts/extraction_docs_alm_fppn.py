# -*- coding: utf-8 -*-
"""
FPPN - Extraction du texte de la documentation externe sur la gestion actif-passif (ALM)
=======================================================================================
Les PDF sont volumineux : le script extrait leur texte page par page (aucune modification des fichiers d'origine),
signale les pages sans texte (pages scannées ou images) et produit une archive ZIP légère à transmettre :
un fichier .txt par document (repère « ===== Page n / N ===== » avant chaque page) et un journal
(nom, taille, empreinte SHA-256, nombre de pages, pages sans texte, nombre de caractères).
Utilisation (VS Code ou terminal Windows) : python extraction_docs_alm_fppn.py
"""
# ============================================================================= PARAMÈTRES
DOSSIER = r"C:\Dev\Missions\05-FPPN\07-Collecte documentaire\ALM"
FICHIERS = [  # vide = tous les PDF du dossier
    "2024081723_jerome-lavoie-pilote.pdf",
    "ALM_ONT.pdf",
    "Ch13_Gestion_Actif-Passif_French.pdf",
    "Construction d'un modèle ALM.pdf",
    "Gestion des Risques Bancaires(Actif PAssif).pdf",
]
SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses\envoi_ALM"
TAILLE_PARTIE_MO = 20
# =============================================================================
import csv
import hashlib
import importlib
import os
import subprocess
import sys
import zipfile

for paquet, module in (("pypdf", "pypdf"),):
    try:
        importlib.import_module(module)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", paquet])
from pypdf import PdfReader  # noqa: E402


def empreinte(chemin):
    h = hashlib.sha256()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1 << 20), b""):
            h.update(bloc)
    return h.hexdigest()


os.makedirs(SORTIE, exist_ok=True)
noms = FICHIERS or sorted(n for n in os.listdir(DOSSIER) if n.lower().endswith(".pdf"))
journal = []
textes = []
for nom in noms:
    chemin = os.path.join(DOSSIER, nom)
    if not os.path.isfile(chemin):
        print("INTROUVABLE :", chemin)
        continue
    print("Lecture :", nom)
    ligne = {"fichier": nom, "taille_octets": os.path.getsize(chemin), "sha256": empreinte(chemin), "pages": 0,
             "pages_sans_texte": "", "caracteres": 0, "remarque": ""}
    sortie_txt = os.path.join(SORTIE, os.path.splitext(nom)[0] + ".txt")
    try:
        r = PdfReader(chemin)
        if r.is_encrypted:
            r.decrypt("")
        n = len(r.pages)
        vides, morceaux = [], []
        for i, p in enumerate(r.pages, 1):
            try:
                t = p.extract_text() or ""
            except Exception as e:  # page illisible : signalée, jamais bloquante
                t, ligne["remarque"] = "", f"page {i} illisible : {type(e).__name__}"
            if len(t.strip()) < 20:
                vides.append(str(i))
            morceaux.append(f"\n===== Page {i} / {n} =====\n{t}")
        texte = "".join(morceaux)
        with open(sortie_txt, "w", encoding="utf-8") as h:
            h.write(f"Document : {nom}\nPages : {n}\n" + texte)
        ligne.update(pages=n, pages_sans_texte=", ".join(vides), caracteres=len(texte))
        if len(vides) > n / 2:
            ligne["remarque"] = (ligne["remarque"] + " ; " if ligne["remarque"] else "") + "document surtout scanné : transmettre aussi le PDF"
        textes.append(sortie_txt)
    except Exception as e:
        ligne["remarque"] = f"Illisible : {type(e).__name__} {str(e)[:120]}"
    journal.append(ligne)
    print(f"   {ligne['pages']} pages, {len(ligne['pages_sans_texte'].split(', ')) if ligne['pages_sans_texte'] else 0} sans texte, {ligne['caracteres']} caractères {ligne['remarque']}")

fj = os.path.join(SORTIE, "journal_ALM.csv")
with open(fj, "w", newline="", encoding="utf-8-sig") as h:
    w = csv.DictWriter(h, fieldnames=list(journal[0].keys()), delimiter=";")
    w.writeheader()
    w.writerows(journal)
archive = os.path.join(SORTIE, "FPPN_documentation_ALM_texte.zip")
with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
    for f in textes + [fj]:
        z.write(f, os.path.basename(f))
taille, k = os.path.getsize(archive), 0
if taille > TAILLE_PARTIE_MO * 1024 * 1024:
    with open(archive, "rb") as f:
        while True:
            bloc = f.read(TAILLE_PARTIE_MO * 1024 * 1024)
            if not bloc:
                break
            k += 1
            with open(f"{archive}.{k:03d}", "wb") as p:
                p.write(bloc)
    os.remove(archive)
print(f"\n{len(textes)} documents extraits ; archive de {taille / 1e6:.2f} Mo" + (f" en {k} parties" if k else "") + f" dans {SORTIE}")
print("Envoyer l'archive (ou ses parties) ; si le journal signale un document surtout scanné, envoyer aussi ce PDF.")
