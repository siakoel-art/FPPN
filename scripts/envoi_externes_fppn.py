# -*- coding: utf-8 -*-
"""
FPPN - Préparation de l'envoi des documents externes volumineux (Code CIMA, référentiel comptable des SFD)
========================================================================================================
Pour chaque PDF, le script produit dans le dossier de sortie :
  1. <nom>_texte.txt : le texte de toutes les pages, page par page (léger, suffit à Adenka si le PDF n'est pas scanné) ;
  2. <nom>_extrait.pdf : les seules pages qui contiennent les mots utiles à la mission (provisions, créances en souffrance…) ;
  3. si le PDF est scanné ou si son texte est illisible (police mal encodée), des parties <nom>_partie_01.pdf, _02… de 15 Mo
     au plus, copiées sans modification : Adenka les lit alors par reconnaissance de caractères (OCR) ;
  4. journal_envoi.csv : nom, pages, taille et empreinte SHA-256 de chaque fichier produit, pour contrôler la réception.
Aucun fichier source n'est modifié.

Installation (une seule fois) :  python -m pip install pypdf
Utilisation :                    python envoi_externes_fppn.py
Puis transmettre à Adenka le contenu du dossier SORTIE (les .txt et _extrait.pdf d'abord, les parties seulement si demandé).
"""
# ============================================================================= PARAMÈTRES
DOSSIER = r"C:\Dev\Missions\05-FPPN\07-Collecte documentaire\Collecte additionnelle_rapport-v10"
FICHIERS = ["Référentiel comptable spécifique des SFD_1.pdf", "CODE-CIMA-2019.pdf"]
SORTIE = r"C:\Dev\Missions\05-FPPN\07-Collecte documentaire\Collecte additionnelle_rapport-v10\envoi_adenka"
TAILLE_PARTIE_MO = 15
FORCER_DECOUPAGE = True  # True : découpe toujours le PDF complet en parties (texte des deux PDF illisible à l'envoi du 23/10/2026)
MOTS = {  # pages à extraire, par fichier (mots cherchés sans tenir compte des majuscules)
    "Référentiel comptable spécifique des SFD_1.pdf": ["souffrance", "provision", "dépréciation", "impayé", "29 -", "299", "ancienneté"],
    "CODE-CIMA-2019.pdf": ["article 334", "article 335", "article 336", "article 337", "article 338", "provision mathématique",
                           "sinistres à payer", "tardifs", "table de mortalité", "taux d'escompte", "taux d'actualisation"],
}
# =============================================================================
import csv
import hashlib
import os

from pypdf import PdfReader, PdfWriter


def empreinte(chemin):
    h = hashlib.sha256()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1 << 20), b""):
            h.update(bloc)
    return h.hexdigest()


def ecrire(writer, chemin):
    with open(chemin, "wb") as f:
        writer.write(f)


os.makedirs(SORTIE, exist_ok=True)
journal = []
for nom in FICHIERS:
    src = os.path.join(DOSSIER, nom)
    base = os.path.splitext(nom)[0].replace(" ", "_")
    lecteur = PdfReader(src)
    n = len(lecteur.pages)
    textes = []
    for i, page in enumerate(lecteur.pages, 1):
        try:
            textes.append(page.extract_text() or "")
        except Exception as e:  # page illisible : on la signale sans arrêter
            textes.append(f"[page {i} illisible : {e}]")
    moyenne = sum(len(t.strip()) for t in textes) / max(n, 1)
    print(f"{nom} : {n} pages, {os.path.getsize(src) / 1e6:.1f} Mo, {moyenne:.0f} caractères par page")

    # 1. texte intégral
    txt = os.path.join(SORTIE, base + "_texte.txt")
    with open(txt, "w", encoding="utf-8") as f:
        for i, t in enumerate(textes, 1):
            f.write(f"\n===== PAGE {i} =====\n{t}\n")
    journal.append((os.path.basename(txt), n, "texte intégral"))

    # 2. extrait des pages utiles
    mots = [m.lower() for m in MOTS.get(nom, [])]
    pages = [i for i, t in enumerate(textes) if any(m in t.lower() for m in mots)]
    if pages:
        w = PdfWriter()
        for i in pages:
            w.add_page(lecteur.pages[i])
        ext = os.path.join(SORTIE, base + "_extrait.pdf")
        ecrire(w, ext)
        journal.append((os.path.basename(ext), len(pages), "pages " + ", ".join(str(i + 1) for i in pages)))

    # 3. PDF scanné ou texte illisible : découpage en parties de TAILLE_PARTIE_MO
    lisibles = sum(1 for t in textes if " de " in t.lower() and " la " in t.lower()) / max(n, 1)
    print(f"  pages au texte lisible : {lisibles:.0%}")
    if FORCER_DECOUPAGE or moyenne < 200 or lisibles < 0.5:
        print("  texte absent ou illisible : découpage du PDF complet en parties")
        taille_page = os.path.getsize(src) / max(n, 1)
        par_partie = max(1, int(TAILLE_PARTIE_MO * 1e6 / taille_page))
        for k, debut in enumerate(range(0, n, par_partie), 1):
            w = PdfWriter()
            for i in range(debut, min(n, debut + par_partie)):
                w.add_page(lecteur.pages[i])
            part = os.path.join(SORTIE, f"{base}_partie_{k:02d}.pdf")
            ecrire(w, part)
            journal.append((os.path.basename(part), min(n, debut + par_partie) - debut, f"pages {debut + 1} à {min(n, debut + par_partie)}"))

with open(os.path.join(SORTIE, "journal_envoi.csv"), "w", newline="", encoding="utf-8-sig") as f:
    wr = csv.writer(f, delimiter=";")
    wr.writerow(["fichier", "pages", "contenu", "taille_octets", "sha256"])
    for fic, pages, contenu in journal:
        p = os.path.join(SORTIE, fic)
        wr.writerow([fic, pages, contenu, os.path.getsize(p), empreinte(p)])
        print(f"  {fic} : {os.path.getsize(p) / 1e6:.2f} Mo")
print("Terminé. Dossier à transmettre :", SORTIE)
