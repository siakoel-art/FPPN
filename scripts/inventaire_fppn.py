# -*- coding: utf-8 -*-
"""
FPPN - Inventaire et résumé des documents reçus du client (volet actuariel)
==========================================================================

Ce script :
  1. vérifie et installe les bibliothèques Python nécessaires ;
  2. parcourt tout le dossier des documents reçus (sous-dossiers et fichiers ZIP compris) ;
  3. produit une fiche par fichier : type, taille, feuilles, colonnes, pages, extrait du texte,
     années trouvées, thèmes détectés, doublons ;
  4. lit les scans et les images avec Tesseract (reconnaissance de texte), si Tesseract est installé ;
  5. écrit deux fichiers Excel dans le dossier de sortie :
       - Inventaire_FPPN_complet.xlsx : version complète, à garder pour toi ;
       - Inventaire_FPPN_masque.xlsx  : sans extraits de données (noms, matricules, montants),
                                        à envoyer pour la suite des analyses.

Utilisation : ouvrir ce fichier dans VS Code, vérifier la partie PARAMÈTRES, puis lancer
(bouton « Run Python File » en haut à droite, ou dans le terminal : python inventaire_fppn.py).

Le script ne modifie ni ne déplace aucun fichier du dossier source. Il n'envoie rien sur Internet,
sauf pour installer les bibliothèques manquantes (pip).
"""

# =============================================================================
# PARAMÈTRES (à vérifier avant de lancer)
# =============================================================================
DOSSIER_SOURCE = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client\01-download"
DOSSIER_SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses"

TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # chemin d'installation par défaut sous Windows
OCR_ACTIF = True            # False pour ne pas lire les scans (beaucoup plus rapide)
OCR_PAGES_MAX = 3           # nombre de pages lues par scan
OCR_LANGUE = "fra"          # français (installer le pack « French » avec Tesseract)
LIGNES_ANALYSEES_MAX = 5000 # lignes lues par feuille Excel ou CSV pour repérer les dates et les mots-clés
TAILLE_ZIP_MAX_MO = 300     # au-delà, un fichier contenu dans un ZIP est listé sans être lu

# =============================================================================
# 1. INSTALLATION DES BIBLIOTHÈQUES MANQUANTES
# =============================================================================
import importlib
import subprocess
import sys

BIBLIOTHEQUES = {          # nom à importer : nom à installer avec pip
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
    "pymupdf": "pymupdf",
    "docx": "python-docx",
    "pptx": "python-pptx",
    "pytesseract": "pytesseract",
    "PIL": "Pillow",
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
import os
import re
import time
import unicodedata
import warnings
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, date

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import pymupdf as fitz
except ImportError:  # anciennes versions
    import fitz

warnings.filterwarnings("ignore")

try:
    import pytesseract
    from PIL import Image
    if os.path.exists(TESSERACT_EXE):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
    pytesseract.get_tesseract_version()
    OCR_DISPONIBLE = OCR_ACTIF
    if OCR_LANGUE not in pytesseract.get_languages(config=""):
        print(f"Attention : la langue '{OCR_LANGUE}' n'est pas installée dans Tesseract. Lecture en anglais.")
        OCR_LANGUE = "eng"
except Exception:
    OCR_DISPONIBLE = False

# =============================================================================
# 3. THÈMES ET MOTS-CLÉS DU PÉRIMÈTRE ACTUARIEL
# =============================================================================
THEMES = {
    "Mutualistes et effectifs": ["mutualiste", "adherent", "effectif", "matricule", "date de naissance",
                                  "naissance", "ayant droit", "ayants droit", "conjoint", "enfant", "retraite",
                                  "grade", "anciennete", "date d'entree", "date d'adhesion"],
    "Cotisations": ["cotisation", "precompte", "retenue", "solde", "tresor", "appel de cotisation",
                    "cotisations encaissees", "cotisations appelees"],
    "Prestations et sinistres": ["prestation", "sinistre", "allocation", "deces", "invalidite", "capital deces",
                                 "frais funeraires", "remboursement", "soins", "prise en charge", "facture",
                                 "dossier", "beneficiaire", "cgrae", "arriere"],
    "Provisions et engagements": ["provision", "engagement", "reserve", "psap", "ibnr", "tardif",
                                  "provision mathematique", "provision technique", "actuari",
                                  "dotation", "reprise"],
    "Épargne": ["epargne", "plan epargne", "fin de carriere", "compte individuel", "interet", "taux servi",
                "solde epargne"],
    "Placements et trésorerie": ["placement", "depot a terme", "dat", "obligation", "action", "titre",
                                 "banque", "echeance", "rendement", "tresorerie", "releve bancaire",
                                 "rapprochement bancaire", "coupon"],
    "Prêts aux mutualistes": ["pret", "credit", "grenier du policier", "caisse de credit", "encours",
                              "impaye", "remboursement de pret", "echeancier"],
    "Immobilier et patrimoine": ["immobilier", "immeuble", "terrain", "location-vente", "location vente",
                                 "patrimoine", "sci", "construction", "expertise"],
    "États financiers et comptabilité": ["bilan", "compte de resultat", "balance", "grand livre",
                                         "etats financiers", "journal", "plan comptable", "pcms",
                                         "syscohada", "commissaire aux comptes", "annexe"],
    "Réglementation et gouvernance": ["statuts", "reglement interieur", "airms", "uemoa", "agrement",
                                      "assemblee generale", "conseil d'administration", "proces-verbal",
                                      "proces verbal", "deliberation", "decision", "convention"],
}
THEMES_COEUR = ["Mutualistes et effectifs", "Cotisations", "Prestations et sinistres",
                "Provisions et engagements", "Épargne", "Placements et trésorerie", "Prêts aux mutualistes"]

EXT_EXCEL = {".xlsx", ".xlsm", ".xltx", ".xls"}
EXT_CSV = {".csv", ".txt"}
EXT_PDF = {".pdf"}
EXT_WORD = {".docx", ".doc"}
EXT_PPT = {".pptx", ".ppt"}
EXT_IMAGE = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif"}
EXT_ZIP = {".zip"}

RE_ANNEE = re.compile(r"(?<!\d)(20(?:1[0-9]|2[0-9]))(?!\d)")


def sans_accents(texte):
    texte = unicodedata.normalize("NFKD", str(texte))
    return "".join(c for c in texte if not unicodedata.combining(c)).lower()


def themes_detectes(texte):
    t = sans_accents(texte)
    scores = {}
    for theme, mots in THEMES.items():
        n = 0
        for mot in mots:
            if len(mot) <= 4:  # mots courts : mot entier seulement (ex. « dat », « sci »)
                n += len(re.findall(r"\b" + re.escape(mot) + r"\b", t))
            else:
                n += t.count(mot)
        if n:
            scores[theme] = n
    return scores


def annees_trouvees(texte):
    annees = sorted(set(RE_ANNEE.findall(str(texte))))
    return ", ".join(annees)


def nettoyer(texte, longueur=600):
    texte = re.sub(r"\s+", " ", str(texte or "")).strip()
    return texte[:longueur]


def taille_lisible(octets):
    for unite in ["o", "Ko", "Mo", "Go"]:
        if octets < 1024:
            return f"{octets:.0f} {unite}"
        octets /= 1024
    return f"{octets:.1f} To"


def empreinte(donnees):
    return hashlib.md5(donnees).hexdigest()


def empreinte_fichier(chemin):
    h = hashlib.md5()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloc)
    return h.hexdigest()


def chemin_long(chemin):
    """Sous Windows, permet d'ouvrir les chemins de plus de 260 caractères (fréquent avec OneDrive)."""
    if os.name == "nt":
        chemin = os.path.abspath(chemin)
        if not chemin.startswith("\\\\?\\"):
            return "\\\\?\\" + chemin
    return chemin


# =============================================================================
# 4. LECTURE DES DIFFÉRENTS TYPES DE FICHIERS
#    Chaque fonction reçoit un flux binaire (fichier ou membre de ZIP) et renvoie un dictionnaire.
# =============================================================================
def valeur_date(v):
    if isinstance(v, (datetime, date)):
        return v.year if 1900 <= v.year <= 2100 else None
    return None


def resumer_tableau(lignes, nb_lignes_total, nom_feuille):
    """lignes : liste de listes (premières lignes d'une feuille). Repère l'en-tête, les colonnes, les dates."""
    lignes = [l for l in lignes if any(v not in (None, "") for v in l)]
    if not lignes:
        return {"feuille": nom_feuille, "lignes": 0, "colonnes": 0, "entete": "", "apercu": "",
                "annees": "", "texte": ""}
    # en-tête : première ligne dont au moins la moitié des cellules est du texte
    idx_entete = 0
    for i, l in enumerate(lignes[:20]):
        remplies = [v for v in l if v not in (None, "")]
        textes = [v for v in remplies if isinstance(v, str)]
        if len(remplies) >= 2 and len(textes) >= 0.5 * len(remplies):
            idx_entete = i
            break
    entete = [nettoyer(v, 40) for v in lignes[idx_entete] if v not in (None, "")]
    corps = lignes[idx_entete + 1:]
    annees = Counter()
    for l in corps:
        for v in l:
            a = valeur_date(v)
            if a:
                annees[a] += 1
    texte_annees = annees_trouvees(" ".join(str(v) for l in lignes[:200] for v in l if v is not None))
    if annees:
        periode = f"dates de {min(annees)} à {max(annees)}"
    elif texte_annees:
        periode = texte_annees
    else:
        periode = ""
    apercu = " | ".join(" ; ".join(nettoyer(v, 25) for v in l if v not in (None, "")) for l in corps[:3])
    texte = " ".join(str(v) for l in lignes[:LIGNES_ANALYSEES_MAX] for v in l if isinstance(v, str))
    nb_col = max(len(l) for l in lignes[:200])
    return {"feuille": nom_feuille, "lignes": nb_lignes_total, "colonnes": nb_col,
            "entete": ", ".join(entete[:30]), "apercu": nettoyer(apercu, 500),
            "annees": periode, "texte": texte}


def lire_excel(flux, ext):
    feuilles = []
    if ext == ".xls":
        classeur = pd.read_excel(flux, sheet_name=None, header=None, nrows=LIGNES_ANALYSEES_MAX, engine="xlrd")
        for nom, df in classeur.items():
            lignes = df.where(pd.notna(df), None).values.tolist()
            feuilles.append(resumer_tableau(lignes, len(df), nom))
    else:
        wb = openpyxl.load_workbook(flux, read_only=True, data_only=True)
        for ws in wb.worksheets:
            lignes = []
            for i, ligne in enumerate(ws.iter_rows(values_only=True)):
                if i >= LIGNES_ANALYSEES_MAX:
                    break
                lignes.append(list(ligne))
            total = ws.max_row or len(lignes)
            feuilles.append(resumer_tableau(lignes, total, ws.title))
        wb.close()
    return feuilles


def lire_csv(donnees):
    for encodage in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texte = donnees.decode(encodage)
            break
        except UnicodeDecodeError:
            continue
    try:
        separateur = csv.Sniffer().sniff(texte[:5000], delimiters=";,\t|").delimiter
    except csv.Error:
        separateur = ";"
    lecteur = csv.reader(io.StringIO(texte), delimiter=separateur)
    lignes = []
    for i, l in enumerate(lecteur):
        if i < LIGNES_ANALYSEES_MAX:
            lignes.append(l)
    total = texte.count("\n")
    return [resumer_tableau(lignes, total, "(CSV)")]


def ocr_image(image):
    return pytesseract.image_to_string(image, lang=OCR_LANGUE)


def lire_pdf(donnees):
    doc = fitz.open(stream=donnees, filetype="pdf")
    nb_pages = doc.page_count
    texte = ""
    for i in range(min(nb_pages, 5)):
        texte += doc[i].get_text() + "\n"
    est_scan = len(texte.strip()) < 50 * min(nb_pages, 5)
    ocr = ""
    if est_scan and OCR_DISPONIBLE:
        for i in range(min(nb_pages, OCR_PAGES_MAX)):
            pix = doc[i].get_pixmap(dpi=200)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            ocr += ocr_image(image) + "\n"
    doc.close()
    return {"pages": nb_pages, "scan": est_scan, "texte": texte if not est_scan else ocr,
            "ocr_fait": bool(ocr)}


def lire_word(donnees, ext):
    if ext == ".doc":
        return {"texte": "", "remarque": "Ancien format Word (.doc) : à ouvrir et réenregistrer en .docx pour lecture."}
    import docx
    d = docx.Document(io.BytesIO(donnees))
    paragraphes = [p.text for p in d.paragraphs if p.text.strip()]
    texte = "\n".join(paragraphes)
    for t in d.tables[:20]:
        for ligne in t.rows[:50]:
            texte += "\n" + " ; ".join(c.text for c in ligne.cells)
    return {"texte": texte, "paragraphes": len(paragraphes), "tableaux": len(d.tables)}


def lire_ppt(donnees, ext):
    if ext == ".ppt":
        return {"texte": "", "remarque": "Ancien format PowerPoint (.ppt) : à réenregistrer en .pptx."}
    from pptx import Presentation
    p = Presentation(io.BytesIO(donnees))
    morceaux = []
    for s in p.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                morceaux.append(sh.text_frame.text)
    return {"texte": "\n".join(morceaux), "diapositives": len(p.slides)}


def lire_image(donnees):
    if not OCR_DISPONIBLE:
        return {"texte": "", "remarque": "Image : Tesseract non disponible, à lire à la main."}
    image = Image.open(io.BytesIO(donnees))
    return {"texte": ocr_image(image), "ocr_fait": True}


# =============================================================================
# 5. ANALYSE D'UN FICHIER
# =============================================================================
def analyser(nom, chemin_affiche, donnees, taille, date_modif, origine_zip=""):
    ext = os.path.splitext(nom)[1].lower()
    fiche = {"Dossier": os.path.dirname(chemin_affiche), "Fichier": os.path.basename(nom),
             "Contenu dans le ZIP": origine_zip, "Extension": ext, "Taille": taille_lisible(taille),
             "Taille (octets)": taille, "Modifié le": date_modif, "Type": "", "Pages / feuilles": "",
             "Lignes (max)": "", "Colonnes / en-têtes": "", "Années trouvées": "", "Scan": "",
             "Thèmes détectés": "", "Thème principal": "", "Pertinence": "", "Résumé": "",
             "Extrait": "", "Empreinte": empreinte(donnees) if donnees is not None else "",
             "Doublon de": "", "Remarque / erreur": ""}
    texte_analyse = nom
    feuilles = []
    try:
        if donnees is None:
            fiche["Type"] = "Non lu"
            fiche["Remarque / erreur"] = "Fichier trop volumineux dans le ZIP : à extraire et relancer."
        elif ext in EXT_EXCEL:
            fiche["Type"] = "Excel"
            feuilles = lire_excel(io.BytesIO(donnees), ext)
        elif ext in EXT_CSV:
            fiche["Type"] = "CSV / texte"
            feuilles = lire_csv(donnees)
        elif ext in EXT_PDF:
            fiche["Type"] = "PDF"
            r = lire_pdf(donnees)
            fiche["Pages / feuilles"] = f"{r['pages']} page(s)"
            fiche["Scan"] = "Oui" if r["scan"] else "Non"
            if r["scan"] and not r["ocr_fait"]:
                fiche["Remarque / erreur"] = "Scan non lu (Tesseract absent ou désactivé) : à lire à la main."
            elif r["scan"]:
                fiche["Remarque / erreur"] = f"Scan lu par reconnaissance de texte ({min(r['pages'], OCR_PAGES_MAX)} page(s))."
            texte_analyse += " " + r["texte"]
            fiche["Extrait"] = nettoyer(r["texte"], 600)
        elif ext in EXT_WORD:
            fiche["Type"] = "Word"
            r = lire_word(donnees, ext)
            if "paragraphes" in r:
                fiche["Pages / feuilles"] = f"{r['paragraphes']} paragraphe(s), {r['tableaux']} tableau(x)"
            fiche["Remarque / erreur"] = r.get("remarque", "")
            texte_analyse += " " + r["texte"]
            fiche["Extrait"] = nettoyer(r["texte"], 600)
        elif ext in EXT_PPT:
            fiche["Type"] = "PowerPoint"
            r = lire_ppt(donnees, ext)
            if "diapositives" in r:
                fiche["Pages / feuilles"] = f"{r['diapositives']} diapositive(s)"
            fiche["Remarque / erreur"] = r.get("remarque", "")
            texte_analyse += " " + r["texte"]
            fiche["Extrait"] = nettoyer(r["texte"], 600)
        elif ext in EXT_IMAGE:
            fiche["Type"] = "Image"
            fiche["Scan"] = "Oui"
            r = lire_image(donnees)
            fiche["Remarque / erreur"] = r.get("remarque", "Image lue par reconnaissance de texte.")
            texte_analyse += " " + r["texte"]
            fiche["Extrait"] = nettoyer(r["texte"], 600)
        else:
            fiche["Type"] = "Autre"
            fiche["Remarque / erreur"] = "Type non lu par le script."
    except Exception as e:
        fiche["Remarque / erreur"] = f"Erreur de lecture : {type(e).__name__} : {nettoyer(e, 200)}"

    if feuilles:
        fiche["Pages / feuilles"] = f"{len(feuilles)} feuille(s) : " + ", ".join(str(f["feuille"]) for f in feuilles[:15])
        fiche["Lignes (max)"] = max(f["lignes"] for f in feuilles)
        fiche["Colonnes / en-têtes"] = " || ".join(f"[{f['feuille']}] {f['entete']}" for f in feuilles[:10])
        fiche["Extrait"] = nettoyer(" || ".join(f"[{f['feuille']}] {f['apercu']}" for f in feuilles[:5]), 600)
        texte_analyse += " " + " ".join(f["entete"] + " " + f["texte"] for f in feuilles)
        annees_cellules = [f["annees"] for f in feuilles if f["annees"]]
        fiche["Années trouvées"] = " ; ".join(annees_cellules[:5])

    if not fiche["Années trouvées"]:
        fiche["Années trouvées"] = annees_trouvees(texte_analyse)

    scores = themes_detectes(texte_analyse)
    if scores:
        tri = sorted(scores.items(), key=lambda kv: -kv[1])
        fiche["Thèmes détectés"] = ", ".join(f"{t} ({n})" for t, n in tri)
        fiche["Thème principal"] = tri[0][0]
    fiche["Pertinence"] = pertinence(scores, fiche)
    fiche["Résumé"] = rediger_resume(fiche, feuilles)
    fiche["_scores"] = scores
    fiche["_feuilles"] = feuilles
    return fiche


def pertinence(scores, fiche):
    coeur = sum(n for t, n in scores.items() if t in THEMES_COEUR)
    lignes = fiche["Lignes (max)"] or 0
    if coeur >= 10 or (coeur >= 3 and isinstance(lignes, int) and lignes >= 100):
        return "Forte"
    if coeur >= 3 or sum(scores.values()) >= 5:
        return "Moyenne"
    # statuts, règlements et états financiers : toujours utiles (règles des prestations, provisions au bilan)
    if scores.get("Réglementation et gouvernance") or scores.get("États financiers et comptabilité"):
        return "Moyenne"
    return "Faible"


def rediger_resume(fiche, feuilles):
    morceaux = [fiche["Type"] or "Fichier"]
    if feuilles:
        for f in feuilles[:4]:
            m = f"feuille « {f['feuille']} » : {f['lignes']} lignes × {f['colonnes']} colonnes"
            if f["entete"]:
                m += f" (colonnes : {nettoyer(f['entete'], 150)})"
            morceaux.append(m)
        if len(feuilles) > 4:
            morceaux.append(f"et {len(feuilles) - 4} autre(s) feuille(s)")
    elif fiche["Pages / feuilles"]:
        morceaux.append(fiche["Pages / feuilles"])
    if fiche["Scan"] == "Oui":
        morceaux.append("document scanné")
    if fiche["Années trouvées"]:
        morceaux.append(f"années : {nettoyer(fiche['Années trouvées'], 80)}")
    if fiche["Thème principal"]:
        morceaux.append(f"thème principal : {fiche['Thème principal']}")
    morceaux = [m[:1].upper() + m[1:] for m in morceaux if m]
    return nettoyer(". ".join(morceaux) + ".", 900)


# =============================================================================
# 6. PARCOURS DU DOSSIER (ZIP COMPRIS)
# =============================================================================
def parcourir(dossier):
    fichiers = []
    for racine, _, noms in os.walk(chemin_long(dossier)):
        for nom in noms:
            if nom.startswith("~$") or nom.lower() in ("desktop.ini", "thumbs.db"):
                continue
            fichiers.append(os.path.join(racine, nom))
    return fichiers


def relatif(chemin):
    c = chemin.replace("\\\\?\\", "")
    try:
        return os.path.relpath(c, DOSSIER_SOURCE)
    except ValueError:
        return c


def traiter_zip(chemin, fiches, n_total, compteur):
    with zipfile.ZipFile(chemin) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            nom = info.filename
            try:
                nom = nom.encode("cp437").decode("utf-8")  # noms accentués dans les ZIP Windows
            except Exception:
                pass
            date_modif = datetime(*info.date_time).strftime("%Y-%m-%d %H:%M")
            donnees = None
            if info.file_size <= TAILLE_ZIP_MAX_MO * 1024 * 1024:
                donnees = z.read(info)
            ext = os.path.splitext(nom)[1].lower()
            if ext in EXT_ZIP and donnees is not None:
                fiches.append({"Fichier": os.path.basename(nom), "Contenu dans le ZIP": relatif(chemin),
                               "Type": "ZIP", "Remarque / erreur": "ZIP dans un ZIP : à extraire puis relancer.",
                               "Taille": taille_lisible(info.file_size), "Taille (octets)": info.file_size,
                               "Dossier": os.path.dirname(nom), "Extension": ext, "Modifié le": date_modif,
                               "Empreinte": empreinte(donnees), "_scores": {}, "_feuilles": []})
                continue
            fiches.append(analyser(nom, os.path.join(relatif(chemin), nom), donnees, info.file_size,
                                   date_modif, origine_zip=relatif(chemin)))


def lancer():
    debut = time.time()
    print("=" * 70)
    print("Inventaire des documents FPPN")
    print(f"Dossier source : {DOSSIER_SOURCE}")
    print(f"Reconnaissance de texte (Tesseract) : {'oui' if OCR_DISPONIBLE else 'non'}")
    print("=" * 70)
    if not os.path.isdir(chemin_long(DOSSIER_SOURCE)):
        print("ERREUR : le dossier source est introuvable. Vérifie DOSSIER_SOURCE en haut du script.")
        return
    chemins = parcourir(DOSSIER_SOURCE)
    print(f"{len(chemins)} fichier(s) trouvé(s). Analyse en cours...\n")
    fiches = []
    for i, chemin in enumerate(chemins, 1):
        rel = relatif(chemin)
        print(f"[{i}/{len(chemins)}] {rel[:110]}")
        try:
            taille = os.path.getsize(chemin)
            date_modif = datetime.fromtimestamp(os.path.getmtime(chemin)).strftime("%Y-%m-%d %H:%M")
            ext = os.path.splitext(chemin)[1].lower()
            if ext in EXT_ZIP:
                traiter_zip(chemin, fiches, len(chemins), i)
                continue
            with open(chemin, "rb") as f:
                donnees = f.read()
            fiches.append(analyser(os.path.basename(chemin), rel, donnees, taille, date_modif))
        except Exception as e:
            fiches.append({"Dossier": os.path.dirname(rel), "Fichier": os.path.basename(rel),
                           "Remarque / erreur": f"Erreur : {type(e).__name__} : {nettoyer(e, 200)}",
                           "_scores": {}, "_feuilles": []})

    # doublons (même contenu) : le fichier hors ZIP, puis le premier par ordre alphabétique, est l'original
    fiches.sort(key=lambda f: (bool(f.get("Contenu dans le ZIP")), f.get("Dossier", ""), f.get("Fichier", "")))
    vus = {}
    for f in fiches:
        e = f.get("Empreinte")
        if not e:
            continue
        if e in vus:
            f["Doublon de"] = vus[e]
        else:
            vus[e] = os.path.join(f.get("Contenu dans le ZIP", ""), f.get("Dossier", ""), f.get("Fichier", ""))

    os.makedirs(chemin_long(DOSSIER_SORTIE), exist_ok=True)
    horodatage = datetime.now().strftime("%Y%m%d_%H%M")
    complet = os.path.join(DOSSIER_SORTIE, f"Inventaire_FPPN_complet_{horodatage}.xlsx")
    masque = os.path.join(DOSSIER_SORTIE, f"Inventaire_FPPN_masque_{horodatage}.xlsx")
    ecrire_excel(fiches, complet, masquer=False)
    ecrire_excel(fiches, masque, masquer=True)

    duree = (time.time() - debut) / 60
    print("\n" + "=" * 70)
    print(f"Terminé en {duree:.1f} minute(s). {len(fiches)} fichier(s) analysé(s).")
    print(f"Version complète (à garder) : {complet}")
    print(f"Version masquée (à envoyer)  : {masque}")
    print("=" * 70)


# =============================================================================
# 7. ÉCRITURE DE L'EXCEL (CHARTE ADENKA)
# =============================================================================
BLEU_NUIT, ORANGE, GRIS, GRIS_CLAIR = "112A40", "EFAC62", "BBBCBC", "F2F2F2"
POLICE = "Helvetica"
BORD = Border(*(Side(style="thin", color=GRIS),) * 4)

COLONNES_INVENTAIRE = ["N°", "Dossier", "Fichier", "Contenu dans le ZIP", "Type", "Extension", "Taille",
                       "Modifié le", "Pages / feuilles", "Lignes (max)", "Colonnes / en-têtes",
                       "Années trouvées", "Scan", "Thème principal", "Thèmes détectés", "Pertinence",
                       "Doublon de", "Remarque / erreur"]
LARGEURS = {"N°": 6, "Dossier": 30, "Fichier": 40, "Contenu dans le ZIP": 24, "Type": 12, "Extension": 10,
            "Taille": 10, "Modifié le": 16, "Pages / feuilles": 30, "Lignes (max)": 11,
            "Colonnes / en-têtes": 60, "Années trouvées": 24, "Scan": 7, "Thème principal": 26,
            "Thèmes détectés": 45, "Pertinence": 11, "Doublon de": 30, "Remarque / erreur": 40,
            "Résumé": 80, "Extrait": 80, "Thème": 30, "Score": 8}


def feuille(wb, titre, colonnes, lignes, note=""):
    ws = wb.create_sheet(titre)
    ws.sheet_view.showGridLines = False
    ws["A1"] = titre
    ws["A1"].font = Font(name=POLICE, size=14, bold=True)
    ws["A2"] = note
    ws["A2"].font = Font(name=POLICE, size=9, italic=True, color="53565A")
    for j, c in enumerate(colonnes, 1):
        cell = ws.cell(4, j, c)
        cell.font = Font(name=POLICE, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=BLEU_NUIT)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = BORD
        ws.column_dimensions[get_column_letter(j)].width = LARGEURS.get(c, 18)
    for i, ligne in enumerate(lignes, 5):
        for j, v in enumerate(ligne, 1):
            if isinstance(v, str):
                v = v[:32000]
                v = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", v)  # caractères refusés par Excel
            cell = ws.cell(i, j, v)
            cell.font = Font(name=POLICE, size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = BORD
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=GRIS_CLAIR)
        if "Pertinence" in colonnes:
            p = ws.cell(i, colonnes.index("Pertinence") + 1)
            if p.value == "Forte":
                p.font = Font(name=POLICE, size=9, bold=True, color="B5651D")
    ws.freeze_panes = "D5" if len(colonnes) > 4 else "A5"
    if lignes:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(colonnes))}{len(lignes) + 4}"
    return ws


def ecrire_excel(fiches, chemin, masquer):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ordre = {"Forte": 0, "Moyenne": 1, "Faible": 2}
    fiches = sorted(fiches, key=lambda f: (f.get("Dossier", ""), f.get("Fichier", "")))
    note_masque = ("Version masquée : extraits, aperçus de lignes et noms de colonnes supprimés."
                   if masquer else "Version complète : contient des extraits de données. Ne pas diffuser.")

    # Synthèse
    ws = wb.create_sheet("Synthèse")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "FPPN : inventaire des documents reçus du client"
    ws["A1"].font = Font(name=POLICE, size=16, bold=True)
    ws["A2"] = f"Généré le {datetime.now():%d/%m/%Y à %H:%M}. {note_masque}"
    ws["A2"].font = Font(name=POLICE, size=10, italic=True, color="53565A")
    stats = [("Nombre de fichiers", len(fiches)),
             ("Taille totale", taille_lisible(sum(f.get("Taille (octets)") or 0 for f in fiches))),
             ("Pertinence forte", sum(f.get("Pertinence") == "Forte" for f in fiches)),
             ("Pertinence moyenne", sum(f.get("Pertinence") == "Moyenne" for f in fiches)),
             ("Pertinence faible", sum(f.get("Pertinence") == "Faible" for f in fiches)),
             ("Scans", sum(f.get("Scan") == "Oui" for f in fiches)),
             ("Doublons", sum(bool(f.get("Doublon de")) for f in fiches)),
             ("Fichiers en erreur ou non lus", sum("rreur" in str(f.get("Remarque / erreur", "")) or
                                                    f.get("Type") in ("Non lu", "Autre") for f in fiches))]
    for i, (k, v) in enumerate(stats, 4):
        ws.cell(i, 1, k).font = Font(name=POLICE, size=11)
        c = ws.cell(i, 2, v)
        c.font = Font(name=POLICE, size=11, bold=True, color=BLEU_NUIT)
    r = len(stats) + 6
    ws.cell(r, 1, "Répartition par type").font = Font(name=POLICE, size=11, bold=True, color=ORANGE)
    for i, (t, n) in enumerate(Counter(f.get("Type", "") for f in fiches).most_common(), r + 1):
        ws.cell(i, 1, t or "(inconnu)").font = Font(name=POLICE, size=10)
        ws.cell(i, 2, n).font = Font(name=POLICE, size=10)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18

    # Inventaire
    lignes = []
    for n, f in enumerate(fiches, 1):
        ligne = []
        for c in COLONNES_INVENTAIRE:
            v = n if c == "N°" else f.get(c, "")
            if masquer and c == "Colonnes / en-têtes":
                v = "(masqué)" if v else ""
            ligne.append(v)
        lignes.append(ligne)
    feuille(wb, "Inventaire", COLONNES_INVENTAIRE, lignes, note_masque)

    # Résumés
    col_res = ["N°", "Dossier", "Fichier", "Pertinence", "Thème principal", "Résumé"] + ([] if masquer else ["Extrait"])
    lignes = []
    for n, f in enumerate(fiches, 1):
        resume = f.get("Résumé", "")
        if masquer:
            resume = re.sub(r"\(colonnes : [^)]*\)", "(colonnes masquées)", resume)
        ligne = [n, f.get("Dossier", ""), f.get("Fichier", ""), f.get("Pertinence", ""),
                 f.get("Thème principal", ""), resume]
        if not masquer:
            ligne.append(f.get("Extrait", ""))
        lignes.append(ligne)
    feuille(wb, "Résumés", col_res, lignes, note_masque)

    # Par thème
    lignes = []
    for theme in THEMES:
        liste = [(f, f["_scores"][theme]) for f in fiches if theme in f.get("_scores", {})]
        for f, s in sorted(liste, key=lambda x: -x[1]):
            lignes.append([theme, s, f.get("Pertinence", ""), f.get("Dossier", ""), f.get("Fichier", ""),
                           f.get("Type", ""), f.get("Années trouvées", "")])
    feuille(wb, "Par thème", ["Thème", "Score", "Pertinence", "Dossier", "Fichier", "Type", "Années trouvées"],
            lignes, "Score = nombre de mots-clés du thème trouvés dans le nom, les colonnes et le texte du fichier.")

    # Priorités
    prio = [f for f in fiches if f.get("Pertinence") in ("Forte", "Moyenne") and not f.get("Doublon de")]
    prio.sort(key=lambda f: (ordre.get(f.get("Pertinence"), 3), -sum(f.get("_scores", {}).values())))
    lignes = [[n, f.get("Pertinence", ""), f.get("Thème principal", ""), f.get("Dossier", ""),
               f.get("Fichier", ""), f.get("Type", ""), f.get("Lignes (max)", ""), f.get("Années trouvées", ""),
               f.get("Remarque / erreur", "")] for n, f in enumerate(prio, 1)]
    feuille(wb, "Priorités", ["N°", "Pertinence", "Thème principal", "Dossier", "Fichier", "Type",
                              "Lignes (max)", "Années trouvées", "Remarque / erreur"], lignes,
            "Fichiers à analyser en premier pour le volet actuariel (doublons exclus).")

    for w in wb.worksheets:
        w.sheet_properties.tabColor = ORANGE if w.title in ("Synthèse", "Priorités") else BLEU_NUIT
    wb.save(chemin_long(chemin))


if __name__ == "__main__":
    # Option : lancer avec un autre dossier source en argument
    #   python inventaire_fppn.py "D:\autre\dossier" "D:\sortie"
    if len(sys.argv) >= 2:
        DOSSIER_SOURCE = sys.argv[1]
    if len(sys.argv) >= 3:
        DOSSIER_SORTIE = sys.argv[2]
    lancer()
