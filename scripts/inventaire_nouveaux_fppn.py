# -*- coding: utf-8 -*-
"""
FPPN - Inventaire rapide des nouveaux fichiers reçus (22/10/2026)
================================================================
Liste chaque fichier des dossiers indiqués : dossier, nom, taille, date, empreinte SHA-256 et, selon le type,
feuilles Excel (nombre de lignes et de colonnes, trois premières lignes), nombre de pages PDF, première ligne des
fichiers texte. Ne modifie, ne déplace et n'envoie aucun fichier. Résultat : un classeur Excel léger, à transmettre.
Utilisation (VS Code ou terminal Windows) : python inventaire_nouveaux_fppn.py
"""
# ============================================================================= PARAMÈTRES
RACINE = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client"
ELEMENTS = [  # dossiers et fichiers reçus le 22/10/2026
    r"07-Nouveaux Trésorerie et placements",
    r"COURRIER DE NOMINATION DÉFINITIVE.pdf",
    r"ORGANIGRAMME FPPN 2021(1).pdf",
    r"02-Nouveaux_GL et Balances",
    r"03-Nouveaux Juridiques",
    r"04-Nouveaux Personnel et paie",
    r"05-Nouveaux Prestations",
    r"06-Nouveaux Revue Budgétaire",
]
SORTIE = r"C:\Dev\Missions\05-FPPN\05-Analyses\FPPN_inventaire_nouveaux_fichiers_20261022.xlsx"
# =============================================================================
import datetime as dt
import hashlib
import importlib
import os
import subprocess
import sys

for paquet, module in (("openpyxl", "openpyxl"), ("python-calamine", "python_calamine"), ("pypdf", "pypdf")):
    try:
        importlib.import_module(module)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", paquet])
import openpyxl  # noqa: E402
from python_calamine import CalamineWorkbook  # noqa: E402
from pypdf import PdfReader  # noqa: E402


def empreinte(chemin):
    h = hashlib.sha256()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1 << 20), b""):
            h.update(bloc)
    return h.hexdigest()


def apercu(lignes):
    return " || ".join(" | ".join(str(v)[:30] for v in l if v not in (None, ""))[:250] for l in lignes[:3])


fichiers = []
for el in ELEMENTS:
    p = os.path.join(RACINE, el)
    if os.path.isfile(p):
        fichiers.append(p)
    elif os.path.isdir(p):
        for d, _, noms in os.walk(p):
            fichiers += [os.path.join(d, n) for n in sorted(noms)]
    else:
        print("INTROUVABLE :", p)

lignes_f, lignes_s = [], []
for i, f in enumerate(fichiers, 1):
    ext = os.path.splitext(f)[1].lower()
    rel = os.path.relpath(os.path.dirname(f), RACINE)
    info = {"n": i, "dossier": rel, "fichier": os.path.basename(f), "extension": ext, "taille_octets": os.path.getsize(f),
            "modifie_le": dt.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%d/%m/%Y %H:%M"), "sha256": empreinte(f),
            "feuilles_ou_pages": None, "remarque": ""}
    print(f"[{i}/{len(fichiers)}] {rel}\\{info['fichier']}")
    try:
        if ext in (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"):
            wb = CalamineWorkbook.from_path(f)
            info["feuilles_ou_pages"] = len(wb.sheet_names)
            for nom in wb.sheet_names:
                t = wb.get_sheet_by_name(nom).to_python()
                lignes_s.append({"n": i, "fichier": info["fichier"], "feuille": nom, "lignes": len(t),
                                 "colonnes": max((len(l) for l in t), default=0), "apercu_3_premieres_lignes": apercu(t)})
        elif ext == ".pdf":
            r = PdfReader(f)
            info["feuilles_ou_pages"] = len(r.pages)
            texte = (r.pages[0].extract_text() or "").strip() if r.pages else ""
            info["remarque"] = "PDF texte : " + texte[:200].replace("\n", " ") if texte else "PDF scanné (pas de texte) : à me transmettre"
        elif ext in (".csv", ".txt"):
            with open(f, "r", encoding="utf-8", errors="replace") as h:
                info["remarque"] = "Première ligne : " + h.readline()[:200].strip()
    except Exception as e:  # fichier illisible : signalé, jamais bloquant
        info["remarque"] = f"Illisible : {type(e).__name__} {str(e)[:120]}"
    lignes_f.append(info)

wb = openpyxl.Workbook()
for titre, lignes in (("Fichiers", lignes_f), ("Feuilles Excel", lignes_s)):
    ws = wb.active if titre == "Fichiers" else wb.create_sheet(titre)
    ws.title = titre
    if lignes:
        ws.append(list(lignes[0].keys()))
        for l in lignes:
            ws.append(list(l.values()))
os.makedirs(os.path.dirname(SORTIE), exist_ok=True)
wb.save(SORTIE)
print(f"\n{len(lignes_f)} fichiers, {len(lignes_s)} feuilles Excel, {sum(l['taille_octets'] for l in lignes_f) / 1e6:.1f} Mo")
print("Inventaire écrit :", SORTIE)
