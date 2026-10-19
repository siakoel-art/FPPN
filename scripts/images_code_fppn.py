# -*- coding: utf-8 -*-
"""
FPPN - Images d'extraits de code mis en forme comme dans VS Code sous Windows (pour le rapport)
==============================================================================================

Chaque image reproduit la fenêtre de VS Code sous Windows telle qu'elle apparaît sur le poste de la mission
(thème sombre, barre de menus en anglais, fil d'Ariane C:\\Dev\\Missions\\05-FPPN\\05-Analyses\\scripts,
barre d'état Python 3.14) et y place un extrait réel d'un script de la mission, avec ses numéros de ligne.
Le terminal n'est affiché que pour une sortie réellement obtenue. Ce sont des reconstitutions, pas des captures.
Utilisation : python images_code_fppn.py <dossier_scripts> <dossier_sortie>
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont
from pygments import lex
from pygments.lexers import PythonLexer
from pygments.token import Comment, Keyword, Name, Number, Operator, String, Token

SCRIPTS, SORTIE = sys.argv[1], sys.argv[2]
os.makedirs(SORTIE, exist_ok=True)
MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"
UI = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
CHEMIN = ["C:", "Dev", "Missions", "05-FPPN", "05-Analyses", "scripts"]

# Couleurs du thème Dark Modern de VS Code
FOND, BARRE, BORD, TXT_UI, NUM, SEL = "#1f1f1f", "#181818", "#2b2b2b", "#cccccc", "#6e7681", "#2a2d2e"
COUL = [(Comment, "#6a9955"), (String.Doc, "#ce9178"), (String, "#ce9178"), (Number, "#b5cea8"),
        (Keyword.Constant, "#569cd6"), (Keyword, "#c586c0"), (Name.Builtin, "#dcdcaa"), (Name.Function, "#dcdcaa"),
        (Name.Class, "#4ec9b0"), (Operator.Word, "#569cd6"), (Name, "#9cdcfe"), (Operator, "#d4d4d4")]

# (nom, fichier, début de l'extrait, nombre de lignes, terminal)
EXTRAITS = [
    ("code_extraction", "extraction_fppn.py", "def format_tableur(donnees):", 12,
     ["PS C:\\Dev\\Missions\\05-FPPN\\05-Analyses> python scripts\\extraction_fppn.py",
      "Terminé en 1.6 minutes.", "Archive compressée : 91.5 Mo, en 5 partie(s).",
      "Fichiers à envoyer (toutes les parties et LISEZMOI_parties.txt) : C:\\Dev\\Missions\\05-FPPN\\05-Analyses\\Extraction_FPPN_20261018_1939\\A_ENVOYER"]),
    ("code_total", "base_fppn.py", "        # Ligne de total : montant égal à la somme des autres lignes", 13, None),
    ("code_double", "base_fppn.py", "    # Montants saisis deux fois à la suite dans le fichier source", 16, None),
    ("code_grenier", "grenier_fppn.py", "def montant(v):", 9, None),
    ("code_sc", "analyses_fppn.py", "s = []\nfor ex in EX:", 22, None),
    ("code_stats", "analyses_complements_fppn.py", "def wilson(k, n, z=1.959964):", 7, None),
    ("code_statut", "ratios_airms_fppn.py", "def statut(test, v, ex):", 21,
     ["PS C:\\Dev\\Missions\\05-FPPN\\05-Analyses> python scripts\\ratios_airms_fppn.py base resultats\\resultats.json resultats",
      "36 ratios ; 246 comptes classés"]),
]


def couleur(tt):
    for t, c in COUL:
        if tt in t:
            return c
    return "#d4d4d4"


def extrait(fichier, debut, n):
    texte = open(os.path.join(SCRIPTS, fichier), encoding="utf-8").read()
    pos = texte.index(debut)
    i = texte[:pos].count("\n")
    lignes = texte.split("\n")
    return lignes[i:i + n], i + 1


def icone_python(d, x, y, t=14):
    d.rounded_rectangle([x, y, x + t * 0.62, y + t * 0.62], 3, fill="#3572a5")
    d.rounded_rectangle([x + t * 0.38, y + t * 0.38, x + t, y + t], 3, fill="#ffd43b")


def image(nom, fichier, debut, n, terminal):
    lignes, premiere = extrait(fichier, debut, n)
    fm, fu, fus = ImageFont.truetype(MONO, 22), ImageFont.truetype(UI, 21), ImageFont.truetype(UI, 19)
    W, hl = 1900, 33
    h_titre, h_onglet, h_ariane, h_etat = 46, 48, 36, 34
    h_term = (60 + 30 * len(terminal)) if terminal else 0
    h_code = 20 + hl * len(lignes) + 16
    H = h_titre + h_onglet + h_ariane + h_code + h_term + h_etat
    img = Image.new("RGB", (W, H), FOND)
    d = ImageDraw.Draw(img)
    # Barre de titre et menus (Windows)
    d.rectangle([0, 0, W, h_titre], fill=BARRE)
    d.polygon([(14, 23), (30, 9), (36, 13), (36, 33), (30, 37), (14, 23)], fill="#23a9f2")
    d.polygon([(18, 23), (30, 31), (30, 15)], fill=BARRE)
    x = 60
    for m in ("File", "Edit", "Selection", "View", "Go", "Run", "Terminal", "Help"):
        d.text((x, 12), m, fill=TXT_UI, font=fu)
        x += d.textlength(m, font=fu) + 26
    fsym = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
    d.text((x + 20, 10), "\u2190   \u2192", fill="#8b8b8b", font=fsym)
    d.rounded_rectangle([700, 8, 1340, 38], 6, fill="#2a2a2a", outline=BORD)
    d.text((714, 13), "Search", fill="#8b8b8b", font=fus)
    for k, xx in enumerate((1600, 1640, 1680, 1720)):
        d.rectangle([xx, 15, xx + 22, 33], outline="#b0b0b0", width=2)
    d.line([1770, 24, 1790, 24], fill=TXT_UI, width=2)
    d.rectangle([1815, 15, 1832, 32], outline=TXT_UI, width=2)
    d.line([1860, 15, 1878, 33], fill=TXT_UI, width=2)
    d.line([1878, 15, 1860, 33], fill=TXT_UI, width=2)
    # Barre d'activités (gauche)
    ya = h_titre
    d.rectangle([0, ya, 60, H - h_etat], fill=BARRE)
    for k in range(7):
        yy = ya + 22 + k * 56
        if k == 0:
            d.rectangle([0, yy - 8, 3, yy + 32], fill="#0078d4")
            d.rectangle([20, yy, 40, yy + 26], outline="#d7d7d7", width=2)
        elif k == 1:
            d.ellipse([20, yy, 36, yy + 16], outline="#868686", width=2)
            d.line([34, yy + 15, 42, yy + 24], fill="#868686", width=3)
        elif k == 2:
            for (cx, cy) in ((24, yy + 4), (24, yy + 24), (38, yy + 12)):
                d.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], outline="#868686", width=2)
            d.line([24, yy + 8, 24, yy + 20], fill="#868686", width=2)
        elif k == 3:
            d.polygon([(22, yy), (40, yy + 13), (22, yy + 26)], outline="#868686")
        elif k == 4:
            for (bx, by) in ((20, yy + 8), (30, yy + 8), (20, yy + 18), (32, yy - 2)):
                d.rectangle([bx, by, bx + 8, by + 8], outline="#868686", width=2)
        elif k == 5:
            d.polygon([(26, yy), (34, yy), (34, yy + 10), (42, yy + 26), (18, yy + 26), (26, yy + 10)], outline="#868686")
        else:
            icone_python(d, 20, yy + 2, 22)
    # Onglet du fichier
    yo = h_titre
    d.rectangle([60, yo, W, yo + h_onglet], fill=BARRE)
    lo = 60 + 46 + d.textlength(fichier, font=fu) + 50
    d.rectangle([60, yo, lo, yo + h_onglet], fill=FOND)
    d.line([60, yo, lo, yo], fill="#0078d4", width=2)
    icone_python(d, 76, yo + 17)
    d.text((100, yo + 13), fichier, fill="#ffffff", font=fu)
    d.text((lo - 30, yo + 12), "\u00d7", fill=TXT_UI, font=fu)
    d.polygon([(W - 108, yo + 16), (W - 94, yo + 24), (W - 108, yo + 32)], outline=TXT_UI)
    d.rectangle([W - 76, yo + 15, W - 58, yo + 33], outline=TXT_UI, width=2)
    d.text((W - 42, yo + 8), "...", fill=TXT_UI, font=fu)
    # Fil d'Ariane
    yb = yo + h_onglet
    x = 80
    for k, p in enumerate(CHEMIN + [fichier, "..."]):
        if p == fichier:
            icone_python(d, x, yb + 10, 13)
            x += 20
        d.text((x, yb + 7), p, fill="#a9a9a9", font=fus)
        x += d.textlength(p, font=fus)
        if k < len(CHEMIN) + 1:
            d.text((x + 6, yb + 5), "\u203a", fill="#a9a9a9", font=fu)
            x += 26
    # Code
    yc = yb + h_ariane + 10
    d.rectangle([60, yc - 6, W - 120, yc + hl - 6], fill=SEL)
    lw = d.textlength(str(premiere + len(lignes)), font=fm)
    for k, ligne in enumerate(lignes):
        yy = yc + k * hl
        num = str(premiere + k)
        d.text((60 + 30 + lw - d.textlength(num, font=fm), yy), num, fill="#cccccc" if k == 0 else NUM, font=fm)
        xx = 60 + 30 + lw + 40
        for tt, val in lex(ligne, PythonLexer()):
            val = val.rstrip("\n")
            if not val:
                continue
            d.text((xx, yy), val, fill=couleur(tt), font=fm)
            xx += d.textlength(val, font=fm)
    # Minimap
    for k in range(min(len(lignes), 40)):
        d.line([W - 105, yc + k * 4, W - 105 + min(80, len(lignes[k]) * 1.2), yc + k * 4], fill="#4a4a4a", width=2)
    # Terminal (sortie réellement obtenue)
    if terminal:
        yt = yc + h_code - 10
        d.line([60, yt, W, yt], fill=BORD, width=1)
        xx = 80
        for k, t in enumerate(("Problems", "Output", "Debug Console", "Terminal", "Ports")):
            if t == "Terminal":
                d.rectangle([xx - 8, yt + 8, xx + d.textlength(t, font=fus) + 8, yt + 36], fill="#2a2a2a")
            d.text((xx, yt + 12), t, fill="#ffffff" if t == "Terminal" else "#9d9d9d", font=fus)
            xx += d.textlength(t, font=fus) + 34
        for k, t in enumerate(terminal):
            d.text((90, yt + 52 + k * 30), t, fill="#cccccc", font=ImageFont.truetype(MONO, 19))
    # Barre d'état
    ye = H - h_etat
    d.rectangle([0, ye, W, H], fill=BARRE)
    d.line([0, ye, W, ye], fill=BORD)
    d.ellipse([72, ye + 9, 88, ye + 25], outline="#a9a9a9", width=2)
    d.text((94, ye + 7), "0", fill="#a9a9a9", font=fus)
    d.polygon([(122, ye + 25), (131, ye + 9), (140, ye + 25)], outline="#a9a9a9")
    d.text((146, ye + 7), "0", fill="#a9a9a9", font=fus)
    etat = "Ln 1, Col 1     Spaces: 4     UTF-8     LF     { } Python     Python 3.14 (64-bit)"
    d.text((W - d.textlength(etat, font=fus) - 40, ye + 7), etat, fill="#a9a9a9", font=fus)
    img.save(os.path.join(SORTIE, nom + ".png"))
    return nom, premiere, premiere + len(lignes) - 1


if __name__ == "__main__":
    for e in EXTRAITS:
        print(image(*e))
