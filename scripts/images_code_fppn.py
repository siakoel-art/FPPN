# -*- coding: utf-8 -*-
"""
FPPN - Images d'extraits de code mis en forme comme dans VS Code (pour le rapport)
==================================================================================

Chaque image reproduit un extrait réel d'un script de la mission : onglet du fichier, numéros de ligne,
coloration syntaxique (thème sombre). Ce sont des extraits mis en forme, pas des captures d'écran.
Utilisation : python images_code_fppn.py <dossier_scripts> <dossier_sortie>
"""
import io
import os
import sys

from PIL import Image, ImageDraw, ImageFont
from pygments import highlight
from pygments.formatters import ImageFormatter
from pygments.lexers import PythonLexer

SCRIPTS, SORTIE = sys.argv[1], sys.argv[2]
os.makedirs(SORTIE, exist_ok=True)
POLICE = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

EXTRAITS = [
    ("code_extraction", "extraction_fppn.py", "def format_tableur(donnees):", 12),
    ("code_total", "base_fppn.py", "        # Ligne de total : montant égal à la somme des autres lignes", 13),
    ("code_double", "base_fppn.py", "    # Montants saisis deux fois à la suite dans le fichier source", 16),
    ("code_grenier", "grenier_fppn.py", "def montant(v):", 9),
    ("code_sc", "analyses_fppn.py", "s = []\nfor ex in EX:", 22),
    ("code_stats", "analyses_complements_fppn.py", "def wilson(k, n, z=1.959964):", 7),
]


def extrait(fichier, debut, n):
    lignes = open(os.path.join(SCRIPTS, fichier), encoding="utf-8").read().split("\n")
    texte = "\n".join(lignes)
    pos = texte.index(debut)
    i = texte[:pos].count("\n")
    return "\n".join(lignes[i:i + n]), i + 1


def image(nom, fichier, debut, n):
    code, premiere = extrait(fichier, debut, n)
    fmt = ImageFormatter(style="one-dark", font_name="DejaVu Sans Mono", font_size=22, line_numbers=True,
                         line_number_start=premiere, line_number_bg="#282c34", line_number_fg="#636d83",
                         line_pad=6, image_pad=18)
    png = highlight(code, PythonLexer(), fmt)
    corps = Image.open(io.BytesIO(png)).convert("RGB")
    w = max(corps.width, 1500)
    barre_titre, onglets = 44, 46
    img = Image.new("RGB", (w, corps.height + barre_titre + onglets), "#282c34")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, barre_titre], fill="#21252b")
    for k, c in enumerate(("#ff5f56", "#ffbd2e", "#27c93f")):
        d.ellipse([16 + k * 28, 14, 32 + k * 28, 30], fill=c)
    f = ImageFont.truetype(POLICE, 18)
    d.text((w // 2 - 230, 12), f"{fichier} - FPPN - Visual Studio Code", fill="#9da5b4", font=f)
    d.rectangle([0, barre_titre, w, barre_titre + onglets], fill="#21252b")
    d.rectangle([0, barre_titre, 26 + len(fichier) * 11 + 40, barre_titre + onglets], fill="#282c34")
    d.text((18, barre_titre + 12), fichier, fill="#e5c07b", font=ImageFont.truetype(POLICE, 19))
    img.paste(corps, (0, barre_titre + onglets))
    img.save(os.path.join(SORTIE, nom + ".png"))
    return nom, premiere, premiere + n - 1


if __name__ == "__main__":
    for e in EXTRAITS:
        print(image(*e))
