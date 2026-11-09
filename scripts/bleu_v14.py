# -*- coding: utf-8 -*-
"""
FPPN - Mise en bleu des ajouts et corrections du 10/11/2026 dans les livrables déjà produits (PowerPoint et Word).
Le texte est repéré par des phrases exactes (au moins 15 caractères) ; les runs sont découpés pour ne colorer que la phrase.
Utilisation en module : bleuir_pptx(chemin, phrases, titres_entiers), bleuir_docx(chemin, phrases).
"""
from copy import deepcopy

BLEU = "0070C0"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _marques(plein, phrases):
    m = [False] * len(plein)
    for ph in phrases:
        d = plein.find(ph)
        while d >= 0:
            for j in range(d, d + len(ph)):
                m[j] = True
            d = plein.find(ph, d + len(ph))
    return m


def _decouper(runs, textes, marques, colorer, poser_texte):
    pos = 0
    for r, tx in zip(runs, textes):
        m = marques[pos:pos + len(tx)]
        pos += len(tx)
        if not tx or not any(m):
            continue
        if all(m):
            colorer(r)
            continue
        morceaux, cour, deb = [], m[0], 0
        for j in range(1, len(tx) + 1):
            if j == len(tx) or m[j] != cour:
                morceaux.append((tx[deb:j], cour))
                if j < len(tx):
                    cour, deb = m[j], j
        ancre = r
        for t, b in morceaux:
            nr = deepcopy(r)
            poser_texte(nr, t)
            if b:
                colorer(nr)
            ancre.addnext(nr)
            ancre = nr
        r.getparent().remove(r)


# ------------------------------------------------------------------ PowerPoint
def _couleur_a(r):
    from lxml import etree
    rpr = r.find(_A + "rPr")
    if rpr is None:
        rpr = etree.Element(_A + "rPr")
        r.insert(0, rpr)
    sf = rpr.find(_A + "solidFill")
    if sf is not None:
        c0 = sf.find(_A + "srgbClr")
        if c0 is not None and c0.get("val") == "FFFFFF":
            return
        rpr.remove(sf)
    sf = etree.Element(_A + "solidFill")
    etree.SubElement(sf, _A + "srgbClr").set("val", BLEU)
    ln = rpr.find(_A + "ln")
    if ln is not None:
        ln.addnext(sf)
    else:
        rpr.insert(0, sf)


def _texte_a(r, t):
    r.find(_A + "t").text = t


def bleuir_pptx(chemin, phrases, titres_entiers=(), sortie=None):
    from pptx import Presentation
    prs = Presentation(chemin)
    phr = sorted({p for p in phrases if len(p) >= 15}, key=len, reverse=True)
    n = 0
    for sl in prs.slides:
        titre = " ".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name in ("Adenka_Titre", "Title 1", "Titre 1"))
        entiere = any(t and t == titre.strip() for t in titres_entiers)
        for p in sl.element.iter(_A + "p"):
            runs = [r for r in p if r.tag == _A + "r"]
            if not runs:
                continue
            textes = [r.findtext(_A + "t") or "" for r in runs]
            if entiere:
                if "Empreinte Africaine" not in "".join(textes):
                    for r in runs:
                        _couleur_a(r)
                    n += 1
                continue
            m = _marques("".join(textes), phr)
            if any(m):
                _decouper(runs, textes, m, _couleur_a, _texte_a)
                n += 1
    prs.save(sortie or chemin)
    return n


# ------------------------------------------------------------------ Word
def _couleur_w(r):
    from lxml import etree
    rpr = r.find(_W + "rPr")
    if rpr is None:
        rpr = etree.Element(_W + "rPr")
        r.insert(0, rpr)
    for c in rpr.findall(_W + "color"):
        rpr.remove(c)
    c = etree.Element(_W + "color")
    c.set(_W + "val", BLEU)
    # ordre du schéma : la couleur vient après les attributs de police, avant la taille
    apres = [e for e in rpr if e.tag in {_W + n for n in ("spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")}]
    if apres:
        apres[0].addprevious(c)
    else:
        rpr.append(c)


def _texte_w(r, t):
    tt = r.find(_W + "t")
    tt.text = t
    tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def bleuir_docx(chemin, phrases, sortie=None):
    import docx
    d = docx.Document(chemin)
    phr = sorted({p for p in phrases if len(p) >= 15}, key=len, reverse=True)
    n = 0
    for p in d.element.body.iter(_W + "p"):
        runs = [r for r in p if r.tag == _W + "r" and len(r.findall(_W + "t")) == 1]
        if not runs:
            continue
        textes = [r.findtext(_W + "t") or "" for r in runs]
        m = _marques("".join(textes), phr)
        if any(m):
            _decouper(runs, textes, m, _couleur_w, _texte_w)
            n += 1
    d.save(sortie or chemin)
    return n
