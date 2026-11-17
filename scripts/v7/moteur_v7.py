# -*- coding: utf-8 -*-
"""Moteur de corrections du rapport FPPN v6 -> v7 (17/11/2026).
Chaque modification garde la mise en forme des mots inchangés. Les mots ajoutés ou changés passent en bleu 0070C0.
Le texte entre [[ ]] est surligné en jaune (à confirmer). Toute modification est consignée dans JOURNAL."""
import copy
import difflib
import re

from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
BLEU, JAUNE = "0070C0", "FFFF00"
FILLS = {A + t for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")}
APRES_FILL = [A + t for t in ("effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs",
                              "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")]
APRES_HL = APRES_FILL[3:]
JOURNAL = []
TOK = re.compile(r"\w+|\s+|[^\w\s]", re.UNICODE)


class Erreur(Exception):
    pass


def _insere(rpr, el, apres):
    for i, c in enumerate(rpr):
        if c.tag in apres:
            rpr.insert(i, el)
            return
    rpr.append(el)


def _bleu(rpr):
    for c in list(rpr):
        if c.tag in FILLS:
            rpr.remove(c)
    sf = etree.Element(A + "solidFill")
    etree.SubElement(sf, A + "srgbClr").set("val", BLEU)
    _insere(rpr, sf, APRES_FILL)


def _jaune(rpr, oui):
    for c in list(rpr):
        if c.tag == A + "highlight":
            rpr.remove(c)
    if oui:
        h = etree.Element(A + "highlight")
        etree.SubElement(h, A + "srgbClr").set("val", JAUNE)
        _insere(rpr, h, APRES_HL)


def _caracteres(p, strict=True):
    out = []
    for ch in p:
        if ch.tag == A + "r":
            rpr = ch.find(A + "rPr")
            t = ch.find(A + "t")
            for c in (t.text or "") if t is not None else "":
                out.append((c, rpr))
        elif ch.tag in (A + "br", A + "fld"):
            if strict:
                raise Erreur("saut de ligne ou champ dans le paragraphe")
            out.append(("", None))
    return out


def _marques(texte):
    """Retire les [[ ]] et renvoie (texte, masque de surlignage)."""
    plain, mask, hl = [], [], False
    i = 0
    while i < len(texte):
        if texte.startswith("[[", i):
            hl, i = True, i + 2
            continue
        if texte.startswith("]]", i):
            hl, i = False, i + 2
            continue
        plain.append(texte[i])
        mask.append(hl)
        i += 1
    return "".join(plain), mask


def texte_para(p):
    return "".join(c for c, _ in _caracteres(p, False))


def modifie(p, nouveau, tout_bleu=False, lieu=""):
    """Remplace le texte du paragraphe p par `nouveau` (avec [[ ]] possibles)."""
    anc = _caracteres(p)
    if not anc:
        raise Erreur("paragraphe vide " + lieu)
    ancien = "".join(c for c, _ in anc)
    neuf, mask = _marques(nouveau)
    if neuf == ancien and not any(mask):
        return False
    base0 = anc[0][1]
    res = []  # (car, rpr_source, change, surligne)
    if tout_bleu:
        # paragraphe réécrit : si l'original commençait en gras, seule la première phrase (la conclusion) reste en gras
        maigre = base0
        if base0 is not None and base0.get("b") == "1" and len(neuf) > 120:
            maigre = copy.deepcopy(base0)
            maigre.attrib.pop("b")
        fin1 = neuf.find(". ") + 1 if ". " in neuf else len(neuf)
        for k, c in enumerate(neuf):
            res.append((c, base0 if k < fin1 else maigre, True, mask[k]))
    else:
        ta, tn = TOK.findall(ancien), TOK.findall(neuf)
        pa, pn = [0], [0]
        for t in ta:
            pa.append(pa[-1] + len(t))
        for t in tn:
            pn.append(pn[-1] + len(t))
        sm = difflib.SequenceMatcher(None, ta, tn, autojunk=False)
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == "equal":
                for k in range(pa[i2] - pa[i1]):
                    c, r = anc[pa[i1] + k]
                    res.append((c, r, False, mask[pn[j1] + k]))
            elif op in ("replace", "insert"):
                src = anc[min(pa[i1], len(anc) - 1)][1] if pa[i1] < len(anc) else anc[-1][1]
                if op == "insert" and pa[i1] > 0:
                    src = anc[pa[i1] - 1][1]
                for k in range(pn[j1], pn[j2]):
                    res.append((neuf[k], src, True, mask[k]))
    # regroupe en runs
    for ch in list(p):
        if ch.tag == A + "r":
            p.remove(ch)
    end = p.find(A + "endParaRPr")
    pos = list(p).index(end) if end is not None else len(p)
    runs, cle_prec = [], None
    for c, src, change, hl in res:
        rpr = copy.deepcopy(src) if src is not None else etree.Element(A + "rPr")
        for at in ("err", "dirty"):
            rpr.attrib.pop(at, None)
        if change:
            _bleu(rpr)
            _jaune(rpr, hl)
        elif hl:
            _jaune(rpr, True)
        cle = etree.tostring(rpr)
        if runs and cle == cle_prec:
            runs[-1][1].append(c)
        else:
            runs.append([rpr, [c]])
            cle_prec = cle
    for k, (rpr, cs) in enumerate(runs):
        r = etree.Element(A + "r")
        r.append(rpr)
        t = etree.SubElement(r, A + "t")
        t.text = "".join(cs)
        p.insert(pos + k, r)
    JOURNAL.append((lieu, ancien, neuf))
    return True


class Deck:
    def __init__(self, prs):
        self.prs = prs
        self.modifiees = set()

    def slide(self, n):
        return self.prs.slides[n - 1]

    def forme(self, n, nom):
        for sh in self.slide(n).shapes:
            if sh.name == nom:
                return sh
        raise Erreur(f"D{n} : forme {nom} absente")

    def paras(self, n, noms=None, tables=True):
        """Tous les paragraphes de la diapositive : (lieu, nom de forme, a:p)."""
        for sh in self.slide(n).shapes:
            if noms and sh.name not in noms:
                continue
            if sh.has_text_frame:
                for i, pa in enumerate(sh.text_frame.paragraphs):
                    yield f"D{n} {sh.name}#{i}", sh.name, pa._p
            elif sh.has_table and tables:
                for ri, row in enumerate(sh.table.rows):
                    for ci, c in enumerate(row.cells):
                        for i, pa in enumerate(c.text_frame.paragraphs):
                            yield f"D{n} {sh.name}[{ri},{ci}]#{i}", sh.name, pa._p

    def para(self, n, nom, i, debut):
        p = self.forme(n, nom).text_frame.paragraphs[i]._p
        t = texte_para(p)
        if not t.startswith(debut):
            raise Erreur(f"D{n} {nom}#{i} commence par « {t[:60]} », attendu « {debut} »")
        return p

    def reecrit(self, n, nom, i, debut, texte, tout_bleu=True):
        modifie(self.para(n, nom, i, debut), texte, tout_bleu, f"D{n} {nom}#{i}")
        self.modifiees.add(n)

    def rep(self, n, ancien, nouveau, noms=None, mini=1):
        k = 0
        for lieu, _, p in list(self.paras(n, noms)):
            t = texte_para(p)
            if ancien in t and nouveau not in t:
                modifie(p, t.replace(ancien, nouveau), False, lieu)
                k += 1
        if k < mini:
            raise Erreur(f"D{n} : « {ancien} » introuvable")
        if k:
            self.modifiees.add(n)
        return k

    def soit(self, n, fragment, suite, noms=None):
        return self.rep(n, fragment, fragment + suite, noms)

    def ajoute(self, n, nom, apres, modele, texte):
        """Ajoute un paragraphe après l'indice `apres`, avec la mise en forme du paragraphe `modele`."""
        tf = self.forme(n, nom).text_frame
        src = tf.paragraphs[modele]._p
        neuf = copy.deepcopy(src)
        tf.paragraphs[apres]._p.addnext(neuf)
        modifie(neuf, texte, True, f"D{n} {nom}#{apres + 1} (ajout)")
        self.modifiees.add(n)
