# -*- coding: utf-8 -*-
"""Retire des livrables les attributions à une personne (fonction individuelle) au profit de la direction ou du service.

« selon le chef comptable » devient « selon la direction comptable », « le directeur général décide seul » devient
« la direction générale décide seule », « le responsable des cotisations » devient « le service des cotisations ».
Les références à un acte sur le poste lui-même (désignation, mandat du directeur général) sont conservées.

Usage en module : from depersonnaliser import dep ; en ligne de commande : python3 depersonnaliser.py fichier.py|.xlsx …
"""
import re
import sys

REGLES = [
    # personnes nommées
    (r"\bla demande de M\. Wabi\b", "les termes de la mission"),
    (r"\bdemande de M\. Wabi \(associé responsable\)", "termes de la mission"),
    (r"\bdemande de M\. Wabi\b", "termes de la mission"),
    # direction comptable
    (r"\b([Ss])elon le chef comptable\b", r"\1elon la direction comptable"),
    (r"\bpar le chef comptable\b", "par la direction comptable"),
    (r"\bavec le chef comptable\b", "avec la direction comptable"),
    (r"\bdu chef comptable\b", "de la direction comptable"),
    (r"\bau chef comptable\b", "à la direction comptable"),
    (r"\bLe chef comptable\b", "La direction comptable"),
    (r"\ble chef comptable\b", "la direction comptable"),
    (r"\bChef comptable\b", "Direction comptable"),
    (r"\bchef comptable\b", "direction comptable"),
    # service des cotisations, service des sinistres et prestations
    (r"\b([Ll]e|[Dd]u|[Aa]u|[Pp]ar le|[Aa]vec le) responsable des (cotisations|sinistres et prestations)\b", lambda m: {
        "le": "le", "Le": "Le", "du": "du", "Du": "Du", "au": "au", "Au": "Au"}.get(m.group(1), m.group(1)) + " service des " + m.group(2)),
    (r"\bResponsable des (cotisations|sinistres et prestations)\b", r"Service des \1"),
    (r"\bresponsable des (cotisations|sinistres et prestations)\b", r"service des \1"),
    # direction générale : seulement les décisions et signatures attribuées à la personne
    (r"\b[Pp]lacements décidés par le seul directeur général\b", "Placements décidés par la direction générale"),
    (r"\bpar le seul directeur général\b", "par la seule direction générale"),
    (r"\bLe directeur général signe\b", "La direction générale signe"),
    (r"\ble directeur général signe\b", "la direction générale signe"),
    (r"\(chef comptable, ", "(direction comptable, "),
    (r"directeur administratif et financier", "direction administrative et financière"),
    (r"\bChef de service ([Aa]ssurance)", r"Service \1"),
    (r"\bchef de service ([Aa]ssurance)", r"service \1"),
    (r"\bLe directeur général décide seul\b", "La direction générale décide seule"),
    (r"\ble directeur général décide seul\b", "la direction générale décide seule"),
    (r"\bLe directeur général décide\b", "La direction générale décide"),
    (r"\ble directeur général décide\b", "la direction générale décide"),
    (r"\bLe directeur général la signe\b", "La direction générale la signe"),
    (r"\bsignée par le directeur général\b", "signée par la direction générale"),
]
_C = [(re.compile(a), b) for a, b in REGLES]


def dep(texte):
    if not isinstance(texte, str):
        return texte
    for motif, rempl in _C:
        texte = motif.sub(rempl, texte)
    return texte


def xlsx(chemin):
    import openpyxl
    wb = openpyxl.load_workbook(chemin)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    v = dep(c.value)
                    if v != c.value:
                        c.value, n = v, n + 1
                if c.comment is not None:
                    v = dep(c.comment.text)
                    if v != c.comment.text:
                        c.comment.text, n = v, n + 1
    wb.save(chemin)
    return n


if __name__ == "__main__":
    for f in sys.argv[1:]:
        if f.endswith(".xlsx"):
            print(f, xlsx(f), "cellules")
        else:
            s = open(f, encoding="utf-8").read()
            t = dep(s)
            open(f, "w", encoding="utf-8").write(t)
            print(f, sum(1 for a, b in zip(s.splitlines(), t.splitlines()) if a != b), "lignes")
