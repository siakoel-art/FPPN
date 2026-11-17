# -*- coding: utf-8 -*-
"""Affiche, pour une liste de diapositives, chaque paragraphe des zones de texte avec taille, couleur et surlignage.
Utilisation : python dump_v6.py <fichier.pptx> 9 80 105"""
import sys

from pptx import Presentation

p = Presentation(sys.argv[1])
for n in map(int, sys.argv[2:]):
    s = p.slides[n - 1]
    print("=" * 20, n)
    for sh in s.shapes:
        if sh.has_text_frame:
            for i, para in enumerate(sh.text_frame.paragraphs):
                if not para.text.strip():
                    continue
                r = para.runs[0] if para.runs else None
                sz = r.font.size.pt if r is not None and r.font.size else None
                col = None
                try:
                    col = r.font.color.rgb if r is not None and r.font.color and r.font.color.type else None
                except Exception:
                    pass
                hl = "HL" if any("highlight" in rr._r.xml for rr in para.runs) else ""
                print(f"[{sh.name}#{i}] sz={sz} col={col} {hl} nruns={len(para.runs)} | {para.text}")
        elif sh.has_table:
            tb = sh.table
            for ri, row in enumerate(tb.rows):
                print(f"[{sh.name} r{ri}] " + " || ".join(c.text for c in row.cells))
