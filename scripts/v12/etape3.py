# ---- v13, étape 3 : logigrammes par produit, matrices, légende complète (exécuté par patch_v12.py)
# dessin des quatre circuits par produit
rep('''            elif x["logigramme"] == "facture":
                LG.circuit_facture(sl, 4.0)
                y_ = 10.6''', '''            elif x["logigramme"] == "facture":
                LG.circuit_facture(sl, 4.0)
                y_ = 10.6
            elif x["logigramme"] in ("epargne", "gna", "maladie_deces"):
                getattr(LG, "circuit_" + x["logigramme"])(sl, 3.9)
                y_ = 14.4
            elif x["logigramme"] == "cgrae":
                LG.circuit_cgrae(sl, 3.9)
                y_ = 14.3''')
# cases colorées des matrices (risques, criticité) et alignement centré
rep('''            if str(v) in BOULES:
                p.alignment = PP_ALIGN.CENTER
''', '''            if str(v) in BOULES or (spec.get("centre") and j > 0 and i > 0):
                p.alignment = PP_ALIGN.CENTER
            if i > 0 and spec.get("fonds") and (i - 1, j) in spec["fonds"]:
                fond = spec["fonds"][(i - 1, j)]
                coul, gras_c = ("FFFFFF" if fond == "C00000" else NOIR), True
''')
# légende complète de l'échelle de conformité
rep('''LEG_BORD = [("🔴", "Non conforme", "la valeur ne respecte pas la norme"), ("🟢", "Conforme", "la valeur respecte la norme"),
            ("⚪", "À apprécier", "norme non chiffrée"), ("🟡", "En attente", "donnée du FPPN attendue")]''', "LEG_BORD = LEG_CONF")
rep('"legende": [("🔴", "Non conforme", "norme non respectée cette année-là"), ("🟢", "Conforme", "norme respectée cette année-là")],', '"legende": LEG_CONF,')
# nouvelles diapositives
rep('contenu({"titre": "Flux financiers 2025"', bloc("bloc_circuits.py") + '\ncontenu({"titre": "Flux financiers 2025"')
rep("# ---- v12 : indicateurs clés de 2020 au 31/08/2026", bloc("bloc_matrice_criticite.py") + "\n# ---- v12 : indicateurs clés de 2020 au 31/08/2026")
rep("# ---- v11 : recommandations au format du rapport de diagnostic", bloc("bloc_matrices.py") + "\n# ---- v11 : recommandations au format du rapport de diagnostic")
rep('gab(107, {"titre": "Quatre taux à surveiller"', bloc("bloc_matrice_conformite.py") + '\ngab(107, {"titre": "Quatre taux à surveiller"')
rep('contenu({"titre": "Analyses en attente", "sous_titre": "Quatre analyses', bloc("bloc_matrice_donnees.py") + '\ncontenu({"titre": "Analyses en attente", "sous_titre": "Quatre analyses')
# les lignes suivantes d'une case gardent l'alignement de la première
rep('''            for li_, ligne_ in enumerate(str(v).split("\\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                for morceau, gras, surl in morceaux_balises(ligne_):
                    runs_texte(p, morceau, taille, gras_c or gras, coul, surl)''', '''            for li_, ligne_ in enumerate(str(v).split("\\n")):
                if li_:
                    p = c.text_frame.add_paragraph()
                    p.alignment = c.text_frame.paragraphs[0].alignment
                for morceau, gras, surl in morceaux_balises(ligne_):
                    runs_texte(p, morceau, taille, gras_c or gras, coul, surl)''')
# une case avec boule est plus haute (interligne de la police des émojis)
rep("""            n = sum(max(1, math.ceil(len(_brut(seg)) / cpl)) for seg in str(v).split("\\n"))
            h = max(h, n * taille * 1.2 * 0.03528 + 0.16)""", """            n = sum(max(1, math.ceil(len(_brut(seg)) / cpl)) for seg in str(v).split("\\n"))
            h = max(h, n * taille * 1.2 * 0.03528 + 0.16 + (0.22 if any(b_ in str(v) for b_ in BOULES) else 0))""")
