# ---- v12 : ratios par garantie, avant les ratios globaux : matrice de 2020 au 31/08/2026, puis détail des montants
_RGL = V11["ratios_garantie"]
_RGg = {(x["exercice"], x["garantie"]): x for x in _RGL}
_MAT_G = []
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : sinistralité", "Repère 70 % à 80 %"] + [_gar(_g, a, "sinistralite") for a in _AN])
_MAT_G.append(["Toutes garanties : taux de sinistralité réglementaire", "Norme 70 % à 80 %"] + [_rt("SIN", a) for a in _AN])
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : ratio combiné (prestations et frais)", "Au plus 100 %"] + [_gar(_g, a, "ratio_combine") for a in _AN])
for _g in ("Maladie", "Décès"):
    _MAT_G.append([f"{_g} : résultat technique", "Au moins 0 FCFA"] + [_gar(_g, a, "resultat_technique") for a in _AN])
contenu({"titre": "Ratios par garantie : matrice de 2020 au 31/08/2026", "sous_titre": "La maladie est déficitaire depuis 2025, le décès reste très excédentaire",
         "synthese": "La Décision n° 05/2022 ne fixe qu'un taux de sinistralité global ; les ratios par garantie montrent quelle garantie le fait monter"},
        "La Décision n° 05/2022 (§ 1.2.1.2.5) définit un seul taux de sinistralité : toutes les prestations de l'exercice avec la variation des provisions (comptes 60, 6901 et 7901), rapportées à toutes les cotisations acquises (comptes 70, 7906 et 6906). La norme de 70 % à 80 % s'applique à ce taux global. Les ratios par garantie sont un outil d'analyse ; la fourchette y sert de repère.",
        tableau={"titre": tab("Ratios par garantie et taux réglementaire, de 2020 au 31/08/2026"), "entetes": ["Ratio", "Norme ou repère"] + _ENT_AN,
                 "largeurs": [7.2, 3.2] + [2.97] * 7, "taille": 8, "sans_code": True, "legende": LEG_MAT[:3], "lignes": _MAT_G},
        disposition="dessous",
        encadre=["**Sinistralité d'une garantie** = prestations de la garantie / cotisations de la garantie. Elle montre la part des cotisations consommée par les soins (maladie) ou les allocations (décès).",
                 "**Ratio combiné** = (prestations + charges de fonctionnement affectées) / cotisations. Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte, frais compris. Frais répartis au prorata des cotisations, faute de clé fournie par le FPPN.",
                 "**Résultat technique** = cotisations - prestations - charges de fonctionnement affectées, hors produits financiers et éléments exceptionnels. Négatif : la garantie est déficitaire par elle-même."],
        analyse=["**Le taux réglementaire global réunit les deux garanties : il masque une maladie déficitaire et un décès très excédentaire.** "
                 f"En 2025, la maladie consomme {pc('RGM sin m25', _RGg[(2025, 'Maladie')]['sinistralite'])} de ses cotisations, le décès {pc('RGM sin d25', _RGg[(2025, 'Décès')]['sinistralite'])}, pour un taux global de {pc('RGM sin 25', RT['SIN']['valeurs'][2025])}."],
        implication="Le respect de la norme globale ne suffit pas : la cotisation maladie doit être fixée sur le coût de la garantie maladie, sans compter sur l'excédent du décès.",
        valable="comptes de 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire) ; en 2020, le décès comprend la prime d'assurance décès (compte 65180000), hors du taux réglementaire",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.2.1.2.5 ; statuts du FPPN édités le 15/06/2023, art. 17")
contenu({"titre": "Ratios par garantie : montants", "sous_titre": "Cotisations, prestations, frais affectés et résultat technique de chaque garantie",
         "synthese": "Résultat technique = cotisations acquises - prestations (variation des provisions comprise) - charges de fonctionnement affectées, hors produits financiers et éléments exceptionnels"},
        "Charges de fonctionnement (comptes 612, 613, 618 et 62 à 68) réparties au prorata des cotisations, faute de clé fournie par le FPPN. Décès : comptes 6021 et 6024, prime d'assurance décès comprise en 2020, cotisations du compte 702 (décès et MEDOC).",
        tableau={"titre": tab("Sinistralité, frais, ratio combiné et résultat technique par garantie"), "entetes": ["Exercice", "Garantie", "Cotisations", "Prestations", "Sinistralité", "Frais affectés", "Ratio combiné", "Résultat technique"],
                 "largeurs": [2.6, 2.8, 4.6, 4.6, 3.0, 4.6, 3.0, 5.0], "taille": 8, "sans_code": True, "legende": LEG_MAT[:3],
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "2026 (8 mois)", x["garantie"], F(f"RG cot {x['exercice']} {x['garantie']}", x["cotisations"]), F(f"RG pre {x['exercice']} {x['garantie']}", x["prestations"]),
                             _gar(x["garantie"], x["exercice"], "sinistralite"), F(f"RG fr {x['exercice']} {x['garantie']}", x["frais"]), _gar(x["garantie"], x["exercice"], "ratio_combine"),
                             _gar(x["garantie"], x["exercice"], "resultat_technique")] for x in _RGL if x["exercice"] >= 2022]},
        disposition="dessous",
        analyse=[f"**Le ratio combiné de la maladie dépasse 100 % en 2025 ({pc('RG rc m25 (a)', _RGg[(2025, 'Maladie')]['ratio_combine'])}) et sur les huit mois de 2026.** Celui du décès reste sous 70 % chaque année depuis 2022."],
        implication="La cotisation maladie de 3,5 % du salaire de base (statuts, art. 17) ne couvre plus le coût complet de la garantie maladie depuis 2025. Sa révision relève du conseil d'administration et de l'Assemblée générale.",
        valable="comptes 2022 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; frais répartis au prorata des cotisations",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; statuts du FPPN édités le 15/06/2023, art. 17")
