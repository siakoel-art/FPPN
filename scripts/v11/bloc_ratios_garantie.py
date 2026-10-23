# ---- v11 : ratios par garantie et définition du résultat technique (après l'excédent par risque)
_RGL = V11["ratios_garantie"]
contenu({"titre": "Ratios par garantie", "sous_titre": "La maladie consomme plus que ses cotisations depuis 2025, le décès en laisse plus de la moitié",
         "synthese": "Résultat technique = cotisations acquises - prestations (variation des provisions comprise) - charges de fonctionnement affectées, hors produits financiers et éléments exceptionnels"},
        "Ratio combiné = (prestations + charges de fonctionnement affectées) / cotisations. Au-dessus de 100 %, la garantie coûte plus qu'elle ne rapporte. Charges de fonctionnement (comptes 612, 613, 618 et 62 à 68) réparties au prorata des cotisations, faute de clé fournie par le FPPN. Décès : comptes 6021 et 6024, prime d'assurance décès comprise en 2020, cotisations du compte 702 (décès et MEDOC).",
        tableau={"titre": tab("Sinistralité, frais, ratio combiné et résultat technique par garantie"), "entetes": ["Exercice", "Garantie", "Cotisations", "Prestations", "Sinistralité", "Frais affectés", "Ratio combiné", "Résultat technique"],
                 "largeurs": [2.6, 2.8, 4.6, 4.6, 3.0, 4.6, 3.0, 5.0], "taille": 8,
                 "lignes": [[str(x["exercice"]) if x["exercice"] < 2026 else "2026 (8 mois)", x["garantie"], F(f"RG cot {x['exercice']} {x['garantie']}", x["cotisations"]), F(f"RG pre {x['exercice']} {x['garantie']}", x["prestations"]),
                             pc(f"RG sin {x['exercice']} {x['garantie']}", x["sinistralite"]), F(f"RG fr {x['exercice']} {x['garantie']}", x["frais"]), pc(f"RG rc {x['exercice']} {x['garantie']}", x["ratio_combine"]),
                             F(f"RG rt {x['exercice']} {x['garantie']}", x["resultat_technique"])] for x in _RGL if x["exercice"] >= 2022]},
        disposition="dessous",
        analyse=[f"**Le ratio combiné de la maladie dépasse 100 % en 2025 ({pc('RG rc m25 (a)', next(x['ratio_combine'] for x in _RGL if x['exercice'] == 2025 and x['garantie'] == 'Maladie'))}) et sur les huit mois de 2026.** Celui du décès reste sous 70 % chaque année depuis 2022.",
                 "**Le résultat technique isole ce que chaque garantie dégage par elle-même.** Il ne compte ni les produits financiers, ni les produits annexes, ni les éléments exceptionnels, qui expliquent l'écart avec le résultat net."],
        implication="La cotisation maladie de 3,5 % du salaire de base (statuts, art. 17) ne couvre plus le coût complet de la garantie maladie depuis 2025. Sa révision relève du conseil d'administration et de l'Assemblée générale.",
        valable="comptes 2022 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; frais répartis au prorata des cotisations",
        source=SRC_COMPTA + " ; " + DEC + ", § 1.1.1 ; statuts du FPPN édités le 15/06/2023, art. 17")
