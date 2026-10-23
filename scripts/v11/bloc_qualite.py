# ---- v11 : qualité des bases variable par variable (complétude et cohérence)
QB = json.load(open(os.path.join(os.path.dirname(V11J), "qualite.json"), encoding="utf-8"))
_SEL = ["Matricule", "Nom", "Montant", "Date de naissance"]
_bases = []
for x in QB["completude"]:
    if x["base"] not in _bases:
        _bases.append(x["base"])
_lq = []
for bse in _bases:
    lig = [x for x in QB["completude"] if x["base"] == bse]
    n_ = lig[0]["lignes"]
    if bse.startswith("Prêts"):
        vals = {x["variable"]: x["manquantes"] for x in lig}
        _lq.append([bse, N(f"QB n {bse}", n_), "Mécano : " + N(f"QB mec {bse}", vals["Mécano de l'emprunteur"]) + " manquants", "Dates, taux, durée : complets",
                    "Aucune incohérence de dates ni d'encours sur " + N(f"QB co {bse}", sum(1 for c in QB["coherence"] if c["base"] == bse)) + " contrôles", "Sans objet"])
        continue
    vals = {x["variable"]: x["manquantes"] for x in lig}
    dbl = next(c["anomalies"] for c in QB["coherence"] if c["base"] == bse and c["controle"].startswith("Même matricule"))
    neg = next(c["anomalies"] for c in QB["coherence"] if c["base"] == bse and c["controle"].startswith("Montant nul"))
    _lq.append([bse, N(f"QB n {bse}", n_), "Matricule : " + N(f"QB mat {bse}", vals["Matricule"]) + " manquants. Montant : " + N(f"QB mt {bse}", vals["Montant"]),
                "Date de naissance : " + pc(f"QB dn {bse}", vals["Date de naissance"] / n_) + " manquante",
                N(f"QB dbl {bse}", dbl) + " lignes d'un même matricule dans un mois. " + N(f"QB neg {bse}", neg) + " montants nuls ou négatifs",
                "Non" if vals["Date de naissance"] / n_ > 0.5 else "Oui"])
contenu({"titre": "Qualité des bases, variable par variable", "sous_titre": "Les montants et les identifiants sont complets, les dates de naissance absentes",
         "synthese": "Les bases permettent les calculs comptables et statistiques, pas les calculs par âge"},
        "Contrôles de complétude (valeur manquante) et de cohérence (doublons dans un mois, montants nuls ou négatifs, dates incohérentes) sur les lignes retenues de janvier 2020 à août 2026. Calculs par âge : possibles si la date de naissance est présente sur plus de la moitié des lignes.",
        tableau={"titre": tab("Complétude et cohérence des bases de données"), "entetes": ["Base", "Lignes", "Identifiants et montants", "Date de naissance", "Cohérence", "Utilisable pour les calculs par âge ?"],
                 "largeurs": [5.6, 2.8, 6.6, 4.8, 8.4, 3.0], "taille": 8.5, "lignes": _lq},
        disposition="dessous",
        analyse=["**Les lignes multiples d'un même matricule dans un mois ne sont pas des erreurs pour le PEP et le Grenier** : un épargnant peut avoir plusieurs plans et un emprunteur plusieurs prêts. Pour la maladie et le décès, les effectifs comptent les matricules distincts : ces lignes ne gonflent pas les effectifs.",
                 "**Seuls les fichiers CGRAE de février à avril 2024 portent la date de naissance** (105 retraités). Les grades et salaires ne figurent que dans les fichiers détaillés de la maladie d'avril à octobre 2024 et de janvier 2025 (209 670 lignes, complètes)."],
        implication="La qualité des montants autorise les analyses comptables du rapport. L'absence de date de naissance empêche les seuls calculs par âge (mortalité, engagements viagers, projection).",
        valable="lignes retenues des fichiers de précomptes de janvier 2020 à août 2026 ; balance âgée du Grenier au 31/08/2026",
        source=SRC_PREC + " ; " + SRC_GRENIER[9:] + " ; script qualite_bases_fppn.py")
