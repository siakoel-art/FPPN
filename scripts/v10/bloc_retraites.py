# ---- v10 : santé des policiers retraités, ce qui est mesurable (partie 8)
_CGv = R["cgrae"]
_CGs = [c for i, c in enumerate(_CGv) if i % 4 == 0 or i == len(_CGv) - 1]
contenu({"titre": "Santé des retraités : ce que l'on peut mesurer", "sous_titre": f"Les retraités précomptés passent de {N('SR début', _CGv[0]['retraites'])} en {_CGv[0]['libelle']} à {N('SR fin', _CGv[-1]['retraites'])} en {_CGv[-1]['libelle']}",
         "synthese": "Les cotisations des retraités sont mesurées ; leurs prestations, leurs âges et donc l'engagement à vie du FPPN ne le sont pas"},
        "Cotisation moyenne du mois = montant précompté par la CGRAE / nombre de retraités précomptés. L'engagement d'une garantie santé à vie se calcule avec l'âge de chaque retraité, son coût annuel de soins et une table de survie (CIMA F, partie 11).",
        tableau={"titre": tab("Retraités précomptés par la CGRAE"), "entetes": ["Mois", "Retraités précomptés", "Montant précompté", "Cotisation moyenne"], "largeurs": [5.0, 5.0, 6.0, 5.6], "taille": 9,
                 "lignes": [[c["libelle"], N(f"SR n {c['libelle']}", c["retraites"]), F(f"SR m {c['libelle']}", c["montant"]), F(f"SR moy {c['libelle']}", c["montant"] / c["retraites"])] for c in _CGs]},
        disposition="droite",
        analyse=["**Les retraités sont une population cotisante en forte hausse.** Leurs prestations ne sont pas isolées dans les comptes ni dans les fichiers reçus.",
                 f"**Les fichiers retour de la CGRAE de février à avril 2024 donnent la date de naissance de {N('SR dn', RDN['nb'])} retraités, âgés en moyenne de {N('SR âge', RDN['age_moyen'], 1)} ans au 30/09/2026.** Aucun fichier ne donne les dates des autres retraités ni les soins de chacun : l'engagement à vie ne peut pas être calculé (annexe « Méthodes non retenues et pourquoi »)."],
        implication="Tant que ces données manquent, le FPPN ne peut pas savoir si la cotisation des retraités couvre leurs soins jusqu'à la fin de leur vie.",
        valable=f"fichiers retour de la CGRAE reçus le 15/10/2026, {_CGv[0]['libelle']} à {_CGv[-1]['libelle']} (un mois sur quatre affiché, tous les mois dans le classeur) ; ne couvre pas les retraités non précomptés",
        source="Source : " + SRC_PREC[9:] + " ; CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; onglet 22 du classeur des calculs")
