# ---------------------------------------------------------------- PARTIE 12 : PROJECTION DE L'ÉQUILIBRE TECHNIQUE
separateur(12, "Projection de l'équilibre technique", "Avertissement, point de blocage, deux scénarios de 2026 à 2030")
_PA, _PB = V10["projection"]["A"], V10["projection"]["B"]
_TC = V10["tcam"]
_CRP = {x["exercice"]: x["prest_maladie"] for x in V10["croissance_prestations"]}
contenu({"titre": "Projection : avertissement", "sous_titre": "Sans les âges des mutualistes, la projection ne peut être qu'agrégée",
         "synthese": "La projection actuarielle bloque dès sa première étape : aucun fichier reçu ne donne la date de naissance des mutualistes en activité"},
        "Une projection actuarielle de l'équilibre technique suit six étapes. Le tableau indique, pour chacune, la donnée nécessaire, si elle a été reçue, et ce qu'il faut faire à la place.",
        tableau={"titre": tab("Étapes d'une projection actuarielle et point de blocage"), "entetes": ["Étape", "Donnée nécessaire", "Reçue ?", "Conséquence"], "largeurs": [5.0, 9.2, 3.2, 13.8], "taille": 8.5,
                 "lignes": [["1. Population par âge", "Date de naissance et sexe de chaque mutualiste", "Non", f"Blocage : les fichiers de la Solde et le fichier des effectifs ne portent que la date de recrutement ; la CGRAE ne donne la date de naissance que de {N('PJ ret dn', RDN['nb'])} retraités"],
                            ["2. Vieillissement et décès", "Âge de chaque mutualiste et table CIMA H", "Table seulement", "Impossible sans l'étape 1 : la table ne s'applique qu'à un âge"],
                            ["3. Départs à la retraite", "Âge et grade (âge légal de départ)", "Grade seulement", "Impossible : on ne peut dater ni les départs, ni les versements de l'EFC qu'ils déclenchent"],
                            ["4. Entrées", "Promotions de recrues prévues", "Passées seulement", "Possible en tendance (partie 4)"],
                            ["5. Cotisations et prestations par âge", "Salaires et coût des soins par âge", "Non", "Impossible : seuls les totaux annuels sont connus"],
                            ["6. Équilibre de chaque année", "Résultats des étapes 1 à 5", "Partiel", "Remplacé par le prolongement des tendances agrégées de 2023 à 2025"]]},
        disposition="dessous",
        analyse=["**La projection présentée ensuite prolonge les tendances passées, sans les expliquer.** Elle ne dit pas pourquoi les prestations augmentent (vieillissement, prix des soins, nombre de bénéficiaires) et ne peut donc pas tester l'effet d'une mesure ciblée."],
        implication="Le fichier des mutualistes avec date de naissance et sexe (PBC, item 9) est la donnée qui débloque à la fois la projection, la mortalité observée et l'engagement santé des retraités.",
        valable="fichiers reçus jusqu'au 22/10/2026 : précomptes de la Solde (postes 936, 8Z6, 998, 870), fichier des effectifs du FPPN, fichiers retour de la CGRAE (dates de naissance de février à avril 2024 seulement)",
        source="Source : liste des documents demandés (PBC) pointée au 22/10/2026 ; " + SRC_PREC[9:])
contenu({"titre": "Projection agrégée : deux scénarios", "sous_titre": f"Si les tendances de 2023 à 2025 se prolongent, le résultat technique devient négatif dès 2026 ({F('PJ A 2026', _PA[0]['resultat'])})",
         "synthese": f"Scénario A (tendance) : {F('PJ A 2030', _PA[-1]['resultat'])} en 2030 ; scénario B (prestations au rythme des cotisations) : {F('PJ B 2030', _PB[-1]['resultat'])}"},
        f"Chaque poste de 2025 croît chaque année à son taux annuel moyen. Cotisations : {pc('PJ tc', _TC['cotisations'])} ; frais : {pc('PJ tf', _TC['frais'])} (2023 à 2025). Prestations : {pc('PJ tp', _TC['prestations'])} dans le scénario A, taux des cotisations dans le scénario B. Résultat technique = cotisations - prestations - frais. Les années 2021 et 2022 sont écartées : leurs prestations maladie sont anormalement basses.",
        tableau={"titre": tab("Projection agrégée du résultat technique, maladie et décès"), "entetes": ["Exercice", "Cotisations", "Prestations A", "Résultat A", "Prestations B", "Résultat B"], "largeurs": [3.0, 5.4, 5.4, 5.4, 5.4, 5.4], "taille": 9,
                 "lignes": [[str(a["exercice"]), F(f"PJ cot {a['exercice']}", a["cotisations"]), F(f"PJ pA {a['exercice']}", a["prestations"]), F(f"PJ rA {a['exercice']}", a["resultat"]),
                             F(f"PJ pB {a['exercice']}", bb["prestations"]), F(f"PJ rB {a['exercice']}", bb["resultat"])] for a, bb in zip(_PA, _PB)]},
        disposition="dessous",
        analyse=[f"**Au rythme de 2023 à 2025, les prestations croissent de {pc('PJ tp (a)', _TC['prestations'])} par an et les cotisations de {pc('PJ tc (a)', _TC['cotisations'])} : l'équilibre technique ne tient pas.**",
                 f"**La hausse des prestations maladie ralentit déjà : {pc('PJ cr 2024', _CRP[2024])} en 2024, {pc('PJ cr 2025', _CRP[2025])} en 2025.** Prolonger le rythme moyen est donc une hypothèse pessimiste.",
                 "**Le scénario B montre l'équilibre atteint si les prestations suivent les cotisations.** La réalité se situe entre les deux scénarios ; seule une projection par âge permettrait de la situer."],
        implication="Le FPPN doit suivre chaque trimestre le rapport prestations / cotisations de la maladie : c'est l'indicateur avancé du déséquilibre que montre le scénario A.",
        valable="cotisations, prestations et frais de gestion des comptes 2023 à 2025, garanties maladie et décès prises ensemble ; hors épargne, Grenier et produits financiers",
        source=SRC_COMPTA + " ; onglet 25 du classeur des calculs")
contenu({"titre": "Projection agrégée : évolution du résultat", "sous_titre": "L'écart entre les deux scénarios mesure l'incertitude laissée par l'absence des âges",
         "synthese": "Le scénario A sert d'alerte, le scénario B de cible : le pilotage des prestations maladie fait passer de l'un à l'autre"},
        "Résultat technique projeté de 2026 à 2030 dans les deux scénarios, en millions de FCFA.",
        graphique={"titre": fig("Résultat technique projeté selon deux scénarios"), "type": "barres", "categories": [str(a["exercice"]) for a in _PA],
                   "series": [{"nom": "Scénario A : tendance 2023 à 2025", "valeurs": [round(a["resultat"]) for a in _PA], "couleur": ORANGE},
                              {"nom": "Scénario B : prestations au rythme des cotisations", "valeurs": [round(a["resultat"]) for a in _PB], "couleur": BLEU}],
                   "format": FMT_M, "unite": "En millions de FCFA"},
        analyse=[f"**En 2030, l'écart entre les deux scénarios atteint {F('PJ écart', _PB[-1]['resultat'] - _PA[-1]['resultat'])}.**",
                 "**Aucun des deux scénarios n'est une prévision** : ce sont deux bornes construites sur les seules données agrégées disponibles."],
        implication="La projection par âge doit être refaite dès réception du fichier des mutualistes ; elle remplacera ces deux bornes par un scénario central expliqué.",
        valable="mêmes données et hypothèses que la diapositive précédente",
        source=SRC_COMPTA + " ; onglet 25 du classeur des calculs")

