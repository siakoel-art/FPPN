# ---- v10 : annexe « Méthodes non retenues et pourquoi » (conditions d'emploi citées, démonstration chiffrée, conséquence)
_MB = [x for x in V9["mortalite_brute"] if 2021 <= x["annee"] <= 2025]
_DMB = sum(x["deces_adherents"] for x in _MB)
_LAT = {c: TARDM[c]["references_annee_precedente"] / TARDM[c]["factures_sinistres"] for c in TARDM}
_pm25a = next(x["placements_moyens"] for x in R["rendement_placements"] if x["exercice"] == 2025)
SRC_CIRC = "circulaire n° 00230/CIMA/CRCA/PDT/2005"


def methode_non_retenue(titre, sous_titre, synthese, principe, lignes, consequence, a_la_place, valable_, source_):
    contenu({"titre": titre, "sous_titre": sous_titre, "synthese": synthese}, principe,
            tableau={"titre": tab("Conditions d'emploi de la méthode et données du FPPN"), "entetes": ["Condition d'emploi", "Source de la condition", "Ce que montrent les données du FPPN", "Remplie ?"],
                     "largeurs": [6.0, 7.4, 14.2, 3.6], "taille": 8, "lignes": lignes},
            disposition="dessous", analyse=[consequence], implication=a_la_place, valable=valable_, source=source_)


contenu({"titre": "Méthodes non retenues : vue d'ensemble", "sous_titre": "Sept méthodes courantes ne sont pas appliquées : leurs conditions d'emploi ne sont pas remplies",
         "synthese": "Appliquer une méthode hors de ses conditions d'emploi produirait des chiffres précis en apparence et faux en réalité"},
        "Pour chaque méthode, les diapositives suivantes citent les conditions d'emploi dans un texte ou un ouvrage de référence, montrent avec les chiffres du FPPN pourquoi elles ne sont pas remplies, et indiquent la méthode appliquée à la place.",
        tableau={"titre": tab("Méthodes non retenues et méthodes appliquées à la place"), "entetes": ["Méthode non retenue", "Condition principale non remplie", "Méthode appliquée à la place", "Détail"], "largeurs": [6.4, 11.0, 9.0, 4.8], "taille": 8.5,
                 "lignes": [["Chain Ladder et méthode CIMA des tardifs", "Des paiements étalés sur plusieurs années de développement", "Méthode courte sur les factures en retard", ref("Méthode non retenue : Chain Ladder et méthode CIMA")],
                            ["Mortalité observée par âge (A/E)", "La date de naissance de chaque mutualiste", "Taux brut de mortalité", ref("Méthode non retenue : mortalité observée par âge")],
                            ["Provision mathématique santé des retraités", "L'âge et le coût des soins de chaque retraité", "Description des cotisations", ref("Méthode non retenue : provision santé des retraités")],
                            ["Duration et immunisation", "Des flux de passif connus et des actifs à revenu fixe", "Écarts de liquidité", ref("Méthode non retenue : duration et immunisation")],
                            ["Modèle actif-passif stochastique", "Des contrats décrits par âge et des scénarios économiques", "Tests de résistance calibrés sur le FPPN", ref("Méthode non retenue : modèle stochastique")],
                            ["Valeur en risque (VaR)", "Des actifs évalués au prix du marché", "Classes de liquidité", ref("Méthode non retenue : valeur en risque")],
                            ["Capital de solvabilité de type Solvabilité II", "Un cadre réglementaire qui l'impose et des engagements en meilleure estimation", "Marge de solvabilité de la Décision n° 05/2022", ref("Méthode non retenue : capital de solvabilité")]]},
        disposition="dessous",
        analyse=["**Une même cause revient : l'absence de données individuelles (âges, dates de soins, comptes d'épargne).** Les méthodes retenues exploitent au mieux les données agrégées reçues."],
        implication="Chaque méthode non retenue pourra être appliquée dès réception des données citées dans sa diapositive.",
        valable="données reçues jusqu'au 22/10/2026", source="Source : Adenka, d'après les références citées dans chaque diapositive")
methode_non_retenue("Méthode non retenue : Chain Ladder et méthode CIMA", "Les factures se règlent en quelques mois : un triangle annuel n'aurait qu'une seule colonne utile",
                    f"Moins de {pc('AN lat max', max(_LAT.values()))} des factures d'une année sont enregistrées l'année suivante",
                    "Chain Ladder et la méthode CIMA des tardifs projettent le rythme passé de règlement, année de survenance par année de développement. Elles ne sont utiles que si une part notable des sinistres se règle après la fin de l'année de survenance.",
                    [["Paiements étalés sur plusieurs années de développement", SRC_CIRC + " (cadences sur quatre années de survenance)", f"Part des factures enregistrées l'année suivante : {', '.join(pc(f'AN lat {c}', _LAT[c]) for c in sorted(_LAT))} (clôtures {', '.join(str(c) for c in sorted(_LAT))})", "Non"],
                     ["Rythme de règlement stable d'une année à l'autre", "Hypothèse de base de Chain Ladder", f"Délai moyen de paiement des cliniques : {N('AN d20', DELAI[2020], 0)} mois en 2020, {N('AN d21', DELAI[2021], 0)} en 2021, {N('AN d25', DELAI[2025], 2)} en 2025 : le rythme a changé de nature", "Non"],
                     ["Base des prestations datée (soins, dépôt, paiement)", REX + ", art. 10", "Non reçue (PBC, item 71) ; seule l'année de la facture est lisible dans son numéro", "Non"],
                     ["Délai de règlement borné par les textes", SRC_RI + ", art. 52, 57 et 60", "Paiement à trois mois du dépôt, demande de remboursement limitée à trois mois", "Oui"]],
                    "**Appliquée aux données du FPPN, la méthode projetterait le rattrapage des arriérés de 2020 à 2022 comme un rythme normal et surestimerait fortement la provision.**",
                    "Méthode appliquée : soins de l'année facturés l'année suivante, mesurés facture par facture (diapositive « Provision maladie : factures arrivées en retard »).",
                    "factures de cliniques du journal SIN 2023 à 2025 ; délais calculés sur le compte 40110000 de 2020 au 31/08/2026",
                    "Source : " + SRC_CIRC + " ; " + REX + ", art. 10 ; " + SRC_RI + " ; " + SRC_GLT)
methode_non_retenue("Méthode non retenue : mortalité observée par âge", "Sans date de naissance, les décès attendus par la table ne peuvent pas être calculés",
                    f"{N('AN déc', _DMB)} décès de membres de 2021 à 2025 : même avec les âges, l'écart à la table ne serait mesuré qu'à ±{pc('AN ic', 1.96 / _DMB ** 0.5)} près",
                    "Le rapport A/E divise les décès observés par les décès attendus : somme, âge par âge, de l'exposition au risque multipliée par la probabilité de décès de la table CIMA H.",
                    [["Exposition au risque par âge", "Définition de la table de mortalité, " + REX + ", art. 1", "Aucun fichier reçu ne porte la date de naissance des mutualistes", "Non"],
                     ["Décès datés et rattachés à un âge", "Même source", f"Décès de membres comptés par année ({N('AN déc (t)', _DMB)} de 2021 à 2025), sans âge", "Non"],
                     ["Nombre de décès suffisant", "Loi de Poisson : intervalle à 95 % d'environ ±1,96 / racine du nombre de décès", f"Intervalle de ±{pc('AN ic (t)', 1.96 / _DMB ** 0.5)} sur cinq ans", "Partiellement"]],
                    "**La mortalité des policiers ne peut pas être comparée à la table CIMA H.** Seul le taux brut, toutes tranches d'âge confondues, est calculable.",
                    "Méthode appliquée : taux brut de mortalité des membres participants (diapositive « Mortalité brute des membres participants »).",
                    "décès de membres recensés par le FPPN de 2021 à 2025", SRC_RATIO + " ; " + SRC_CIMA[9:])
methode_non_retenue("Méthode non retenue : provision santé des retraités", "L'engagement à vie ne se calcule qu'avec l'âge et le coût des soins de chaque retraité",
                    f"{N('AN ret', CG[-1]['retraites'])} retraités précomptés en {CG[-1]['libelle']}, dont {N('AN ret dn', RDN['nb'])} avec une date de naissance, sans prestations isolées",
                    "Provision mathématique = valeur actuelle des soins futurs - valeur actuelle des cotisations futures, retraité par retraité, avec une table de survie (CIMA F) et un taux d'actualisation.",
                    [["Âge de chaque retraité", "Code CIMA, art. 334-8 (référence technique)", f"Date de naissance pour {N('AN ret dn (t)', RDN['nb'])} retraités seulement (fichiers retour de la CGRAE de février à avril 2024)", "Partiellement"],
                     ["Coût annuel des soins par âge", "Même source", "Prestations des retraités non isolées dans les comptes ni dans les fichiers", "Non"],
                     ["Règle de la garantie (durée, cotisation)", "Bulletin de souscription, annexe 2 de la convention", "Non transmis", "Non"]],
                    "**Calculer une provision avec un âge moyen supposé donnerait un chiffre sans fondement** : la valeur d'une rente à 60 ans varie de plus de deux années de paiement selon la seule table retenue (partie 11).",
                    "Méthode appliquée : description des cotisations des retraités (diapositive « Santé des retraités : ce que l'on peut mesurer »).",
                    f"retraités précomptés par la CGRAE de {CG[0]['libelle']} à {CG[-1]['libelle']} ; dates de naissance des fichiers retour de février à avril 2024", "Source : CONVENTION DE PARTENARIAT ENTRE IPS-CGRAE ET LE FPPN.pdf ; Code CIMA, art. 334-8 ; " + SRC_PREC[9:])
methode_non_retenue("Méthode non retenue : duration et immunisation", "L'immunisation suppose des flux de passif connus et des actifs à revenu fixe",
                    f"Placements financiers moyens 2025 : {pc('AN pm', _pm25a / EP_DUE25)} de l'épargne due ; taux servi non arrêté",
                    "L'immunisation rééquilibre les actifs pour que leur valeur varie comme celle des passifs quand les taux changent. Elle compare les durations (durée moyenne pondérée des flux) de l'actif et du passif.",
                    [["Flux de passif prévisibles", SRC_ALM["IAA"] + ", p. 3 et 23", "Épargne sans échéancier ; taux servi de 3,5 %, 5,5 % ou 14 % selon le document", "Non"],
                     ["Actifs à revenu fixe en volume suffisant", "Même source", f"DAT soldé chaque année et FCP ; placements moyens 2025 de {F('AN pm (t)', _pm25a)}", "Non"],
                     ["Risque principal : le taux d'intérêt", SRC_ALM["ONT"] + ", p. 8", "Risque principal : la liquidité (couverture de l'année inférieure à 40 %)", "Non"]],
                    "**Le risque du FPPN n'est pas une variation de taux, mais un manque de liquidité** : mesurer une duration ne dirait rien de la capacité à payer les sorties de l'année.",
                    "Méthode appliquée : écarts de liquidité par échéance (diapositive « Adossement : écarts de liquidité année par année »).",
                    "bilans au 31/12/2025 et au 31/08/2026", "Source : " + SRC_ALM["IAA"] + " ; " + SRC_ALM["ONT"] + " ; " + SRC_COMPTA[9:])
methode_non_retenue("Méthode non retenue : modèle stochastique", "Un modèle stochastique demande des contrats décrits par âge et des scénarios économiques",
                    "Aucune des deux entrées du modèle n'existe pour le FPPN",
                    "Un modèle actif-passif stochastique projette des milliers de scénarios économiques (méthode de Monte-Carlo) sur des groupes de contrats décrits par âge, sexe, ancienneté et taux garanti.",
                    [["Groupes de contrats par âge, sexe, ancienneté et taux garanti", SRC_ALM["TIC"] + ", p. 54", "Âges non reçus ; taux garanti non arrêté", "Non"],
                     ["Actifs en valeur de marché", "Même source", "Terrains, clinique et créances en valeur comptable non réévaluée", "Non"],
                     ["Générateur de scénarios économiques", "Même source, section 3.2", "Aucune série de prix des actifs du FPPN", "Non"]],
                    "**Le modèle produirait une distribution de résultats à partir d'entrées inventées.**",
                    "Méthode appliquée : trois scénarios déterministes calibrés sur les extrêmes observés au FPPN (diapositive « Adossement : tests de résistance »).",
                    "données reçues jusqu'au 22/10/2026", "Source : " + SRC_ALM["TIC"])
methode_non_retenue("Méthode non retenue : valeur en risque", "La VaR mesure une perte de valeur de marché, alors que les actifs du FPPN ne sont pas évalués au marché",
                    "Les actifs du FPPN sont tenus en valeur comptable ; aucun n'est coté à l'exception du FCP",
                    "La valeur en risque (VaR) est la perte de valeur de marché qui n'est dépassée qu'avec une faible probabilité sur une période donnée.",
                    [["Indicateur adapté à la gestion actif-passif", SRC_ALM["BAN"] + ", p. 9", "La source indique que la VaR n'est pas adaptée à l'ALM, « monde des intérêts courus »", "Non"],
                     ["Séries de prix des actifs", "Définition de la VaR", "Aucune série de prix pour les terrains, la clinique, les prêts et les créances", "Non"]],
                    "**Une VaR ne mesurerait que le FCP, une petite part de l'actif.**",
                    "Méthode appliquée : classes de liquidité des actifs (diapositive « Choix des méthodes d'adossement »).",
                    "bilan au 31/08/2026", "Source : " + SRC_ALM["BAN"])
methode_non_retenue("Méthode non retenue : capital de solvabilité", "Le capital de type Solvabilité II n'est pas exigé du FPPN et demande des engagements en meilleure estimation",
                    "Le cadre applicable est la marge de solvabilité de la Décision n° 05/2022, calculée en partie 10",
                    "Le capital de solvabilité requis (SCR) de la formule standard européenne applique des chocs (par exemple 40 % de rachats massifs) à la meilleure estimation des engagements.",
                    [["Cadre réglementaire qui l'impose", "Décision n° 05/2022 ; " + REX, "Le texte applicable fixe une marge de solvabilité, non un SCR", "Non"],
                     ["Engagements en meilleure estimation", SRC_ALM["TIC"] + ", p. 110", "Épargne sans échéancier ni taux arrêté ; provisions non calculées", "Non"],
                     ["Chocs calibrés sur le portefeuille", "Même source", f"Choc de 40 % calibré sur des contrats européens ; sortie maximale observée au FPPN : {pc('AN tmax', T_MAX)}", "Non"]],
                    "**Le SCR mélangerait un cadre étranger et des engagements non estimés.**",
                    "Méthode appliquée : marge de solvabilité de la Décision n° 05/2022 et tests de résistance calibrés sur le FPPN (partie 10 et partie 9).",
                    "ratios de 2020 au 31/08/2026", "Source : " + SRC_ALM["TIC"] + " ; Décision n° 05/2022 ; " + REX)
