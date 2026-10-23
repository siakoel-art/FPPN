# ---- v12 : boules de couleur (calcul) ; le dessin est dans le moteur des tableaux
BOULES = "🔴🟠🟡🟢⚪"
STATUT_BOULE = {"Conforme": "🟢", "Non conforme": "🔴", "Au-dessus": "🟠", "En attente": "🟡", "À apprécier": "⚪"}


def boule_statut(st):
    """Boule d'un statut de ratio (« Non conforme (indicatif) » compte comme « Non conforme »)."""
    return STATUT_BOULE.get(str(st).split(" (")[0].strip(), "")


def bv(texte, boule):
    """Case de matrice : boule devant la valeur (la valeur reste lisible)."""
    return f"{boule} {texte}" if boule else texte


def boule_fourchette(v, bas, haut):
    """Repère en fourchette (sinistralité par garantie) : 🟢 dans la fourchette, 🟠 en dessous, 🔴 au-dessus."""
    if v is None:
        return ""
    return "🟢" if bas <= v <= haut else ("🟠" if v < bas else "🔴")


def boule_seuil(v, seuil, sens="max"):
    """Seuil simple : sens « max » = la valeur doit rester au plus égale au seuil, « min » = au moins égale."""
    if v is None:
        return ""
    ok = v <= seuil if sens == "max" else v >= seuil
    return "🟢" if ok else "🔴"



# ---- v12 : définition et lecture de chaque ratio (formule et norme : Décision n° 05/2022 ; lecture : Adenka)
def val_ratio(x, a, retraite=False):
    v = (x.get("valeurs_retraitees") if retraite else x["valeurs"]).get(a)
    if v is None:
        return "n.d."
    if abs(v) < 5e-7:
        v = 0.0
    lab = f"{x['code']} {a}{' retraité' if retraite else ''}"
    if x["unite"] == "%":
        return pc(lab, v)
    if x["unite"] == "mois":
        return N(lab, v, 2) + " mois"
    if x["unite"] == "jours":
        return N(lab, v, 0) + " jours"
    if x["unite"] == "FCFA":
        return F(lab, v)
    return N(lab, v, 2)


DEFS = {  # code : (ce que le ratio mesure, comment le lire)
    "SIN": ("La part des cotisations de l'année consommée par les prestations, provisions comprises, toutes garanties confondues.",
            "Sous 70 %, les cotisations sont peu consommées. Au-dessus de 80 %, il ne reste pas assez pour payer les frais et constituer des réserves."),
    "MSE": ("La capacité des fonds propres libres à absorber une perte imprévue, rapportée au minimum exigé par le texte.",
            "Au moins 1 : la marge est suffisante. Sous 1 : les fonds propres libres sont trop faibles pour le volume d'activité."),
    "ROB": ("Le nombre de mois de prestations que la réserve obligatoire permettrait de payer.",
            "Moins de 6 mois : la réserve ne couvre pas un choc de prestations. Plus de 9 mois : réserve au-delà de la norme."),
    "FEP": ("La part des dettes, provisions techniques comprises, couverte par les fonds propres et assimilés.",
            "Sous 1 : les dettes dépassent les fonds propres et assimilés. Le FPPN dépend alors de ses créanciers et de l'épargne des adhérents."),
    "LIQ": ("La capacité à payer les dettes à court terme avec la trésorerie disponible.",
            "Sous 1 : la trésorerie ne suffit pas à payer les dettes exigibles à court terme."),
    "EQS": ("Le financement des emplois longs (immobilisations, prêts) par des ressources longues.",
            "Au-dessus de 1 : une partie des emplois longs est financée par des ressources de court terme, ce qui crée un risque de liquidité."),
    "AUT": ("La capacité des cotisations à payer l'ensemble des charges de l'année.",
            "Sous 1 : les cotisations ne suffisent pas à payer les charges de l'année."),
    "EXN": ("La part des cotisations conservée en excédent à la fin de l'exercice.",
            "Au-dessus de 5 % : excédent supérieur à la norme d'une mutuelle sociale. Négatif : déficit de l'exercice."),
    "TFP": ("La capacité de l'ensemble des produits, hors subventions, à payer l'ensemble des charges.",
            "Sous 100 % : les produits ne couvrent pas les charges et l'exercice est déficitaire."),
    "RL1": ("La part de l'épargne due (PEP et EFC) couverte par les actifs qui lui sont réservés : banques et dépôts à terme de l'épargne, FCP.",
            "Sous 1 : une partie de l'épargne des adhérents n'est garantie par aucun actif réservé."),
    "RL2": ("La capacité à payer les dettes à court terme avec les créances à court terme et la trésorerie.",
            "Sous 1 : les créances et la trésorerie ne suffisent pas à payer les dettes à court terme."),
    "RL3": ("La capacité à payer immédiatement les dettes à court terme avec la seule trésorerie.",
            "Sous 1 : un retrait massif ne pourrait pas être payé avec l'argent disponible."),
    "RL4": ("La réserve de trésorerie des risques longs, en trimestres de sorties d'épargne et de frais.",
            "0 : aucune réserve de trésorerie n'est constituée."),
    "RL5": ("La couverture par les fonds propres des sorties d'épargne des deux derniers exercices.",
            "Négatif : les fonds propres sont négatifs et ne couvrent aucune sortie."),
    "RL6": ("Le rendement des sommes réellement placées (dépôts à terme et FCP).",
            "À comparer au taux du marché et au taux servi aux épargnants : un rendement inférieur au taux servi creuse la dette d'épargne."),
    "RL7": ("L'argent laissé sur des comptes courants qui ne rapportent aucun intérêt.",
            "Plus le solde est élevé, plus le manque à gagner est important. La norme demande un solde proche de zéro."),
    "CHF": ("La part des cotisations absorbée par les charges de fonctionnement.",
            "Au-dessus de 20 % : les frais prennent une part des cotisations qui manque aux prestations et aux réserves."),
    "CHB": ("La part de l'ensemble des produits absorbée par les charges de fonctionnement brutes.",
            "Au-dessus de 15 % : les frais de gestion sont élevés au regard des ressources."),
    "CUP": ("Les frais de fonctionnement engagés pour 1 FCFA de prestations.",
            "Au-dessus de 15 % : la gestion coûte cher par rapport aux prestations servies."),
    "CHP": ("La part des charges de fonctionnement consacrée au personnel.",
            "Au-dessus de 85 % : les salaires laissent peu de place aux autres dépenses de fonctionnement."),
    "REC": ("La part des cotisations appelées effectivement encaissée.",
            "Sous 90 % : une partie des cotisations dues n'est pas encaissée."),
    "FID": ("La part des adhérents d'une année encore adhérents l'année suivante.",
            "Sous 80 % : les départs sont nombreux et la mutualisation s'affaiblit."),
    "DPP": ("Le délai moyen entre la prestation et le paiement du prestataire de soins.",
            "À comparer au délai contractuel des conventions (non transmis) et au délai de trois mois du règlement intérieur (art. 52)."),
    "FOR": ("La part des produits investie dans la formation du personnel.",
            "Au plus 3 % selon la norme."),
    "CRB": ("La croissance du nombre de bénéficiaires d'une année à l'autre.",
            "À comparer à l'objectif du plan de développement stratégique (non transmis)."),
}


def _defs_diapo(k, codes):
    """Diapositive de définitions avant chaque tableau de bord : formule, ce que le ratio mesure, lecture, norme."""
    noms = {1: "solvabilité et réserves", 2: "risques longs", 3: "gestion"}
    contenu({"titre": f"Définition et lecture des ratios ({k}/3) : {noms[k]}", "sous_titre": "Ce que chaque ratio mesure et comment lire sa valeur, avant le tableau de bord",
             "synthese": "Formule et norme de la Décision n° 05/2022 ; la lecture explique ce que la valeur dit de la situation du FPPN"},
            "Formules et normes de l'annexe de la Décision n° 05/2022/COM/UEMOA. La colonne « Comment le lire » est rédigée par Adenka pour un lecteur non spécialiste.",
            tableau={"titre": tab(f"Définition et lecture des ratios, {noms[k]}"), "entetes": ["Ratio", "Formule", "Ce que le ratio mesure", "Comment le lire", "Norme"],
                     "largeurs": [4.8, 7.0, 7.2, 8.8, 3.4], "taille": 8, "sans_code": True,
                     "lignes": [[RT[c]["nom"].replace(" (risques longs)", ""), RT[c]["formule"], DEFS[c][0], DEFS[c][1],
                                 RT[c]["norme"].replace(" (tableau 1) ; 9 mois au § 1.2.2.5.4", "").replace(" (marge disponible au moins égale à la marge exigée)", "")] for c in codes]},
            valable="formules et normes en vigueur depuis le 01/06/2022",
            source="Source : " + DEC + " ; lecture Adenka")


