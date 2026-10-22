# -*- coding: utf-8 -*-
"""Construit scripts/rapport_v9_fppn.py à partir de rapport_v8_fppn.py (remplacements contrôlés et insertions)."""
import os

REPO = "/home/user/FPPN/scripts"
ICI = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(REPO, "rapport_v8_fppn.py"), encoding="utf-8").read()


def rep(a, b, n=1):
    global s
    k = s.count(a)
    assert k == n, (k, a[:100])
    s = s.replace(a, b)


def bloc(nom):
    return open(os.path.join(ICI, nom), encoding="utf-8").read()


# ------------------------------------------------------------------ en-tête, arguments
rep("FPPN - Rapport actuariel et statistique, version 8 (PowerPoint, charte Adenka, Timesaver)",
    "FPPN - Rapport actuariel et statistique, version 9 (PowerPoint, charte Adenka, Timesaver)")
rep("Nouveautés de la version 8 (analyses de sensibilité, statuts et règlement intérieur du 15/06/2023) :",
    """Nouveautés de la version 9 (documents reçus le 22/10/2026, périmètre et validité des analyses) :
- périmètre, domaine et période de validité à trois niveaux : diapositive de synthèse (partie 2), ligne « Valable pour »
  sous chaque tableau et graphique, tableau détaillé en annexe ; colonne « Valable pour » dans le classeur ;
- contrôles croisés (partie 3) : états financiers contre balances, effectifs du FPPN contre fichiers 8Z6, rapport
  8Z6 / 936 par mois, rapprochements bancaires ;
- méthode de sinistralité du FPPN (communiquée le 22/10/2026) contre méthode réglementaire, sinistralité décès des deux
  sources, mortalité brute des membres (partie 5) ;
- taux servis sur l'épargne selon les documents, absence d'intérêts comptabilisés, sensibilité de l'EFC au taux,
  épargne présentée dans le rapport de gestion 2025 (partie 6) ; convention IPS-CGRAE (partie 8) ;
- provision du personnel recalculée agent par agent à partir des DISA, séparée entre Direction Générale, CMP et Grenier,
  avec sensibilités (partie 9) ; corrections (mesure du taux de 1,23 %, part de financement du Grenier, taux servi).

Nouveautés de la version 8 (analyses de sensibilité, statuts et règlement intérieur du 15/06/2023) :""")
rep("Utilisation : python rapport_v8_fppn.py <resultats.json> <complements.json> <ratios.json> <dossier_images> <skill> <sortie.pptx> <dossier_base>",
    "Utilisation : python rapport_v9_fppn.py <resultats.json> <complements.json> <ratios.json> <dossier_images> <skill> <sortie.pptx> <dossier_base>")
rep("              <classeur_sensibilites_recalcule.xlsx> <carte_sensibilites.json>\n",
    "              <classeur_sensibilites_recalcule.xlsx> <carte_sensibilites.json> <donnees_v9.json>\n")
rep("RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8 = sys.argv[1:10]", "RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J = sys.argv[1:11]")
rep('PV = json.load(open(os.path.join(os.path.dirname(RES), "provisions.json"), encoding="utf-8"))\n',
    'PV = json.load(open(os.path.join(os.path.dirname(RES), "provisions.json"), encoding="utf-8"))\nV9 = json.load(open(V9J, encoding="utf-8"))\n')

# ------------------------------------------------------------------ validité : partie courante et ligne « Valable pour »
rep('''def contenu(entete, notes, **k):
    SL.append({"type": "contenu", "entete": entete, "notes": notes, **k})


def separateur(num, titre, sous_titre, notes=""):
    SL.append({"type": "separateur", "num": num, "titre": titre, "sous_titre": sous_titre, "notes": notes})''',
    '''PARTIE = None
VALIDITE = {  # niveau 2 : ligne « Valable pour » par défaut de chaque partie (remplacée par une mention propre si nécessaire)
    "R": "chaque chiffre dans le périmètre de sa partie d'origine, voir l'annexe « Périmètre et validité des analyses »",
    1: "textes du FPPN (statuts et règlement intérieur du 15/06/2023) et circuits décrits en entretien, sans test de leur application",
    2: "mission au 15/10/2026, documents reçus jusqu'au 22/10/2026",
    3: "fichiers de précomptes reçus le 15/10/2026, mois effectivement reçus pour chaque poste (diapositive « Couverture de la période »)",
    4: "policiers en activité précomptés par la Solde, janvier 2024 à septembre 2026 ; taux mesurés sur les mois détaillés cités",
    5: "comptes 2020 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; prestations comptabilisées, sans sinistres tardifs",
    6: "soldes et mouvements comptables de 2020 au 31/08/2026 ; épargne sans échéancier individuel ni intérêts comptabilisés",
    7: "exports du logiciel de prêts du Grenier, arrêtés du 31/12/2021 au 30/09/2026, non rapprochés des contrats",
    8: "retraités précomptés par la CGRAE de février 2024 à septembre 2026 ; produits annexes sur les mois reçus",
    9: "balances de 2020 au 31/08/2026, valeurs comptables non réévaluées ; 2026 provisoire (huit mois)",
    10: "ratios de la Décision n° 05/2022 recalculés sur les balances ; statut opposable de 2022 à 2025, indicatif en 2020, 2021 et 2026",
    11: "tables réglementaires CIMA ; non comparées à la mortalité par âge du FPPN (dates de naissance non reçues)",
    12: "méthodes décrites, non appliquées faute de données",
    13: "chaîne de traitement de la version 9",
    14: "demandes et constats au 22/10/2026",
}


def contenu(entete, notes, **k):
    if any(c in k for c in ("tableau", "graphique", "graphiques")):
        k.setdefault("valable", VALIDITE.get(PARTIE))
    k["_partie"] = PARTIE
    SL.append({"type": "contenu", "entete": entete, "notes": notes, **k})


def separateur(num, titre, sous_titre, notes=""):
    global PARTIE
    PARTIE = num
    SL.append({"type": "separateur", "num": num, "titre": titre, "sous_titre": sous_titre, "notes": notes})''')

rep('''def tableau_avec_analyse(slide, x):
    spec, paras = dict(x["tableau"]), paragraphes(x)
    if not paras:
        tableau(slide, spec)
        return''', '''def valable(slide, x, y, largeur):
    """Niveau 2 de la validité : ligne « Valable pour » sous le tableau ou le graphique."""
    if x.get("valable"):
        txt(slide, 1.4, y, largeur, 0.45, "Valable pour : " + x["valable"] + ".", 8.5, False, GRIS_F)
        return y + 0.5
    return y


def tableau_avec_analyse(slide, x):
    spec, paras = dict(x["tableau"]), paragraphes(x)
    if not paras:
        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - (0.55 if x.get("valable") else 0))
        valable(slide, x, bas + 0.05, sum(spec["largeurs"]))
        return''')
rep('''        bas = tableau(slide, spec)
        xt = 1.4 + sum(spec["largeurs"]) + 0.6
        corps(slide, xt, 3.85, 32.6 - xt, BAS_CONTENU - 3.85, paras)''', '''        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - (0.55 if x.get("valable") else 0))
        valable(slide, x, bas + 0.05, sum(spec["largeurs"]))
        xt = 1.4 + sum(spec["largeurs"]) + 0.6
        corps(slide, xt, 3.85, 32.6 - xt, BAS_CONTENU - 3.85, paras)''')
rep('''        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.3)
        corps(slide, 1.4, bas + 0.3, 31.2, BAS_CONTENU - bas - 0.3, paras, 13)''', '''        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.3 - (0.55 if x.get("valable") else 0))
        bas = valable(slide, x, bas + 0.05, 31.2)
        corps(slide, 1.4, bas + 0.3, 31.2, BAS_CONTENU - bas - 0.3, paras, 13)''')
# graphiques : la hauteur du graphique laisse la place de la ligne, écrite au post-traitement
rep('''            g.update({"x": 1.4, "y": 4.95, "largeur": lg, "hauteur": BAS_CONTENU - 4.95})''',
    '''            g.update({"x": 1.4, "y": 4.95, "largeur": lg, "hauteur": BAS_CONTENU - 4.95 - (0.55 if x.get("valable") else 0)})''')
rep('''            hg = (13.4 - 4.95) if x.get("bandeau") else (BAS_CONTENU - 4.95)''',
    '''            hg = ((13.4 - 4.95) if x.get("bandeau") else (BAS_CONTENU - 4.95)) - (0.55 if x.get("valable") else 0)''')
rep('''        if "graphique" in x:
            _, xt, lt = disposition_graphique(paras)
            corps(sl, xt, 4.0, lt, BAS_CONTENU - 4.0, paras)''', '''        if "graphique" in x:
            lg_v, xt, lt = disposition_graphique(paras)
            corps(sl, xt, 4.0, lt, BAS_CONTENU - 4.0, paras)
            valable(sl, x, BAS_CONTENU - 0.5, lg_v)
        if "graphiques" in x:
            valable(sl, x, (13.4 if x.get("bandeau") else BAS_CONTENU) - 0.5, 31.2)''')

# ------------------------------------------------------------------ corrections
rep('''"Le précompte total sur le salaire de base se répartit en 3,77 % maladie et 1,23 % décès, sur " + N("Lignes taux", T936["lignes"]) + " lignes."''',
    '''"Le précompte maladie vaut 3,77 % du salaire de base sur " + N("Lignes taux", T936["lignes"]) + " lignes de huit mois ; le précompte décès vaut 1,23 % (mesuré en avril 2024 et janvier 2025, confirmé de mars 2024 à septembre 2026)."''')
rep('''"synthese": f"La même répartition, 3,77 % et 1,23 %, se retrouve sur {N('Lignes taux (pas)', T936['lignes'])} lignes des mois détaillés"},''',
    '''"synthese": f"Le taux maladie de 3,77 % se retrouve sur {N('Lignes taux (pas)', T936['lignes'])} lignes de huit mois ; le taux décès de 1,23 % est mesuré en avril 2024 et janvier 2025, puis confirmé par le rapport 8Z6 / 936 (partie 3)"},''')
rep('''"Encours multiplié par 132, financé à 87,67 %, retards maîtrisés mais précomptes incomplets",''',
    '''"Encours multiplié par 132, financé à " + pc("Part financement (ret)", b[2026]["grenier_financement"] / GP["encours"]) + ", retards maîtrisés mais précomptes incomplets",''')
rep('''+ ", sans échéancier ; taux servi de 3,5 %", "Hausse des flux futurs (intérêts servis)", "Décision fixant le taux (non reçue)"],''',
    '''+ ", sans échéancier ; taux servi non établi (3,5 %, 5,5 % ou 14 % selon le document)", "Hausse des flux futurs (intérêts servis, non comptabilisés)", "Décision fixant le taux (non reçue)"],''')
rep('''par an de 2021 à 2025, irrégulier ; taux servi de 3,5 % sur l'épargne ; FCP non revalorisé."''',
    '''par an de 2021 à 2025, irrégulier ; taux servi non établi (3,5 %, 5,5 % ou 14 %) ; FCP non revalorisé."''')
rep('''le rendement minimal au regard du taux servi de 3,5 %."''', '''le rendement minimal au regard du taux servi, qui doit d'abord être arrêté (3,5 %, 5,5 % ou 14 % selon les documents)."''')
rep(''' ; taux servi sur l'épargne : 3,50 %"},''', ''' ; taux servi sur l'épargne : 3,50 % selon le chef comptable, 5,50 % selon le manuel d'épargne"},''')
rep('''contenu({"titre": "Règles des produits", "sous_titre": "Les fichiers confirment les montants de l'EFC et du PEP, sans texte qui les fonde",''',
    '''contenu({"titre": "Règles des produits", "sous_titre": "Les fichiers confirment les montants de l'EFC et du PEP ; les textes reçus le 22/10/2026 ne sont pas adoptés formellement",''')
rep('''"**Les fichiers confirment les déclarations sur l'EFC et le PEP, mais aucune règle n'est appuyée par un texte transmis.**''',
    '''"**Les fichiers confirment les déclarations sur l'EFC et le PEP.** Le cahier des charges EFC et PEP de 2021 et le manuel des produits d'épargne, reçus le 22/10/2026, décrivent ces produits sans décision d'adoption signée (partie 6).''')
rep('''["Méthode du FPPN", "Non communiquée", "Non communiquée : le calcul du FPPN a été demandé le 16/10/2026.", "Inconnue", "En attente"],''',
    '''["Méthode du FPPN", "Prestations payées / cotisations payées, par garantie et global, sans provisions", "Communiquée le 22/10/2026 (étude des ratios) : 12,95 % en décès, 99,23 % en maladie, 78,49 % au total au 30/06/2026.", "70 % à 80 % appliqués à chaque garantie", "Divergent du texte"],''')
rep('''"synthese": "La méthode du FPPN n'a pas été communiquée : seules la formule du texte et sa déclinaison par garantie peuvent être calculées"},''',
    '''"synthese": "La méthode du FPPN, communiquée le 22/10/2026, ignore les provisions et applique la norme globale à chaque garantie"},''')
rep('''implication="Seul le taux global est comparé à la norme de 70 % à 80 %. Le taux par garantie sert à calculer l'excédent de chaque risque. Dès que le FPPN communiquera sa méthode, l'écart avec la formule du texte sera chiffré.",''',
    '''implication="Seul le taux global est comparé à la norme de 70 % à 80 %. Le taux par garantie sert à calculer l'excédent de chaque risque. La méthode du FPPN est détaillée et chiffrée dans la diapositive « Méthode de sinistralité du FPPN ».",''')
# provision du personnel : le recalcul est fait
rep('''{F('PIDR 2025', PI[-1]['provision_196'])} en 2025 ; en attente des documents pour le recalcul"},''',
    '''{F('PIDR 2025', PI[-1]['provision_196'])} en 2025 ; recalcul agent par agent dans les diapositives suivantes"},''')
rep('''                 "**Le recalcul demande le fichier du personnel avec le statut de chaque agent (civil ou policier), l'âge, l'ancienneté et le salaire, et la note de calcul du FPPN.** Ces documents sont dans la PBC (items 76, 78, 103 et 104)."],
        implication="Analyse en attente des documents. Le montant est faible au regard des provisions techniques, mais il entre dans les dettes à long terme et dans les ratios de la partie 10.",''',
    '''                 "**Les DISA reçues le 22/10/2026 donnent l'âge, l'ancienneté et le salaire de chaque agent civil.** Le recalcul est présenté dans les diapositives suivantes."],
        implication="Le montant est faible au regard des provisions techniques, mais il entre dans les dettes à long terme et dans les ratios de la partie 10.",''')
rep('("CMP", r"\\bCMP\\b", "sigle des comptes du FPPN, signification à confirmer"),', '("CMP", r"\\bCMP\\b", "Centre Médical du Policier (rapport de gestion 2023)"),')

# ------------------------------------------------------------------ insertions
rep('contenu({"titre": "Chaîne de traitement"', bloc("bloc_validite.py") + 'contenu({"titre": "Chaîne de traitement"')
rep('separateur(4, "Effectifs et cotisations"', bloc("bloc_controles.py") + 'separateur(4, "Effectifs et cotisations"')
rep('contenu({"titre": "Nature des risques"', bloc("bloc_sinistralite.py") + 'contenu({"titre": "Nature des risques"')
rep('separateur(7, "Prêts du Grenier"', bloc("bloc_epargne.py") + 'separateur(7, "Prêts du Grenier"')
rep('contenu({"titre": "Produits annexes"', bloc("bloc_cgrae.py") + 'contenu({"titre": "Produits annexes"')
rep('# ---- v7 : conclusions des tâches « engagements » et « suffisance des provisions »', bloc("bloc_pidr.py") + '# ---- v7 : conclusions des tâches « engagements » et « suffisance des provisions »')
rep('SL.append({"type": "separateur", "num": "A", "titre": "Annexes"', bloc("bloc_points.py") + 'SL.append({"type": "separateur", "num": "A", "titre": "Annexes"')

# annexe : niveau 3 de la validité (une ligne par partie, puis chaque diapositive à mention propre)
rep('''_sig = sorted({(nom, defi) for nom, _, defi in SIGLES}, key=lambda t: t[0].lower())''', '''TITRES_P = {x["num"]: x["titre"] for x in SL if x["type"] == "separateur"}
VAL_LIGNES = [[f"Partie {p}" if p != "R" else "Résumé", TITRES_P.get(p, ""), "Toutes, sauf mention propre ci-dessous", v] for p, v in VALIDITE.items() if p in TITRES_P]
VAL_LIGNES += [[f"Partie {x['_partie']}" if x["_partie"] != "R" else "Résumé", x["entete"]["titre"], str(i), x["valable"]]
               for i, x in enumerate(SL, 1) if x["type"] == "contenu" and x.get("valable") and x["valable"] != VALIDITE.get(x.get("_partie"))]
for _k in range(0, len(VAL_LIGNES), 14):
    SL.append({"type": "annexe", "titre": f"Périmètre et validité des analyses ({_k // 14 + 1}/{(len(VAL_LIGNES) + 13) // 14})", "notes": "Niveau 3 de la validité : la ligne « Valable pour » de chaque diapositive et sa justification. Colonne « Valable pour » dans l'onglet 9_figures_tableaux du classeur v9.",
               "tableau": {"titre": None, "entetes": ["Partie", "Analyse ou diapositive", "Diapo.", "Valable pour"], "largeurs": [2.6, 9.6, 3.0, 16.0], "taille": 8,
                           "lignes": VAL_LIGNES[_k:_k + 14]},
               "source": "Source : Adenka, d'après les fichiers cités dans chaque diapositive"})
_sig = sorted({(nom, defi) for nom, _, defi in SIGLES}, key=lambda t: t[0].lower())''')
rep('''                      ["Classeur des calculs v8", "Tableaux, ratios, provisions, formules (onglet 11_formules), sensibilités (onglets 13), traçabilité", "FPPN_Calculs_rapport_20261015_v8.xlsx"]]},''',
    '''                      ["Documents reçus le 22/10/2026 (lots 1 à 3, 116 fichiers)", "États financiers, rapports de gestion, cahiers des charges, manuel d'épargne, étude des ratios, DISA, grands livres, rapprochements bancaires, conventions", "Dossier 04-Documents reçus du client ; journaux journal_lot1.csv à journal_lot3.csv"],
                      ["Classeur des calculs v9", "Tableaux, ratios, provisions, formules (onglet 11_formules), sensibilités (onglets 13), provision du personnel (14), mortalité (15), EFC (16), contrôles croisés (17), traçabilité et validité", "FPPN_Calculs_rapport_20261015_v9.xlsx"]]},''')
rep("onglet 11_formules du classeur des calculs v8.", "onglet 11_formules du classeur des calculs v9.")
rep("onglet 11_formules du classeur des calculs v7.", "onglet 11_formules du classeur des calculs v9.")
rep('["rapport_v8_fppn.py, sensibilites_fppn.py, logigramme_fppn.py, classeur_v2_fppn.py"', '["rapport_v9_fppn.py, donnees_v9_fppn.py, sensibilites_fppn.py, logigramme_fppn.py, classeur_v2_fppn.py"')
rep('"notes": "Rapport d\'étape au 15/10/2026, version 8 (analyses de sensibilité et leviers de conformité ;',
    '"notes": "Rapport d\'étape au 15/10/2026, version 9 (documents du 22/10/2026, périmètre et validité, provision du personnel recalculée, contrôles croisés ; analyses de sensibilité et leviers de conformité ;')

# index : validité des figures et des tableaux (colonne du classeur) et lignes de l'annexe
rep('''json.dump({"figures": FIG, "tableaux": TAB, "diapositives": len(SL), "trace": TRACE},''',
    '''for _l in FIG + TAB:
    _l.append(SL[_l[2] - 1].get("valable") or "")
json.dump({"figures": FIG, "tableaux": TAB, "diapositives": len(SL), "trace": TRACE, "validite": VAL_LIGNES},''')

# sigles
rep('    ("SMIG", r"\\bSMIG\\b", "salaire minimum interprofessionnel garanti"),\n',
    '    ("SMIG", r"\\bSMIG\\b", "salaire minimum interprofessionnel garanti"),\n'
    '    ("DISA", r"\\bDISA\\b", "déclaration individuelle des salaires annuels à la CNPS"),\n'
    '    ("IPS-CGRAE", r"IPS-CGRAE", "Institution de prévoyance sociale - Caisse générale de retraite des agents de l\'État"),\n'
    '    ("DG", r"\\bDG\\b", "Direction Générale du FPPN"),\n'
    '    ("DEPA", r"\\bDEPA\\b", "Direction de l\'École de Police d\'Abidjan"),\n'
    '    ("SHA-256", r"SHA-256", "algorithme d\'empreinte numérique servant à vérifier qu\'un fichier n\'a pas été modifié"),\n'
    '    ("S/P", r"\\bS/P\\b", "sinistres sur primes (cotisations)"),\n')
open(os.path.join(REPO, "rapport_v9_fppn.py"), "w", encoding="utf-8").write(s)
print("rapport_v9_fppn.py écrit,", s.count("\n"), "lignes")
