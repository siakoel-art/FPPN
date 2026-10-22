# -*- coding: utf-8 -*-
"""Construit scripts/rapport_v10_fppn.py à partir de rapport_v9_fppn.py (remplacements contrôlés et insertions)."""
import os

REPO = "/home/user/FPPN/scripts"
ICI = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(REPO, "rapport_v9_fppn.py"), encoding="utf-8").read()


def rep(a, b, n=1):
    global s
    k = s.count(a)
    assert k == n, (k, a[:100])
    s = s.replace(a, b)


def tranche(debut, fin, nouveau):
    """Remplace le texte compris entre debut (inclus) et fin (exclu)."""
    global s
    assert s.count(debut) == 1 and s.count(fin) == 1, (debut[:80], fin[:80])
    i, j = s.index(debut), s.index(fin)
    assert i < j
    s = s[:i] + nouveau + s[j:]


def omettre(titre):
    rep('contenu({"titre": "' + titre + '",', '_omis(_snap(), {"titre": "' + titre + '",')


def bloc(nom):
    return open(os.path.join(ICI, nom), encoding="utf-8").read()


# ------------------------------------------------------------------ en-tête, arguments
rep("FPPN - Rapport actuariel et statistique, version 9 (PowerPoint, charte Adenka, Timesaver)",
    "FPPN - Rapport actuariel et statistique, version 10 (PowerPoint, charte Adenka, Timesaver)")
rep("Nouveautés de la version 9 (documents reçus le 22/10/2026, périmètre et validité des analyses) :",
    """Nouveautés de la version 10 (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif) :
- résumé exécutif ouvert par la réponse aux quatre demandes de la mission et un verdict chiffré par demande ;
- provisions : méthode courte maladie, allocations décès (états de fin d'année et 41591000), risques courts,
  intérêts d'épargne dus, dépréciation du Grenier selon sa politique, santé des retraités ;
- adossement actif-passif : méthodes justifiées, écoulement de l'épargne, écarts de liquidité, résistance, gouvernance ;
- projection agrégée de l'équilibre technique avec avertissement ; annexe des méthodes non retenues ;
- validité exacte par diapositive ; renvois numérotés entre diapositives ; mentions internes retirées.

Nouveautés de la version 9 (documents reçus le 22/10/2026, périmètre et validité des analyses) :""")
rep("Utilisation : python rapport_v9_fppn.py", "Utilisation : python rapport_v10_fppn.py")
rep("<carte_sensibilites.json> <donnees_v9.json>\n", "<carte_sensibilites.json> <donnees_v9.json> <donnees_v10.json>\n")
rep("RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J = sys.argv[1:11]", "RES, COMP, RAT, IMG, SKILL, SORTIE, BASE, SENS, CARTE8, V9J, V10J = sys.argv[1:12]")

# ------------------------------------------------------------------ validité par partie, sous-parties
rep('''    12: "méthodes décrites, non appliquées faute de données",
    13: "chaîne de traitement de la version 9",''', '''    "9.2": "bilans au 31/12/2025 et au 31/08/2026 (provisoire) ; sorties d'épargne observées de 2021 à 2025",
    "9.3": "balances de 2020 au 31/08/2026 ; documents sur les actifs reçus le 22/10/2026",
    12: "comptes 2023 à 2025, garanties maladie et décès prises ensemble",
    13: "traitements des fichiers reçus jusqu'au 22/10/2026",''')
rep('''    if x["type"] == "separateur" and x["num"] != "R":
        entrees.append((x["titre"], i))''', '''    if x["type"] == "separateur" and x["num"] != "R" and "." not in str(x["num"]):
        entrees.append((x["titre"], i))''')

# ------------------------------------------------------------------ outils v10 : données communes, omission tracée
rep('SL.append({"type": "couverture"', '''def _snap():
    return (len(TRACE), len(FIG), len(TAB))


def _omis(n0, *a, **k):
    """Diapositive retirée en v10 : on efface les chiffres, figures et tableaux enregistrés pendant l'évaluation de ses arguments."""
    del TRACE[n0[0]:], FIG[n0[1]:], TAB[n0[2]:]


''' + bloc("bloc_communs.py") + '\nSL.append({"type": "couverture"')
rep('version 9 (documents du 22/10/2026, périmètre et validité, provision du personnel recalculée, contrôles croisés ;',
    'version 10 (réponse aux quatre demandes, revue actuarielle complète, adossement actif-passif ; version 9 : documents du 22/10/2026, périmètre et validité, provision du personnel recalculée, contrôles croisés ;')

# ------------------------------------------------------------------ ligne « Valable pour » : hauteur selon la longueur du texte
rep("""def _snap():""", """def _hv(x, largeur=31.2):
    \"\"\"Hauteur (cm) réservée à la ligne « Valable pour » : une ligne par tranche de largeur, 8,5 points.\"\"\"
    if not x.get("valable"):
        return 0
    return 0.40 * math.ceil(len("Valable pour : " + x["valable"] + ".") * 0.158 / largeur) + 0.15


def _snap():""")
rep("""            g.update({"x": 1.4, "y": 4.95, "largeur": lg, "hauteur": BAS_CONTENU - 4.95 - (0.55 if x.get("valable") else 0)})""",
    """            g.update({"x": 1.4, "y": 4.95, "largeur": lg, "hauteur": BAS_CONTENU - 4.95 - _hv(x, lg)})""")
rep("""            hg = ((13.4 - 4.95) if x.get("bandeau") else (BAS_CONTENU - 4.95)) - (0.55 if x.get("valable") else 0)""",
    """            hg = ((13.4 - 4.95) if x.get("bandeau") else (BAS_CONTENU - 4.95)) - _hv(x)""")
rep("""        txt(slide, 1.4, y, largeur, 0.45, "Valable pour : " + x["valable"] + ".", 8.5, False, GRIS_F)
        return y + 0.5""", """        txt(slide, 1.4, y, largeur, _hv(x, largeur) - 0.1, "Valable pour : " + x["valable"] + ".", 8.5, False, GRIS_F)
        return y + _hv(x, largeur)""")
rep("""        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - (0.55 if x.get("valable") else 0))""",
    """        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - 0.25 - _hv(x, sum(spec["largeurs"])))""", 2)
rep("""        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.3 - (0.55 if x.get("valable") else 0))""",
    """        bas = tableau(slide, spec, hauteur_max=BAS_CONTENU - 3.85 - 0.75 - besoin - 0.55 - _hv(x))""")
rep("""        valable(slide, x, bas + 0.05, sum(spec["largeurs"]))""", """        valable(slide, x, bas + 0.3, sum(spec["largeurs"]))""", 2)
rep("""        bas = valable(slide, x, bas + 0.05, 31.2)""", """        bas = valable(slide, x, bas + 0.3, 31.2)""")
rep("""            valable(sl, x, BAS_CONTENU - 0.5, lg_v)""", """            valable(sl, x, BAS_CONTENU - _hv(x, lg_v), lg_v)""")
rep("""            valable(sl, x, (13.4 if x.get("bandeau") else BAS_CONTENU) - 0.5, 31.2)""", """            valable(sl, x, (13.4 if x.get("bandeau") else BAS_CONTENU) - _hv(x), 31.2)""")

# ------------------------------------------------------------------ résumé exécutif : réponse aux quatre demandes en tête
rep('contenu({"titre": "Résumé exécutif",', bloc("bloc_verdicts.py") + 'contenu({"titre": "Résumé exécutif",')
omettre("Réponse aux quatre questions de la mission")
rep('''boite(slide, 16.8, 6.5, 6.6, 4.0, "FPPN\\nCotisations " + F("Cot (schéma)"''', '''boite(slide, 16.8, 6.5, 6.6, 4.0, "FPPN\\nCotisations 2025\\n" + F("Cot (schéma)"''')

# ------------------------------------------------------------------ partie 2 : documents reçus, validité
tranche('contenu({"titre": "Documents reçus le 22/10/2026"', 'contenu({"titre": "Chaîne de traitement"', bloc("bloc_documents.py"))
rep('SRC_22 = "Source : documents du FPPN reçus le 22/10/2026 (lots 1 à 3, contrôlés par empreinte SHA-256)"', 'SRC_22 = "Source : documents du FPPN reçus le 22/10/2026"')
rep('valable="les familles d\'analyses du rapport version 9, sur les documents reçus jusqu\'au 22/10/2026",', 'valable="les familles d\'analyses du rapport, sur les documents reçus jusqu\'au 22/10/2026",')

rep('        source=SRC_COMPTA + " ; " + SRC_PREC[9:] + " ; " + SRC_GRENIER[9:])',
    '        source="Source : balances des comptes du FPPN de 2020 au 31/08/2026 ; fichiers de précomptes et exports du Grenier reçus le 15/10/2026 (noms exacts en annexe « Fichiers sources ») ; traitements Adenka")')

# ------------------------------------------------------------------ partie 6 : intérêts dus
rep('contenu({"titre": "Rapport de gestion 2025 : l\'épargne présentée', bloc("bloc_epargne_interets.py") + 'contenu({"titre": "Rapport de gestion 2025 : l\'épargne présentée')
# ------------------------------------------------------------------ partie 8 : retraités
rep('contenu({"titre": "Produits annexes"', bloc("bloc_retraites.py") + 'contenu({"titre": "Produits annexes"')

# ------------------------------------------------------------------ partie 9 : provisions
rep('''["Maladie", "PSAP dossier par dossier et sinistres tardifs (art. 10)", F("PSAP maladie (p9)", PV["psap_maladie"]["2025"]), "Montant hérité d'environ 2004 ; aucune provision maladie calculée, factures rattachées jusqu'en mars (chef comptable, 19/10/2026)."],''',
    '''["Maladie", "PSAP dossier par dossier et sinistres tardifs (art. 10)", F("PSAP maladie (p9)", PV["psap_maladie"]["2025"]), "Montant hérité d'environ 2004, non calculé ; besoin mesuré de quelques millions de FCFA (méthode courte)."],''')
rep('''"Montant hérité d'environ 2004, inchangé ; calcul annuel des allocations restant à payer, compte à préciser."],''',
    '''"Montant hérité d'environ 2004, inchangé ; le calcul annuel des allocations restant à payer est passé au compte 41591000."],''')
rep('''"Cotisation mensuelle : période non couverte d'un mois au plus, montant probablement faible."],''', '''"Cotisation précomptée sur le mois qu'elle couvre : provision nulle (statuts, art. 88)."],''')
rep('''"Règlement de l'EFC et du PEP non reçu : besoin non mesurable."],''', '''"Intérêts dus non comptabilisés : " + F("PPP int", INT55) + " au 31/08/2026 (PEP 3,5 %, EFC 5,5 %)."],''')
rep('''["Prêts du Grenier", "Dépréciation des créances (PCMS, prudence)", "Aucune depuis 2022", "Encours en retard connu par tranche, barème à appliquer."],''',
    '''["Prêts du Grenier", "Dépréciation des créances (PCMS, prudence)", "Aucune depuis 2022", "Dépréciation calculée avec la politique de crédit du Grenier (40 %, 80 %, 100 %)."],''')
rep('''        implication="Le recalcul de la PSAP, dossier par dossier puis par la méthode CIMA des sinistres tardifs, est la première tâche dès réception des bases de prestations. D'ici là, le résultat de chaque exercice ne tient compte d'aucune variation de provision.",''',
    '''        implication="Les trois diapositives suivantes mesurent la provision maladie nécessaire avec la méthode courte que permettent le règlement intérieur et les factures reçues.",''')
rep('contenu({"titre": "Allocations décès : méthode du FPPN"', bloc("bloc_psap_maladie.py") + 'contenu({"titre": "Allocations décès : méthode du FPPN"')
rep('''contenu({"titre": "Allocations décès : méthode du FPPN", "sous_titre": "Les dossiers non payés sont valorisés au barème, mais le compte utilisé n'est pas connu",
         "synthese": "Question ouverte : où le calcul de fin d'année des allocations restant à payer est-il comptabilisé ?"},
        "Logigramme modifiable, tel que décrit par le chef comptable le 19/10/2026 (à confirmer par document). Barème déclaré : sous-officier 1 000 000 FCFA, commissaire 2 000 000 FCFA, officier à confirmer.",''',
    '''contenu({"titre": "Allocations décès : méthode du FPPN", "sous_titre": "Les dossiers non payés sont valorisés au barème et passés en dette au compte 41591000",
         "synthese": "Les états de fin d'année reçus le 22/10/2026 confirment le circuit décrit en entretien et en donnent le compte"},
        "Logigramme modifiable, tel que décrit par le chef comptable le 19/10/2026 et confirmé par les états de fin d'année 2021 à 2025. Barème appliqué dans les états : sous-officier 1 400 000 FCFA, officier 1 600 000 FCFA, commissaire 2 200 000 FCFA (le barème déclaré en entretien, 1 000 000 et 2 000 000 FCFA, n'est pas celui des fichiers).",''')
rep('''                 "**Question ouverte : le compte de provision décès (15802000) ne bouge pas depuis 2020.** Le rapport ne peut donc pas dire où ce calcul annuel est comptabilisé, ni s'il l'est.",''',
    '''                 "**Le calcul annuel est comptabilisé en dette au compte 41591000, non au compte de provision 15802000**, qui ne bouge pas depuis 2020 (diapositive suivante).",''')
rep('\ncontenu({"titre": "Méthode CIMA des sinistres tardifs"', '\n' + bloc("bloc_deces.py") + 'contenu({"titre": "Méthode CIMA des sinistres tardifs"')
omettre("Méthode CIMA des sinistres tardifs")
omettre("Dépréciation des prêts du Grenier")
tranche('contenu({"titre": "Conclusion : engagements envers les adhérents"', 'contenu({"titre": "Structure de l\'actif"',
        bloc("bloc_conclusions_prov.py") + '\nseparateur("9.2", "Adossement actif-passif", "Liquidité, écoulement de l\'épargne, tests de résistance, gouvernance")\n')
tranche('contenu({"titre": "Sorties d\'épargne observées"', '\nfp = EX["fonds_propres_2025"]',
        bloc("bloc_alm.py") + '\nseparateur("9.3", "Fonds propres et placements", "Fonds propres, placements, actifs immobilisés, adéquation aux engagements")\n')
rep('\ncontenu({"titre": "Rendement des placements"', '\n' + bloc("bloc_placements.py") + 'contenu({"titre": "Rendement des placements"')
rep('''                 "**Le FPPN sert 3,5 % par an sur le PEP et l'EFC** (chef comptable, 19/10/2026) : ce taux est un engagement de rendement, à couvrir par les produits des placements.",''',
    '''                 "**Le taux servi sur l'épargne (3,5 % sur le PEP, 3,5 % ou 5,5 % sur l'EFC selon le document) est un engagement de rendement** que les produits des placements ne couvrent pas (partie 6).",''')
tranche('# ---- v7 : adéquation des placements (article 17)', 'retenir("bilan"', bloc("bloc_concl_placements.py") + "\n")

# ------------------------------------------------------------------ partie 12 : projection ; annexe des méthodes non retenues
tranche('# ---------------------------------------------------------------- PARTIE 12', '# ---------------------------------------------------------------- PARTIE 13', bloc("bloc_projection.py"))
rep('''SL.append({"type": "separateur", "num": "A", "titre": "Annexes", "sous_titre": "Méthodes, listes, abréviations, références", "notes": ""})''',
    '''separateur("A", "Annexes", "Méthodes non retenues, validité, listes, abréviations, références")
''' + bloc("bloc_annexe_methodes.py") + bloc("bloc_annexe_sources.py"))
rep('SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 ; traitements Adenka"',
    'SRC_PREC = "Source : fichiers de précomptes de la Solde, de la CGRAE et du FPPN reçus le 15/10/2026 (dossiers « FPPN-SOLDE », « FPPN-CGRAE », « FPPN-GNA », noms exacts en annexe « Fichiers sources ») ; traitements Adenka"')
rep('SRC_COMPTA = "Source : balances des comptes du FPPN (Sage 100), exercices 2020 à 2025 et du 01/01 au 31/08/2026 (provisoire) ; traitements Adenka"',
    'SRC_COMPTA = "Source : balances des comptes du FPPN de 2020 au 31/08/2026 (provisoire), fichiers « BALANCE … » listés en annexe « Fichiers sources » ; traitements Adenka"')
rep('SRC_GRENIER = "Source : exports du logiciel de prêts du Grenier du Policier (balances âgées, production) reçus le 15/10/2026 ; traitements Adenka"',
    'SRC_GRENIER = "Source : exports du logiciel de prêts du Grenier du Policier « BalanceAgeeDetaille … .xls » et « production_de_la_periode … .xls » reçus le 15/10/2026 ; traitements Adenka"')
rep("résumé exécutif, compréhension du FPPN, démarche, données, analyses par thème, conformité prudentielle, tables de mortalité, méthodes à venir, outils, analyses en attente, annexes.",
    "résumé exécutif (réponse aux quatre demandes), compréhension du FPPN, démarche, données, analyses par thème, adossement actif-passif et placements, conformité prudentielle, tables de mortalité, projection de l'équilibre technique, outils, analyses en attente, annexes.")

# ------------------------------------------------------------------ validité exacte, renvois, mentions internes, annexe de validité (une ligne par diapositive)
rep('''TITRES_P = {x["num"]: x["titre"] for x in SL if x["type"] == "separateur"}
VAL_LIGNES = [[f"Partie {p}" if p != "R" else "Résumé", TITRES_P.get(p, ""), "Toutes, sauf mention propre ci-dessous", v] for p, v in VALIDITE.items() if p in TITRES_P]
VAL_LIGNES += [[f"Partie {x['_partie']}" if x["_partie"] != "R" else "Résumé", x["entete"]["titre"], str(i), x["valable"]]
               for i, x in enumerate(SL, 1) if x["type"] == "contenu" and x.get("valable") and x["valable"] != VALIDITE.get(x.get("_partie"))]''',
    bloc("bloc_validite_v10.py") + '''TITRES_P = {x["num"]: x["titre"] for x in SL if x["type"] == "separateur"}
VAL_LIGNES = [[("Partie " + str(x["_partie"])) if x["_partie"] not in ("R", "A") else ("Résumé" if x["_partie"] == "R" else "Annexe"), x["entete"]["titre"], str(i), x["valable"]]
              for i, x in enumerate(SL, 1) if x["type"] == "contenu" and x.get("valable")]''')
rep('''"notes": "Niveau 3 de la validité : la ligne « Valable pour » de chaque diapositive et sa justification. Colonne « Valable pour » dans l'onglet 9_figures_tableaux du classeur v9.",''',
    '''"notes": "Niveau 3 de la validité : la ligne « Valable pour » de chaque diapositive. Colonne « Valable pour » dans l'onglet 9_figures_tableaux du classeur des calculs.",''')
exec(open(os.path.join(ICI, "patch_mentions.py"), encoding="utf-8").read())
open(os.path.join(REPO, "rapport_v10_fppn.py"), "w", encoding="utf-8").write(s)
print("rapport_v10_fppn.py écrit,", s.count("\n"), "lignes")
