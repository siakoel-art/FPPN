# -*- coding: utf-8 -*-
# v14, étape 8 (10/11/2026) : documents de la dataroom du 09/11/2026 (lots A et B), retouches et commentaires de
# l'utilisateur sur le rapport du 09/11/2026. Tout ajout ou correction est écrit en bleu (passe finale).
#  - retouches de texte de l'utilisateur (GNA, défauts de données, Directeur Général, « Dispositif actuariel ») ;
#  - commentaires : diapositive 7 (actifs mal évalués), diapositives 180-181 (cohérence des titres des ratios) ;
#  - nouvelles diapositives : documents reçus le 09/11/2026, méthode du FPPN pour la provision du personnel,
#    ratios calculés par le FPPN, actifs dont la valeur n'est pas établie, engagements et projets hors bilan,
#    états financiers du Grenier ; barème décès complet ; questions au FPPN ; limites ;
#  - passe finale : positions retouchées par l'utilisateur (assets/retouches_utilisateur_20261110.json) et texte en bleu.
# Exécuté par patch_v12.py après etape7.py.

# ---------------------------------------------------------------- registre du bleu
_BLEU14 = r'''# ---- v14 : ajouts et corrections du 10/11/2026 écrits en bleu
BLEU = "0070C0"
BLEU_T = set()   # titres de diapositives entièrement nouvelles
BLEU_P = []      # morceaux de texte nouveaux ou corrigés dans des diapositives existantes


def bl(t):
    """Enregistre un texte nouveau ou corrigé (écrit en bleu à la fin) et le renvoie tel quel."""
    for m_ in re.split(r"\*\*|\[\[|\]\]", str(t)):
        if len(m_.strip()) >= 4:
            BLEU_P.append(m_.strip())
            BLEU_P.append(_pv14(m_.strip()))
    return t


def _pv14(t):
    """Même transformation que le rendu (pas de point-virgule dans les diapositives)."""
    return re.sub(r" ; (\w)", lambda m_: ". " + m_.group(1).upper(), str(t))


BLEU_L = []      # lignes de tableau entièrement nouvelles (repérées par leur cellule la plus longue)


def bll(t):
    """Ligne de tableau entièrement nouvelle : toute la ligne sera en bleu."""
    BLEU_L.append(str(t))
    BLEU_L.append(_pv14(t))
    return t


def blt(entete):
    """Diapositive entièrement nouvelle : tout son texte sera en bleu."""
    BLEU_T.add(entete["titre"])
    return entete'''
rep('RT = {x["code"]: x for x in Q["ratios"]}', 'RT = {x["code"]: x for x in Q["ratios"]}; exec(' + repr(_BLEU14) + ')')

# ---------------------------------------------------------------- retouches de texte de l'utilisateur
rep("primes pour le GNA", "primes pour GNA")
rep("124 fichiers reçus intacts", "124 fichiers reçus")
rep("Quatre défauts restent ouverts : décembre 2024", "Quatre défauts restent identifiés et non traités : Décembre 2024")
s = s.replace('"Clinique, projet de Cocody et désignation du directeur général"', '"Clinique, projet de Cocody et désignation du Directeur Général"')
s = s.replace('ref("Clinique, projet de Cocody et désignation du directeur général")', 'ref("Clinique, projet de Cocody et désignation du Directeur Général")')

# ---------------------------------------------------------------- barème décès complet (document reçu le 09/11/2026)
rep('''"Barème appliqué dans les fichiers : membre sous-officier 1 400 000 FCFA, officier 1 600 000, commissaire 2 200 000 ; conjoint 600 000 à 700 000 ; enfant 120 000 à 150 000. C'est le barème de la cartographie des risques, non celui déclaré en entretien le 19/10/2026.",''',
    '''"Barème appliqué dans les fichiers : membre sous-officier 1 400 000 FCFA, officier 1 600 000, commissaire 2 200 000 ; conjoint 600 000 à 1 000 000 ; enfant 120 000 à 230 000. C'est le barème de la cartographie des risques, non celui déclaré en entretien le 19/10/2026.",''')

rep('''analyse=["**En 2024 et 2025, le crédit de l'année au compte 41591000 est exactement le total du fichier.**''',
    '''analyse=[bl("**Barème officiel de 2021 à 2026, reçu le 09/11/2026 :** membre 1 400 000, 1 600 000 et 2 200 000 FCFA, conjoint 600 000, 700 000 et 1 000 000 FCFA, enfant 120 000, 150 000 et 230 000 FCFA, selon le corps (sous-officier, officier, commissaire). Frais funéraires de 300 000 FCFA pour tous depuis 2024."), "**En 2024 et 2025, le crédit de l'année au compte 41591000 est exactement le total du fichier.**''')
rep('''« GRAND LIVRE FPPN 2024.xlsx » (compte 41591000) ; « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx » ;''',
    '''« GRAND LIVRE FPPN 2024.xlsx » (compte 41591000) ; « CARTOGRAPHIE DES RISQUES COUVERTS ET DES PRESTATIONS OFFERTES.docx » ; « OFFRES ET BAREMES DES ALLOCATIONS ; CAPITAUX DECES ET FRAIS FUNERAIRES.docx » ;''')

# ---------------------------------------------------------------- provision du personnel : méthode du FPPN confirmée
rep('implication="Le compte 196 couvre la provision recalculée selon les deux approches. [[La méthode du FPPN est celle des droits acquis]] : sa note de calcul est demandée (PBC, item 113).",',
    'implication="Le compte 196 couvre la provision recalculée selon les deux approches. " + bl("La méthode du FPPN est celle des droits acquis : ses fichiers de calcul de 2020 à 2025, reçus le 09/11/2026, le confirment") + " (" + ref("Provision du personnel : méthode du FPPN et compte 196") + ").",')

_PIDR_FPPN = r'''
# ---- v14 : fichiers de calcul de la provision retraite du FPPN (reçus le 09/11/2026)
_PF = {2020: (45, 24155943, 24155943), 2021: (41, 22342542, 22542542), 2022: (64, 28325794, 28325794), 2023: (62, 35312179, 35312179),
       2024: (61, 34753547, 34753547), 2025: (78, 43666678, 43666678)}
_PSYa = {x["annee"]: x for x in PSY}
contenu(blt({"titre": "Provision du personnel : méthode du FPPN et compte 196", "sous_titre": "Le compte 196 est le calcul du FPPN en droits acquis, au franc près de 2022 à 2025",
         "synthese": "Le FPPN calcule l'indemnité de chaque agent sur son ancienneté au 31/12, comme s'il partait ce jour-là : sans actualisation, ni mortalité, ni départ anticipé"}),
        "Fichiers « FICHIER PERSONNEL CIVIL FPPN POUR PROVISION CONGES ET RETRAITE EN 2020 » à « … EN 2025 », onglet RETRAITE : ancienneté, salaire mensuel moyen, indemnité par tranche (30 %, 35 %, 40 %), total et écriture de dotation au compte 196.",
        tableau={"titre": tab("Provision du personnel : calcul du FPPN, compte 196 et recalcul Adenka"), "entetes": ["Exercice", "Agents (FPPN)", "Calcul du FPPN", "Compte 196", "Écart", "Agents (Adenka)", "Droits acquis (Adenka)", "Provision recalculée (Adenka)"],
                 "largeurs": [2.4, 2.8, 4.4, 4.4, 3.2, 3.0, 5.2, 5.6], "taille": 8.5,
                 "lignes": [[str(a_), str(v_[0]), nb(v_[1]) + " FCFA", nb(v_[2]) + " FCFA", nb(v_[2] - v_[1]) + " FCFA",
                             N(f"PF agents {a_}", _PSYa[a_]["agents"]) if a_ in _PSYa else "n.d.",
                             F(f"PF acquis {a_}", _PSYa[a_]["acquis"]) if a_ in _PSYa else "n.d.",
                             F(f"PF dbo {a_}", _PSYa[a_]["dbo"]) if a_ in _PSYa else "n.d."] for a_, v_ in _PF.items()]},
        disposition="dessous",
        analyse=["**De 2022 à 2025, le compte 196 est exactement le total du calcul du FPPN.** En 2021, le compte dépasse le fichier de 200 000 FCFA ; en 2020, ils sont égaux.",
                 f"**En 2025, les droits acquis recalculés par Adenka ({F('PF acq 2025 b', _PSYa[2025]['acquis'])}, {N('PF ag 2025 b', _PSYa[2025]['agents'])} agents) sont à {pc('PF écart acquis', (43666678 - _PSYa[2025]['acquis']) / 43666678)} du calcul du FPPN (43 666 678 FCFA, 78 agents).** La méthode est la même ; l'écart vient d'un agent de plus et des salaires retenus."],
        implication="La méthode des droits acquis surestime l'engagement par rapport aux unités de crédit projetées (" + F("PF dbo 2025 c", _PSYa[2025]["dbo"]) + " en 2025), car elle ignore l'actualisation, la mortalité et les départs avant la retraite. Le FPPN est prudent, mais sa méthode n'est pas celle des normes actuarielles.",
        valable="personnel civil du FPPN, exercices 2020 à 2025 ; calcul Adenka sur les DISA 2021, 2022, 2023 et 2025",
        source="Source : fichiers « FICHIER PERSONNEL CIVIL FPPN POUR PROVISION CONGES ET RETRAITE EN 2020.xlsx » à « … EN 2025.xlsx », reçus le 09/11/2026 ; balances 2020 à 2025 (compte 19600000) ; calculs Adenka (onglets 14b et 14c du classeur)")

'''
rep("\n\ndef _lignes_pidr(liste):", _PIDR_FPPN + "\n\ndef _lignes_pidr(liste):")

# ---------------------------------------------------------------- documents reçus le 09/11/2026
_DOC0911 = r'''contenu(blt({"titre": "Documents reçus le 09/11/2026", "sous_titre": "980 fichiers nouveaux ou modifiés, contrôlés un par un, qui confirment les chiffres et ajoutent des faits nouveaux",
         "synthese": "Les balances 2020 à 2026 sont inchangées ; les nouveaux documents précisent la provision du personnel, le barème décès, les ratios du FPPN, les actifs et la gouvernance"}),
        "Les deux dossiers de la dataroom ont été inventoriés (5 366 fichiers) et comparés, par empreinte SHA-256, aux fichiers déjà reçus. Les 980 fichiers utiles au volet actuariel ont été transmis et contrôlés par empreinte : aucun n'a été altéré.",
        tableau={"titre": tab("Documents reçus le 09/11/2026 et usage dans le rapport"), "entetes": ["Famille", "Documents", "Ce qu'ils apportent"], "largeurs": [3.6, 13.6, 14.0], "taille": 8,
                 "lignes": [["Comptes", "Balances et grands livres 2020 à 2026 (nouveau tirage) ; balance générale au 31/08/2026 refaite", "Soldes identiques compte par compte de 2020 à 2025. Balance 2026 refaite : quatre reclassements, dont 816 670 FCFA du compte 41191000 (épargne) au compte 41591000 (allocations décès)"],
                            ["Provisions", "Fichiers de calcul de la provision pour congés et retraite 2020 à 2025", "Méthode du FPPN pour la provision du personnel (" + ref("Provision du personnel : méthode du FPPN et compte 196") + ")"],
                            ["Prestations", "Offres et barèmes des allocations décès et frais funéraires ; cartographie des risques", "Barème décès complet ; frais funéraires de 300 000 FCFA depuis 2024"],
                            ["Ratios", "Indicateurs et ratios prudentiels calculés par le FPPN, 2022 à 2026", "Comparaison avec le calcul Adenka (" + ref("Ratios prudentiels : calculs du FPPN et d'Adenka") + ")"],
                            ["Gouvernance et contrôle", "PV des assemblées générales et des conseils d'administration de 2020 à 2026, rapports du commissaire aux comptes de 2022 à 2025, courriers de l'AIRMS", "Résultats et réserves confirmés ; actifs dont la valeur n'est pas établie ; procès sur l'EFC ; projets (" + ref("Actifs dont la valeur n'est pas établie") + ")"],
                            ["Grenier", "États financiers 2022 à 2025, balances, grands livres, production, documents prioritaires", "Comptes propres du Grenier (" + ref("Grenier du Policier : ses propres états financiers") + ")"],
                            ["Cotisations", "Fichiers aller-retour CGRAE et GNA modifiés ; tableaux de centralisation aller-retour", "Lignes ajoutées dans douze fichiers CGRAE et GNA ; écarts mensuels aller-retour à rapprocher des écarts du rapport"]]},
        disposition="dessous",
        analyse=["**Les excédents votés en assemblée générale sont ceux des balances :** 2 083 908 202 FCFA (2021), 2 060 332 677 FCFA (2022), 930 629 102 FCFA (2023) et 792 069 641 FCFA (2024). L'AIRMS a calculé pour 2022 la même sinistralité (9,06 %) et les mêmes frais de fonctionnement (47,35 %) que le rapport."],
        implication="Gouvernance : l'AIRMS a jugé nulle la prorogation de deux ans des mandats votée le 17/04/2026 et a mis le FPPN en demeure de transmettre ses comptes 2025. Ces points relèvent du volet juridique ; ils ne changent pas les calculs actuariels.",
        valable="fichiers des dossiers « 03-Adenka-Audit » et « 05-Adenka Audit Juridique et Fiscal » de la dataroom, téléchargés le 09/11/2026",
        source="Source : dataroom du FPPN au 09/11/2026 ; inventaire « FPPN_Inventaire_dataroom_20261109.xlsx » ; PV des AGO des 16/06/2022, 15/06/2023, 26/06/2024 et 30/06/2025 ; courriers de l'AIRMS du 29/09/2023, du 29/05/2026 et de septembre 2026")
'''
rep('contenu({"titre": "Chaîne de traitement"', _DOC0911 + 'contenu({"titre": "Chaîne de traitement"')

# ---------------------------------------------------------------- ratios calculés par le FPPN
_RAT_FPPN = r'''# ---- v14 : ratios calculés par le FPPN (document « INDICATEURS ET RATIOS PRUDENTIELS UEMOA. ACTUALISES 2 AA »)
def _rv(c_, a_, pct=False, dec=2, corr=False):
    v_ = (RT[c_].get("valeurs_retraitees") or {}).get(a_) if corr else RT[c_]["valeurs"].get(a_)
    if v_ is None:
        return "n.d."
    return pc(f"RF {c_} {a_} {corr}", v_) if pct else N(f"RF {c_} {a_} {corr}", v_, dec)


_RF = [  # code, libellé, norme, valeurs du FPPN 2024, 2025, 2026 (1er semestre), pourcentage, lecture corrigée, explication
    ("SIN", "Taux de sinistralité", "70 % à 80 %", "40,34 %", "57,48 %", "82,87 %", True, False, "Identique en 2024 et 2025 ; 2026 : périodes différentes"),
    ("CHF", "Charges de fonctionnement / cotisations", "Au plus 20 %", "36,06 %", "36,33 %", "45,07 %", True, False, "Identique à 0,01 point près"),
    ("CHB", "Charges de fonctionnement brut / produits", "Au plus 15 %", "32,20 %", "34,86 %", "44,87 %", True, False, "Identique en 2025 ; écart en 2024 non expliqué"),
    ("CHP", "Charges de personnel / charges de fonctionnement", "Au plus 85 %", "24,8 %", "26,57 %", "23,91 %", True, False, "Identique en 2024 et 2025"),
    ("AUT", "Autonomie financière", "Au moins 1", "1,29", "1,051", "0,78", False, False, "Identique"),
    ("ROB", "Réserve obligatoire (mois de prestations)", "6 à 9 mois", "5 mois", "9 mois", "n.d.", False, False, "Calcul du FPPN non détaillé : 9 mois en 2025 contre 5,01 mois"),
    ("LIQ", "Liquidité", "Au moins 1", "0,31", "1,006", "1,16", False, True, "Le FPPN retraite l'épargne en dette longue, comme la lecture corrigée"),
    ("FEP", "Fonds propres et assimilés / dettes", "Au moins 1", "0,40", "0,31", "0,20", False, False, "Le FPPN classe l'épargne non échue en dettes à plus d'un an"),
    ("EQS", "Équilibre stable", "Au plus 1", "0,48", "0,38", "0,36", False, True, "Retraitement des ressources stables non détaillé"),
    ("EXN", "Taux d'excédent net", "Au plus 5 %", "25,18 %", "12,81 %", "27,32 %", True, False, "Identique en 2024 et 2025"),
    ("TFP", "Financement propre", "Au moins 100 %", "132,48 %", "113,47 %", "78,65 %", True, False, "Identique en 2024 et 2025"),
]
contenu(blt({"titre": "Ratios prudentiels : calculs du FPPN et d'Adenka", "sous_titre": "Mêmes sinistralité et frais, mais réserve obligatoire, liquidité et fonds propres divergents",
         "synthese": "Ratios du FPPN de 2024, 2025 et du 1er semestre 2026, comparés aux ratios recalculés par Adenka sur les balances (2026 : au 31/08/2026)"}),
        "Colonnes « FPPN » : valeurs du document du FPPN, reçu le 09/11/2026. Colonnes « Adenka » : valeurs du tableau de bord ; lecture corrigée selon le PCMS pour la liquidité et l'équilibre stable, que le FPPN retraite aussi.",
        tableau={"titre": tab("Ratios prudentiels selon le FPPN et selon Adenka"), "entetes": ["Ratio", "Norme", "FPPN 2024", "Adenka 2024", "FPPN 2025", "Adenka 2025", "FPPN 2026 (S1)", "Adenka 31/08/2026", "Lecture de l'écart"],
                 "largeurs": [5.4, 2.6, 2.3, 2.3, 2.3, 2.3, 2.5, 2.6, 8.9], "taille": 7.5,
                 "lignes": [[lib_, nor_, f24_, _rv(c_, 2024, pct_, corr=cor_), f25_, _rv(c_, 2025, pct_, corr=cor_), f26_, _rv(c_, 2026, pct_, corr=cor_), exp_]
                            for c_, lib_, nor_, f24_, f25_, f26_, pct_, cor_, exp_ in _RF]},
        disposition="dessous",
        analyse=["**Même avec les chiffres du FPPN, la sinistralité 2025 (57,48 %), les fonds propres (0,31) et les frais de fonctionnement (36,33 %) restent hors norme.** Le verdict de non-conformité ne change pas.",
                 "**Le FPPN trouve 9 mois de réserve obligatoire en 2025, contre 5,01 mois recalculés sur les balances.** Le compte 11100000 porte 820 245 819 FCFA de réserves légales fin 2025 : la base de prestations retenue par le FPPN doit être demandée."],
        implication="Les ratios du FPPN et ceux du rapport doivent être réconciliés avant toute transmission à l'AIRMS : le FPPN doit transmettre le détail de ses calculs, notamment la réserve obligatoire et le retraitement des ressources stables.",
        valable="ratios de la Décision n° 05/2022 : valeurs du FPPN pour 2024, 2025 et le 1er semestre 2026 ; valeurs Adenka de 2024, 2025 et au 31/08/2026 (balance provisoire)",
        source="Source : document du FPPN « INDICATEURS ET RATIOS PRUDENTIELS UEMOA. ACTUALISES 2 AA.docx », reçu le 09/11/2026 ; " + SRC_REG[9:])
'''
rep('contenu({"titre": "Liquidité et équilibre : deux lectures"', _RAT_FPPN + 'contenu({"titre": "Liquidité et équilibre : deux lectures"')

# ---------------------------------------------------------------- actifs dont la valeur n'est pas établie ; engagements hors bilan
_ACTIFS = r'''# ---- v14 : actifs dont la valeur n'est pas établie et engagements hors bilan (documents du 09/11/2026)
_ACT_CONT = 375000000 + 3407519850
contenu(blt({"titre": "Actifs dont la valeur n'est pas établie", "sous_titre": "Deux terrains sans titre et deux créances sans accord de remboursement pèsent sur la valeur de l'actif",
         "synthese": f"Sans le terrain d'Attoban et la créance sur l'Hôpital de la Police, les fonds propres baisseraient de {nb(_ACT_CONT)} FCFA"}),
        "Les actifs ci-dessous figurent au bilan à leur coût. Les documents reçus le 09/11/2026 montrent que leur valeur réalisable n'est pas établie.",
        tableau={"titre": tab("Actifs dont la valeur n'est pas établie au 31/08/2026"), "entetes": ["Actif", "Montant", "Situation", "Document"], "largeurs": [6.4, 4.4, 13.4, 7.0], "taille": 8,
                 "lignes": [["Terrain « Terre rouge », Cocody Riviera Attoban (4 ha)", "375 000 000 FCFA", "Pas de titre foncier. Le vendeur a perdu son procès contre les propriétaires terriens et sa société (MAXFOR Développement) est en liquidation. Le commissaire aux comptes demande de sortir cet actif du patrimoine", "Rapports du commissaire aux comptes, exercices 2022 à 2025 ; PV du CA du 06/05/2025"],
                            ["Créance sur l'Hôpital de la Police nationale (compte 47111100)", F("AV hpn", b[2026]["creance_hopital_police"]), "Inchangée depuis 2020. Selon la direction générale, ces dépenses pourraient être des subventions : le commissaire aux comptes demande alors de les passer en charges", "Rapport du commissaire aux comptes au CA, exercice 2024 ; PV du CA du 06/05/2025"],
                            ["Terrain SONEB, Songon (10 ha)", "450 000 000 FCFA", "Acheté en 2023, sans titre foncier. Remplacé en 2026 par une parcelle de 9 ha à Jacqueville, dotée d'une lettre d'attribution de 2021", "Protocole d'accord du 16/02/2026"],
                            ["Créance sur l'État (soins des policiers blessés en service commandé)", "3 932 124 730 FCFA", "Recouvrement demandé avec l'appui de l'AIRMS ; comité de recouvrement créé. Le lien avec la créance sur l'Hôpital de la Police n'est pas établi", "Courrier de l'AIRMS du 05/01/2026 ; PV du CA de mars 2026"]]},
        disposition="dessous",
        analyse=[f"[[**Si le terrain d'Attoban et la créance sur l'Hôpital de la Police sortaient de l'actif, les fonds propres passeraient de {F('AV fp25', b[2025]['fonds_propres'])} à {F('AV fp25s', b[2025]['fonds_propres'] - _ACT_CONT)} au 31/12/2025, et de {F('AV fp26', b[2026]['fonds_propres'])} à {F('AV fp26s', b[2026]['fonds_propres'] - _ACT_CONT)} au 31/08/2026.**]] Ce calcul reste une sensibilité tant que le CA n'a pas décidé.",
                 "**Certains actifs ne sont donc pas correctement évalués.** Leur valeur au bilan surestime ce que le FPPN pourrait mobiliser pour payer l'épargne."],
        implication="Le FPPN doit faire évaluer ses terrains et sa clinique, obtenir un accord écrit sur les créances, et déprécier ce qui n'est pas recouvrable. L'adossement et les fonds propres du rapport sont des maxima.",
        valable="bilan au 31/12/2025 et au 31/08/2026 (balance provisoire) ; documents reçus le 09/11/2026",
        source="Source : rapports du commissaire aux comptes (CLS Audit & Conseil) sur les exercices 2022 à 2025 ; PV des CA des 06/05/2025 et 02/03/2026 ; protocole SONEB du 16/02/2026 ; courrier de l'AIRMS du 05/01/2026 ; " + SRC_COMPTA[9:])
contenu(blt({"titre": "Engagements et projets hors bilan", "sous_titre": "Un procès perdu, un projet immobilier de 19,5 milliards FCFA et une prise de participation",
         "synthese": "Ces éléments ne figurent pas, ou pas entièrement, dans les comptes analysés ; ils réduisent la trésorerie disponible pour l'épargne"}),
        "Faits relevés dans les procès-verbaux du conseil d'administration et de l'assemblée générale, et dans les documents de financement reçus le 09/11/2026.",
        tableau={"titre": tab("Engagements et projets hors bilan relevés dans les documents"), "entetes": ["Élément", "Montant", "Ce que disent les documents", "Effet pour l'adossement"], "largeurs": [5.2, 4.8, 12.0, 9.2], "taille": 8,
                 "lignes": [["Procès perdu sur l'EFC", "263 000 000 FCFA", "Somme bloquée sur les comptes de la BOA ; la direction générale demande de réviser les textes de l'EFC (CA du 12/08/2025)", "Trésorerie indisponible ; objet du litige et jugement non transmis"],
                            ["Projet immobilier d'Azaguié (23 ha, 265 lots, 800 appartements)", "19 539 377 111 FCFA", "Plan de masse approuvé par le CA du 12/08/2025 ; prix de vente envisagé de 22 à 30 millions FCFA par logement (CA du 28/05/2026) ; budget 2026 : 2 750 000 000 FCFA d'investissements", "Engagement sans plan de financement transmis ; risque d'usage de l'épargne"],
                            ["Entrée au capital de SOMAVIE", "Non chiffré", "Projet présenté au CA de mars 2026, non approuvé par le conseil", "Risque de participation dans une mutuelle d'assurance vie en difficulté"],
                            ["Clinique (Centre médical du policier)", "187 916 783 FCFA d'apport en 2024", "Chiffre d'affaires 942 760 686 FCFA, charges 1 130 677 469 FCFA (CA du 02/06/2025)", "La clinique consomme des ressources au lieu d'en produire"],
                            ["Crédit court terme BOA (2021)", "500 000 000 FCFA", "Taux de 6,80 % hors taxes (notification du 21/07/2021)", "Coût du financement court terme supérieur au rendement des placements"]]},
        disposition="dessous",
        analyse=["**Le projet d'Azaguié représente plus que l'épargne due** (" + F("AV epargne26", b[2026]["epargne_totale"]) + " au 31/08/2026). Sans plan de financement, il ne peut être financé que par l'épargne ou par l'emprunt."],
        implication="Avant tout engagement, le CA doit disposer d'une étude de faisabilité, d'un plan de financement et d'une analyse de l'effet sur l'adossement de l'épargne.",
        valable="décisions et faits relevés de 2021 au 28/05/2026",
        source="Source : PV des CA du 11/06/2024, 02/06/2025, 12/08/2025, 02/03/2026 et 28/05/2026 ; rapport de l'AG mixte du 17/04/2026 ; notification de crédit BOA du 21/07/2021 ; " + SRC_COMPTA[9:])
'''
rep("# ---- v10 : adossement actif-passif (méthodes justifiées", _ACTIFS + "# ---- v10 : adossement actif-passif (méthodes justifiées")

# ---------------------------------------------------------------- Grenier : états financiers propres
_GREN = r'''# ---- v14 : états financiers du Grenier (consultant, juin 2026)
contenu(blt({"titre": "Grenier du Policier : ses propres états financiers", "sous_titre": "Le Grenier est bénéficiaire, mais ne déprécie aucun prêt et inscrit l'épargne PEP et EFC à son passif",
         "synthese": "Résultat net de 30 543 376 FCFA en 2025 ; encours de 829 623 319 FCFA, identique aux balances âgées ; coût du risque nul"}),
        "États financiers du Grenier établis par un consultant selon le référentiel comptable des systèmes financiers décentralisés (SFD) de l'UMOA, exercices 2022 à 2025.",
        tableau={"titre": tab("Comptes du Grenier du Policier, 2024 et 2025"), "entetes": ["Poste", "2024", "2025", "Lecture"], "largeurs": [8.6, 4.8, 4.8, 13.0], "taille": 8.5,
                 "lignes": [["Encours de crédit", "566 643 134 FCFA", "829 623 319 FCFA", "Identique aux balances âgées du logiciel de prêts"],
                            ["Produits d'intérêts et commissions", "126 424 675 FCFA", "163 822 737 FCFA", "Hausse de 29,58 %"],
                            ["Charges d'exploitation", "116 101 460 FCFA", "133 288 861 FCFA", "Personnel et frais généraux"],
                            ["Provisions pour créances douteuses", "0 FCFA", "0 FCFA", "Aucune dépréciation des prêts en retard"],
                            ["Résultat net", "10 323 215 FCFA", "30 543 376 FCFA", "Bénéficiaire avant toute dépréciation"],
                            ["Épargne PEP et EFC inscrite au passif", "8 493 055 501 FCFA", "8 493 055 501 FCFA", "Montant figé depuis 2023, sans flux de trésorerie"],
                            ["Épargne ordinaire collectée", "101 524 680 FCFA", "182 245 130 FCFA", "Dépôts des membres"],
                            ["Compte de liaison avec le FPPN", "204 060 848 FCFA", "817 121 335 FCFA", "Écart de rapprochement de 9 696 068 FCFA mis en attente"]]},
        disposition="dessous",
        analyse=["**Le résultat du Grenier disparaîtrait avec la dépréciation prévue par sa propre politique de crédit** (" + ref("Prêts du Grenier : dépréciation selon la politique de crédit") + "). En 2024, il a recouvré 47 % des crédits ordinaires octroyés dans l'année et 81 % des crédits flash (rapport du commissaire aux comptes).",
                 "**La même épargne PEP et EFC figure dans les livres du FPPN (comptes 41195100 et 41195200) et dans ceux du Grenier.** Le propriétaire de cette épargne doit être fixé, et la double inscription supprimée à la consolidation."],
        implication="Le Grenier collecte une épargne ordinaire et accorde des crédits : ces activités relèvent de la réglementation des SFD. Le FPPN doit clarifier le statut du Grenier et constituer les dépréciations.",
        valable="exercices 2024 et 2025 du Grenier du Policier, états établis en juin 2026",
        source="Source : « ETAT FINANCIER AU 31 DECEMBRE 2025 » du Grenier du Policier (rapport financier, juin 2026) ; rapport du commissaire aux comptes au CA, exercice 2024 ; " + SRC_GRENIER[9:])
'''
rep('_omis(_snap(), {"titre": "Méthode CIMA des sinistres tardifs"', _GREN + '_omis(_snap(), {"titre": "Méthode CIMA des sinistres tardifs"')

# ---------------------------------------------------------------- demande 3 : commentaire « certains actifs ne sont pas correctement évalués »
rep('''                            ["Réserve de trésorerie des risques longs (norme un trimestre)", N("V3 RL4", RT["RL4"]["valeurs"][2026], 2), "Aucune réserve constituée", ref("Calcul pas à pas : réserves des risques longs 2025")]]},''',
    '''                            ["Réserve de trésorerie des risques longs (norme un trimestre)", N("V3 RL4", RT["RL4"]["valeurs"][2026], 2), "Aucune réserve constituée", ref("Calcul pas à pas : réserves des risques longs 2025")], ["Actifs dont la valeur n'est pas établie (terrain d'Attoban, créance sur l'Hôpital de la Police)", nb(375000000 + 3407519850) + " FCFA", bll("Valeur réalisable non démontrée ; le commissaire aux comptes demande de sortir le terrain de l'actif"), ref("Actifs dont la valeur n'est pas établie")]]},''')
rep('''analyse=["**L'épargne, dette exigible à tout moment, finance des actifs longs : terrains, clinique, créance sur l'hôpital et prêts du Grenier.** Même au taux de sortie moyen observé, l'actif mobilisable dans l'année ne couvre pas les paiements de l'année."],''',
    '''analyse=["**L'épargne, dette exigible à tout moment, finance des actifs longs : terrains, clinique, créance sur l'hôpital et prêts du Grenier.** Même au taux de sortie moyen observé, l'actif mobilisable dans l'année ne couvre pas les paiements de l'année. " + bl("Certains de ces actifs ne sont pas correctement évalués : terrains sans titre foncier, créance sur l'hôpital inchangée depuis 2020.")],''')

# ---------------------------------------------------------------- constat sur les fonds propres
rep('''     "Les pertes reportées dépassent l'ensemble des réserves",''',
    '''     "Les pertes reportées dépassent l'ensemble des réserves. " + bl("Déficit cumulé reconnu par l'AG du 26/06/2024"),''')
rep('''     "Aucun coussin pour absorber un choc de prestations, une perte sur les prêts du Grenier ou une baisse de valeur des actifs",''',
    '''     "Aucun coussin pour absorber un choc de prestations, une perte sur les prêts du Grenier ou une baisse de valeur des actifs. " + bl("Actifs contestés : 3 782 519 850 FCFA de moins"),''')

# ---------------------------------------------------------------- limites
rep('''"Terrains et clinique sans évaluation récente ; créance sur l'Hôpital de la Police sans accord de remboursement", "Valeur réalisable des actifs longs non démontrée",''',
    '''"Terrains et clinique sans évaluation récente ; créance sur l'Hôpital de la Police sans accord de remboursement" + bl(" ; terrain d'Attoban sans titre, à sortir de l'actif selon le commissaire aux comptes"), "Valeur réalisable des actifs longs non démontrée",''')
rep('''"Balance au 31/08/2026 provisoire ; deux versions des comptes 2023", "Situation intermédiaire et excédent 2023 à confirmer",''',
    '''"Balance au 31/08/2026 provisoire" + bl(", refaite le 09/11/2026 (quatre reclassements)") + " ; deux versions des comptes 2023", "Situation intermédiaire et excédent 2023 à confirmer",''')

# ---------------------------------------------------------------- questions au FPPN : réponses partielles et nouvelles questions
rep('''["Provisions", "À quel grade correspond l'allocation de 2 200 000 FCFA des états décès, et quel texte fonde l'exclusion du suicide ?", _n(116), "Moyen"],''',
    '''["Provisions", "À quel grade correspond l'allocation de 2 200 000 FCFA des états décès, et quel texte fonde l'exclusion du suicide ? " + bl("Grade reçu le 09/11/2026 : commissaire de police."), _n(116), "Moyen"],''')
rep('''["Placements", "Quelle est la nature de la créance de 3 407 519 850 FCFA sur l'hôpital de la Police, et un remboursement est-il prévu ?", _n(87), "Moyen"],''',
    '''["Placements", "Quelle est la nature de la créance de 3 407 519 850 FCFA sur l'hôpital de la Police, et un remboursement est-il prévu ? " + bl("Selon le commissaire aux comptes, ce seraient des subventions ; le CA du 06/05/2025 a confié le dossier au PCA. Quelle décision a été prise ?"), _n(87), "Urgent"],''')
rep('''["Gouvernance", "Quelle version des comptes 2023 a été approuvée, et pourquoi le résultat au 31/08/2026 diffère-t-il entre états financiers et balance ?", _n(110), "Urgent"],''',
    '''["Gouvernance", "Quelle version des comptes 2023 a été approuvée, et pourquoi le résultat au 31/08/2026 diffère-t-il entre états financiers et balance ? " + bl("Comptes 2023 approuvés par l'AGO du 26/06/2024 (excédent de 930 629 102 FCFA)."), _n(110), "Urgent"],''')
rep('''    ["Conformité", "Quel est l'effectif total de la Police nationale au 31/12 de 2020 à 2025 et au 31/08/2026 ?", "89", "Moyen"],
]''', '''    ["Conformité", "Quel est l'effectif total de la Police nationale au 31/12 de 2020 à 2025 et au 31/08/2026 ?", "89", "Moyen"],
    ["Épargne", bll("Quel est l'objet du procès perdu sur l'EFC, et pourquoi 263 000 000 FCFA sont-ils bloqués à la BOA ?"), "Nouveau", "Urgent"],
    ["Placements", bll("Quels sont l'étude de faisabilité et le plan de financement du projet d'Azaguié (19 539 377 111 FCFA) ?"), "Nouveau", "Urgent"],
    ["Placements", bll("Quel est le dossier de la créance de 3 932 124 730 FCFA sur l'État, et quel est son lien avec la créance sur l'hôpital ?"), "Nouveau", "Urgent"],
    ["Grenier", bll("Qui est propriétaire de l'épargne PEP et EFC inscrite à la fois dans les livres du FPPN et dans ceux du Grenier ?"), "Nouveau", "Urgent"],
    ["Ratios", bll("Comment le FPPN calcule-t-il 9 mois de réserve obligatoire en 2025 et retraite-t-il les ressources stables ?"), "Nouveau", "Moyen"],
    ["Conformité", bll("Le rapport définitif de la mission de contrôle de l'AIRMS (19 au 22/08/2025) peut-il être transmis ?"), "Nouveau", "Urgent"],
]''')

# ---------------------------------------------------------------- titres renommés (commentaires sur la cohérence des titres) et passe finale
_RENOMS = r'''# ---- v14 : titres renommés (cohérence des titres des ratios, retouche « Dispositif actuariel »)
_REN14 = {"Définition et lecture des ratios (1/3) : solvabilité et réserves": "Solvabilité et réserves : définition des ratios",
          "Tableau de bord (1/4) : solvabilité et réserves": "Solvabilité et réserves : tableau de bord",
          "Définition et lecture des ratios (2/3) : risques longs": "Risques longs : définition des ratios",
          "Tableau de bord (2/4) : risques longs": "Risques longs : tableau de bord",
          "Définition et lecture des ratios (3/3) : gestion": "Gestion : définition des ratios",
          "Tableau de bord (3/4) : gestion": "Gestion : tableau de bord",
          "Tableau de bord (4/4) : données en attente": "Données en attente : tableau de bord",
          "Dispositif actuariel : qui calcule quoi": "Dispositif actuariel"}


def _ren14(o_):
    if isinstance(o_, str):
        for a_, b_ in _REN14.items():
            o_ = o_.replace(a_, b_)
        return o_
    if isinstance(o_, list):
        return [_ren14(v_) for v_ in o_]
    if isinstance(o_, tuple):
        return tuple(_ren14(v_) for v_ in o_)
    if isinstance(o_, dict):
        return {k_: _ren14(v_) for k_, v_ in o_.items()}
    return o_


for _i, _x in enumerate(SL):
    if _x.get("type") == "contenu":
        _t0 = _x["entete"].get("titre")
        SL[_i] = _ren14(_x)
        if _t0 in _REN14 and not _t0.startswith("Dispositif actuariel"):
            BLEU_P.append(_REN14[_t0])
'''
rep("_POS = {}\nfor i, x in enumerate(SL, 1):", _RENOMS + "_POS = {}\nfor i, x in enumerate(SL, 1):")

_FIN14 = r'''
if "BLEU_P" in globals():  # v14 : rapport seulement (les présentations dérivées n'ont pas de registre bleu)
    # ---- v14 : retouches de mise en page de l'utilisateur (rapport du 09/11/2026 retouché le 09/11/2026 au soir)
    from copy import deepcopy as _dc
    _R14 = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "retouches_utilisateur_20261110.json"), encoding="utf-8"))["corrections"]


    def _cle14(sl):
        t_ = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"]
        if t_:
            return {"titre": t_[0]}
        txt_ = [sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
        return {"layout": sl.slide_layout.name, "signature": txt_[:2]}


    _n14 = 0
    for sl in prs.slides:
        k_ = _cle14(sl)
        for c_ in _R14:
            if {kk: c_[kk] for kk in k_ if kk in c_} != k_ or set(c_) - {"formes"} != set(k_):
                continue
            for f_ in c_["formes"]:
                mm_ = [sh for sh in sl.shapes if sh.name == f_["nom"]]
                if f_["rang"] >= len(mm_):
                    continue
                sh = mm_[f_["rang"]]
                g_ = f_["apres"]
                if f_["hauteur_si_texte"] is None:
                    sh.left, sh.top, sh.width, sh.height = (Emu(v_) for v_ in g_)
                else:
                    sh.left, sh.top, sh.width = (Emu(v_) for v_ in g_[:3])
                    if sh.has_text_frame and sh.text_frame.text == f_["hauteur_si_texte"]:
                        sh.height = Emu(g_[3])
                _n14 += 1
    print("v14 : retouches de mise en page appliquées :", _n14)

    # ---- v14 : texte nouveau ou corrigé en bleu
    from lxml import etree as _et14
    _NS14 = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


    def _couleur14(r_el):
        rpr = r_el.find(_NS14 + "rPr")
        if rpr is None:
            rpr = _et14.SubElement(r_el, _NS14 + "rPr")
            r_el.remove(rpr)
            r_el.insert(0, rpr)
        sf = rpr.find(_NS14 + "solidFill")
        if sf is not None:
            c0 = sf.find(_NS14 + "srgbClr")
            if c0 is not None and c0.get("val") in ("FFFFFF",):
                return
            rpr.remove(sf)
        sf = _et14.Element(_NS14 + "solidFill")
        _et14.SubElement(sf, _NS14 + "srgbClr").set("val", BLEU)
        ln = rpr.find(_NS14 + "ln")
        if ln is not None:
            ln.addnext(sf)
        else:
            rpr.insert(0, sf)


    def _paragraphes14(sl):
        for sh in sl.shapes:
            if sh.has_text_frame:
                for p_ in sh.text_frame.paragraphs:
                    yield p_
            if getattr(sh, "has_table", False) and sh.has_table:
                for row_ in sh.table.rows:
                    for c_ in row_.cells:
                        for p_ in c_.text_frame.paragraphs:
                            yield p_
            if sh.shape_type == 6:  # groupe
                for s2 in sh.shapes:
                    if s2.has_text_frame:
                        for p_ in s2.text_frame.paragraphs:
                            yield p_


    _PHR14 = sorted({p_ for p_ in BLEU_P if len(p_) >= 15}, key=len, reverse=True)
    _LIG14 = set(BLEU_L)
    _nb14 = 0
    for sl in prs.slides:
        t_ = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.name == "Adenka_Titre"]
        entiere = bool(t_) and t_[0] in BLEU_T
        if not entiere and _LIG14:
            for sh in sl.shapes:
                if getattr(sh, "has_table", False) and sh.has_table:
                    for row_ in sh.table.rows:
                        if any(c_.text_frame.text in _LIG14 for c_ in row_.cells):
                            for c_ in row_.cells:
                                for p_ in c_.text_frame.paragraphs:
                                    for r_ in p_._p:
                                        if r_.tag == _NS14 + "r":
                                            _couleur14(r_)
        for p_ in _paragraphes14(sl):
            runs_ = [r_ for r_ in p_._p if r_.tag == _NS14 + "r"]
            if not runs_:
                continue
            if entiere:
                if "Empreinte Africaine" in "".join((r_.findtext(_NS14 + "t") or "") for r_ in runs_):
                    continue
                for r_ in runs_:
                    _couleur14(r_)
                _nb14 += 1
                continue
            plein = "".join((r_.findtext(_NS14 + "t") or "") for r_ in runs_)
            marques = [False] * len(plein)
            for ph_ in _PHR14:
                dep_ = plein.find(ph_)
                while dep_ >= 0:
                    for j_ in range(dep_, dep_ + len(ph_)):
                        marques[j_] = True
                    dep_ = plein.find(ph_, dep_ + len(ph_))
            if not any(marques):
                continue
            _nb14 += 1
            pos_ = 0
            for r_ in runs_:
                tx_ = r_.findtext(_NS14 + "t") or ""
                m_ = marques[pos_:pos_ + len(tx_)]
                pos_ += len(tx_)
                if not tx_ or not any(m_):
                    continue
                if all(m_):
                    _couleur14(r_)
                    continue
                # découpage du run en morceaux homogènes
                morceaux, cour, deb = [], m_[0], 0
                for j_ in range(1, len(tx_) + 1):
                    if j_ == len(tx_) or m_[j_] != cour:
                        morceaux.append((tx_[deb:j_], cour))
                        if j_ < len(tx_):
                            cour, deb = m_[j_], j_
                ancre = r_
                for txt_, bleu_ in morceaux:
                    nr_ = _dc(r_)
                    nr_.find(_NS14 + "t").text = txt_
                    if bleu_:
                        _couleur14(nr_)
                    ancre.addnext(nr_)
                    ancre = nr_
                r_.getparent().remove(r_)
    print("v14 : paragraphes passés en bleu :", _nb14, "; diapositives entièrement nouvelles :", len(BLEU_T))
'''
rep("prs.save(SORTIE)\nfinaliser(SORTIE)\nos.remove(tmp)", _FIN14 + "prs.save(SORTIE)\nfinaliser(SORTIE)\nos.remove(tmp)")
rep('CJ[k] = {"Google Shape;138;p25": "Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "09 novembre 2026"}',
    'CJ[k] = {"Google Shape;138;p25": "Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "10 novembre 2026"}')
