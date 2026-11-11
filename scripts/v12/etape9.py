# -*- coding: utf-8 -*-
# v15, étape 9 (11/11/2026) : analyse approfondie de la sinistralité (lecture dans les deux sens), à partir des grands livres
# des comptes 2021 à 2026 (écritures des comptes 601, 6018 et 60800007 classées une à une), des balances, des PV et des
# rapports du commissaire aux comptes. Cinq diapositives nouvelles, écrites en bleu. Exécuté par patch_v12.py après etape8.py.
# Données : assets/sinistralite_v15.json (produit par scripts/sinistralite_gl_fppn.py).

_SIN15 = r'''
# ---- v15 : sinistralité lue dans les deux sens (grands livres 2021 à 2026)
SQ = {int(k_): v_ for k_, v_ in json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "sinistralite_v15.json"), encoding="utf-8")).items()}
_A15 = list(range(2021, 2027))
_LIB15 = ["2021", "2022", "2023", "2024", "2025", "31/08/2026"]
contenu(blt({"titre": "Sinistralité : lire le taux dans les deux sens", "sous_titre": "Pour une mutuelle sociale, un taux bas n'est pas une réussite et un taux qui monte n'est pas forcément un progrès",
         "synthese": "Le bon repère n'est pas le seul taux de sinistralité, mais le ratio combiné : prestations et frais rapportés aux cotisations, qui doit rester sous 100 %"}),
        "Lecture comparée d'un taux de sinistralité selon qu'on raisonne en assureur ou en mutuelle sociale, appliquée aux chiffres du FPPN.",
        tableau={"titre": tab("Lire un taux de sinistralité : assureur, mutuelle sociale et situation du FPPN"), "entetes": ["Niveau du taux", "Pour un assureur", "Pour une mutuelle sociale (norme AIRMS)", "Ce que l'on observe au FPPN"],
                 "largeurs": [4.6, 7.4, 9.0, 10.2], "taille": 8.5,
                 "lignes": [["Sous 70 %", "Très rentable : la prime dépasse largement le coût des sinistres", "Anomalie : les cotisations ne reviennent pas assez aux adhérents en soins. L'AIRMS l'a relevé pour 2022 (9,06 %)", "2020 à 2024 : taux réglementaire de 8,92 % à 40,34 %, excédents de 2 083 908 202 FCFA (2021) et 2 060 332 677 FCFA (2022)"],
                            ["De 70 % à 80 %", "Correct, la marge reste positive", "Cible : il reste 20 % à 30 % pour les frais (au plus 20 %) et les réserves", "31/08/2026 : 75,20 %, dans la norme pour la première fois, budget 2026 calé sur 80 %"],
                            ["Au-dessus de 80 %", "Marge faible", "Pas assez de marge pour les frais et la réserve obligatoire", "Maladie seule : 73,53 % (2025) puis 97,40 % (31/08/2026)"],
                            ["Ratio combiné au-dessus de 100 %", "Perte technique", "La garantie consomme plus qu'elle ne reçoit, frais compris : elle puise dans les réserves", "Maladie : 109,87 % en 2025 et 141,52 % sur les huit mois de 2026, avec des frais de 36,34 % (2025) et 44,12 % (31/08/2026) des cotisations"]]},
        disposition="dessous",
        analyse=["**Un taux qui approche la norme est d'abord une correction :** de 2020 à 2022, le FPPN rendait moins de 10 % des cotisations en prestations, ce que l'AIRMS lui a demandé de corriger (courrier du 29/09/2023).",
                 "**Mais la question utile est de savoir si le taux mesure bien les soins, et si l'équilibre tient frais compris.** Les deux diapositives suivantes répondent : le taux est gonflé par des dépenses étrangères aux soins, et la maladie perd de l'argent même sans elles."],
        implication="Le FPPN doit piloter le ratio combiné par garantie (soins et frais), et non le seul taux réglementaire global, qui masque le déficit de la maladie derrière l'excédent du décès.",
        valable="exercices 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire)",
        source="Source : " + DEC + ", annexe ; courrier de l'AIRMS du 29/09/2023 ; PV des AGO du 16/06/2022 et du 15/06/2023 ; rapport de l'AG mixte du 17/04/2026 (budget 2026) ; " + SRC_COMPTA[9:])
_CAT15 = ["Clinique du FPPN (CMP)", "Autres prestataires de soins", "Dons (autorités, élus, délégués)", "Gadgets, imprimés et cartes", "Honoraires d'un cabinet (factures trimestrielles)", "Régularisations comptables"]
_LIBC15 = {"Clinique du FPPN (CMP)": "Factures de la clinique du FPPN (CMP)", "Autres prestataires de soins": "Hôpitaux, cliniques, pharmacies et autres soins",
           "Dons (autorités, élus, délégués)": "Dons (autorités de la Police, élus, délégués, partenaires, obsèques)", "Gadgets, imprimés et cartes": "Gadgets, imprimés et cartes",
           "Honoraires d'un cabinet (factures trimestrielles)": "[[Honoraires trimestriels d'un cabinet (factures n° 00444…), nature à justifier]]", "Régularisations comptables": "Régularisations comptables (avances anciennes, frais de recouvrement)"}
contenu(blt({"titre": "Prestations maladie : ce que contient le compte", "sous_titre": f"En 2025, {pc('SQ part ns 2025', SQ[2025]['non_soins'] / SQ[2025]['prest'])} des « prestations maladie » ne sont pas des soins",
         "synthese": f"Dons, gadgets et honoraires imputés aux prestations : {F('SQ ns 2025', SQ[2025]['non_soins'])} en 2025 et {F('SQ ns 2026', SQ[2026]['non_soins'])} sur les huit mois de 2026"}),
        "Chaque écriture des comptes de prestations maladie (601, 6018 et 60800007) des grands livres 2021 à 2026 a été classée selon son libellé. Les montants sont nets des annulations.",
        tableau={"titre": tab("Écritures des comptes de prestations maladie par nature réelle (FCFA)"), "entetes": ["Nature réelle"] + _LIB15,
                 "largeurs": [9.4, 3.6, 3.6, 3.6, 3.6, 3.6, 3.8], "taille": 8.5,
                 "lignes": [[_LIBC15[c_]] + [nb(SQ[a_]["cats"].get(c_, 0.0)) for a_ in _A15] for c_ in _CAT15]
                          + [["Total des prestations maladie comptabilisées"] + [nb(SQ[a_]["prest"]) for a_ in _A15],
                             ["Dont dépenses étrangères aux soins"] + [nb(SQ[a_]["non_soins"]) for a_ in _A15]]},
        disposition="dessous",
        analyse=[f"**Les dons aux autorités de la Police, aux élus et aux délégués sont payés sur le compte des soins :** {F('SQ dons 2025', SQ[2025]['cats']['Dons (autorités, élus, délégués)'])} en 2025, avec des libellés comme « DON DU FPPN AUX AUTORITES/TABASKI » ou « … /RAMADAN 26 ».",
                 f"**La clinique du FPPN reçoit {pc('SQ part cmp 2025', SQ[2025]['part_cmp'])} des soins de 2025.** Ses factures mensuelles (« FACT CMP/… ») sont fixées par le FPPN lui-même : le taux mesure en partie un transfert interne. Le cabinet facture chaque trimestre de 20 447 817 à 32 173 310 FCFA par des écritures d'opérations diverses (journal OD), sur le compte des soins."],
        implication="Ces dépenses doivent sortir des prestations : les dons vers les charges de fonctionnement, après décision des organes, et les gadgets et imprimés vers les frais de gestion. La nature des honoraires du cabinet doit être justifiée.",
        valable="écritures des grands livres des comptes 2021 à 2025 et du 01/01 au 31/08/2026, classées par libellé",
        source="Source : « GRAND LIVRE 2021.xlsx », « GRAND LIVRE DES COMPTES 2022.xlsx », « GRAND LIVRE DES COMPTES FPPN 2023.xlsx », « GRAND LIVRE FPPN 2024.xlsx », « GRAND LIVRE DES COMPTES FPPN 2025.xlsx », « GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx » (dataroom du 09/11/2026) ; classement Adenka (onglet 35 du classeur)")
contenu(blt({"titre": "Sinistralité des soins seuls", "sous_titre": "Même sans les dépenses étrangères aux soins, la maladie devient déficitaire en 2026",
         "synthese": f"Taux maladie des soins seuls : {pc('SQ scs 2025', SQ[2025]['sc_soins'])} en 2025 et {pc('SQ scs 2026', SQ[2026]['sc_soins'])} au 31/08/2026. Ratio combiné des soins seuls : {pc('SQ cbs 2026', SQ[2026]['comb_soins'])} en 2026"}),
        "Soins seuls = prestations maladie comptabilisées moins les dons, gadgets, imprimés, honoraires du cabinet et régularisations. Ratio combiné = (prestations + frais affectés à la maladie) / cotisations maladie. Taux réglementaire = prestations maladie et décès / cotisations maladie et décès.",
        tableau={"titre": tab("Sinistralité comptable et sinistralité des soins seuls"), "entetes": ["Indicateur"] + _LIB15, "largeurs": [10.0, 3.5, 3.5, 3.5, 3.5, 3.5, 3.7], "taille": 8.5,
                 "lignes": [["Maladie : taux comptable"] + [nb(100 * SQ[a_]["sc"], 2) + " %" for a_ in _A15],
                            ["Maladie : taux des soins seuls"] + [nb(100 * SQ[a_]["sc_soins"], 2) + " %" for a_ in _A15],
                            ["Part de la clinique du FPPN dans les soins"] + [nb(100 * SQ[a_]["part_cmp"], 2) + " %" for a_ in _A15],
                            ["Maladie : ratio combiné comptable"] + [nb(100 * SQ[a_]["comb"], 2) + " %" for a_ in _A15],
                            ["Maladie : ratio combiné des soins seuls"] + [nb(100 * SQ[a_]["comb_soins"], 2) + " %" for a_ in _A15],
                            ["Toutes garanties : taux réglementaire comptable"] + [nb(100 * SQ[a_]["sin"], 2) + " %" for a_ in _A15],
                            ["Toutes garanties : taux réglementaire des soins seuls"] + [nb(100 * SQ[a_]["sin_soins"], 2) + " %" for a_ in _A15]]},
        disposition="dessous",
        analyse=[f"**Une partie de la hausse est artificielle :** en 2025, le taux réglementaire passe de {pc('SQ sin 2025', SQ[2025]['sin'])} à {pc('SQ sins 2025', SQ[2025]['sin_soins'])} sans les dépenses étrangères aux soins, sous la norme de 70 %. La conformité apparente de 2026 ({pc('SQ sin 2026', SQ[2026]['sin'])}) tombe à {pc('SQ sins 2026', SQ[2026]['sin_soins'])}.",
                 f"**Mais la tendance des soins est réelle :** le taux des soins seuls gagne de 17,08 à 21,65 points par an depuis 2023. Avec des frais de 44,12 % des cotisations, la maladie consomme {pc('SQ cbs 2026 b', SQ[2026]['comb_soins'])} de ses cotisations en 2026, soins seuls. Le décès, à 21,76 %, finance l'écart."],
        implication="La hausse n'est donc ni une bonne nouvelle ni une mauvaise en soi : elle corrige un sous-service ancien, mais elle est gonflée par des dépenses hors soins et elle rend la maladie déficitaire. La cotisation maladie, la tarification de la clinique et les frais doivent être revus ensemble.",
        valable="exercices 2021 à 2025 et du 01/01 au 31/08/2026 (huit mois, balance provisoire)",
        source="Source : grands livres des comptes 2021 au 31/08/2026 (dataroom du 09/11/2026) ; " + SRC_COMPTA[9:] + " ; onglets 28 et 35 du classeur")
contenu(blt({"titre": "Pourquoi le taux était proche de zéro de 2020 à 2022", "sous_titre": "Des raisons réelles, mais aussi des dépenses de santé portées à l'actif au lieu des charges",
         "synthese": "Un taux de 6 % pour une garantie santé est anormal : les documents montrent ce qui l'explique et ce qu'il cachait"}),
        "Chaque explication est rattachée au document qui l'établit. Les montants portés à l'actif sont ceux des comptes 471111 des balances.",
        tableau={"titre": tab("Explications du faible taux de sinistralité de 2020 à 2022"), "entetes": ["Explication", "Effet sur le taux", "Document"], "largeurs": [12.6, 9.4, 9.2], "taille": 8,
                 "lignes": [["Garantie limitée aux soins liés à une hospitalisation", "Peu de soins remboursés : taux bas réel", "Règlement intérieur, art. 41 et 45"],
                            ["Produits ajoutés en juin 2023 : santé des policiers retraités, MEDOC (30 % du panier MUGEF-CI), frais funéraires", "Hausse des prestations à partir de 2023", "PV de l'AGO du 15/06/2023"],
                            ["Clinique du FPPN mise en service en 2022 et 2023", "Accès aux soins élargi : hausse des prestations", "Comptes 47111104 et 47111105 ; PV du CA du 02/06/2025"],
                            ["Dépenses de la clinique payées par le FPPN et inscrites à l'actif : 554 617 537 FCFA débités au compte « Polyclinique du FPPN » en 2021, 511 767 833 FCFA aux « charges de fonctionnement au CMP » en 2022", "[[Taux sous-estimé : en 2022, ces charges nettes des remboursements (469 980 456 FCFA) auraient porté le taux maladie de 6,47 % à 32,86 %]]", "Comptes 47111101 et 47111105 des balances 2021 et 2022"],
                            ["Créance sur l'Hôpital de la Police : 3 407 519 850 FCFA de dépenses depuis 2004, inscrites à l'actif", "Dépenses de santé jamais passées en charges (subventions possibles selon la direction générale)", "Rapport du commissaire aux comptes au CA, exercice 2024"],
                            ["Créance sur l'État pour les soins des policiers blessés en service : 3 932 124 730 FCFA", "[[Soins financés par le FPPN, non passés en prestations]]", "Courrier de l'AIRMS du 05/01/2026"],
                            ["Provisions maladie et obsèques figées à 4 829 298 290 FCFA depuis 2020", "Aucune variation de provision dans la charge de l'année", "Balances 2020 au 31/08/2026"]]},
        disposition="dessous",
        analyse=["**Le taux proche de zéro ne mesurait pas le coût réel des soins :** une partie de la santé des policiers était financée par le FPPN mais inscrite à l'actif (clinique, hôpital, État), ce qui gonflait les excédents et l'actif.",
                 "**Ces excédents ont financé des actifs peu liquides** (clinique, terrains) plutôt que des prestations ou la réserve obligatoire, que l'AIRMS demandait de constituer (courrier du 29/09/2023)."],
        implication="Le taux bas n'était donc pas qu'un « bon résultat » : il reflétait une couverture étroite et un classement comptable qui sortait des dépenses de santé du compte de résultat. Ces montants restent à reclasser ou à déprécier.",
        valable="exercices 2020 à 2022, éclairés par les documents reçus jusqu'au 09/11/2026",
        source="Source : règlement intérieur du FPPN ; PV de l'AGO du 15/06/2023 ; rapport du commissaire aux comptes au CA (exercice 2024) ; courriers de l'AIRMS du 29/09/2023 et du 05/01/2026 ; " + SRC_COMPTA[9:])
contenu(blt({"titre": "Sinistralité : ce qu'il faut en retenir", "sous_titre": "La hausse corrige un sous-service ancien, mais elle est en partie artificielle et rend la maladie déficitaire",
         "synthese": "Ni le taux bas d'hier ni la conformité apparente d'aujourd'hui ne doivent être lus comme de bonnes nouvelles"}),
        "Synthèse de l'analyse de la sinistralité dans les deux sens.",
        tableau={"titre": tab("Sinistralité : ce qui rassure, ce qui inquiète, ce qu'il faut faire"), "entetes": ["Ce qui rassure", "Ce qui inquiète", "Ce qu'il faut faire"], "largeurs": [10.4, 10.4, 10.4], "taille": 8.5,
                 "lignes": [["Les cotisations reviennent davantage aux adhérents : 75,20 % au 31/08/2026, dans la norme", "Une partie du taux n'est pas du soin : dons, gadgets, honoraires (" + nb(SQ[2025]["non_soins"]) + " FCFA en 2025)", "Sortir des prestations les dépenses hors soins et justifier les honoraires du cabinet"],
                            ["L'accès aux soins s'élargit (clinique, retraités, MEDOC)", "La clinique du FPPN reçoit environ 7 FCFA de soins sur 10 et fixe elle-même ses factures", "Tenir une comptabilité séparée de la clinique et facturer au coût réel, contrôlé"],
                            ["Le décès reste très excédentaire (21,76 % au 31/08/2026)", "La maladie est déficitaire même en soins seuls (ratio combiné de 119,07 % en 2026)", "Recalculer la cotisation maladie (3,5 % du salaire de base) et ramener les frais à 20 % au plus"],
                            ["Le budget 2026 vise une sinistralité de 80 %", "Le taux est piloté par le budget, pas par une mesure du risque", "Calculer les provisions dossier par dossier et suivre le ratio combiné chaque trimestre"]]},
        disposition="dessous",
        analyse=["**Message clé : un taux de sinistralité ne vaut que par ce qu'il contient.** Au FPPN, il a d'abord été trop bas parce que des dépenses de santé étaient inscrites à l'actif, puis il a été gonflé par des dépenses étrangères aux soins. Le vrai signal est le ratio combiné de la maladie, qui dépasse 100 % en 2026."],
        implication="Ces constats complètent le constat STAT-2 et les recommandations de l'axe « Conformité et gouvernance ».",
        valable="exercices 2020 à 2025 et du 01/01 au 31/08/2026 (balance provisoire)",
        source="Source : diapositives précédentes ; grands livres 2021 au 31/08/2026 ; " + SRC_COMPTA[9:])
'''
rep('contenu({"titre": "Frais de gestion et résultat", "sous_titre": "Les excédents se réduisent et deviennent un déficit en 2026",',
    _SIN15 + 'contenu({"titre": "Frais de gestion et résultat", "sous_titre": "Les excédents se réduisent et deviennent un déficit en 2026",')
rep('''f"La hausse du S/C maladie est statistiquement significative, d'environ {N('Pente S/C (ret)', C['tendances'][1]['pente'], 2)} points par an de 2020 à 2025. Les frais dépassent la norme de 20 % des cotisations.",''',
    '''f"La hausse du S/C maladie est statistiquement significative, d'environ {N('Pente S/C (ret)', C['tendances'][1]['pente'], 2)} points par an de 2020 à 2025. Les frais dépassent la norme de 20 % des cotisations. " + bl("Sans les dons, gadgets et honoraires imputés aux soins, le taux maladie est de 53,30 % en 2025 et de 74,95 % au 31/08/2026."),''')
rep('''"**Le compte global empêche le suivi par type de soins et par prestataire.**"''',
    '''"**Le compte global empêche le suivi par type de soins et par prestataire.** " + bl("La lecture des grands livres montre qu'il contient aussi des dons, des gadgets et des honoraires (diapositive suivante).")''')
