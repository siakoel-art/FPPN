# -*- coding: utf-8 -*-
# v16, étape 10 (11/11/2026) : documents de la dataroom du 11/11/2026 (32 fichiers nouveaux), écrits en bleu.
#  - mémo du service comptable sur les comptes 47 et 27 ;
#  - rapport d'expertise du siège (19/03/2019) rapproché du compte 10620000 ;
#  - balances et grands livres du Grenier 2021 à 2026 (compte de liaison 37921400, fonds de dotation, épargne) ;
#  - déclarations mensuelles des salaires (EDI ITS) : personnel sans numéro CNPS, hors provision retraite ;
#  - état des comptes bancaires 2020-2026 : origine du fonds commun de placement ;
#  - conventions de production des cartes, PV d'arrêté de caisse 2024, inventaire physique du 31/08/2026.
# Exécuté par patch_v12.py après etape9.py.

_DOC16 = r'''
# ---- v16 : documents reçus le 11/11/2026
contenu(blt({"titre": "Documents reçus le 11/11/2026", "sous_titre": "32 fichiers nouveaux, lus en entier : quatre apportent des faits nouveaux pour le volet actuariel",
         "synthese": "Le mémo sur les comptes 47 et 27, l'expertise du siège, les comptes du Grenier et les déclarations de salaires changent la lecture de l'actif et des provisions"}),
        "Inventaire du 11/11/2026 : 263 fichiers, dont 199 déjà reçus, 30 pièces d'achats sans objet pour le volet actuariel, 2 doublons et 32 nouveaux.",
        tableau={"titre": tab("Documents reçus le 11/11/2026 et usage dans le rapport"), "entetes": ["Document", "Ce qu'il apporte", "Diapositive"], "largeurs": [10.2, 15.0, 6.0], "taille": 8,
                 "lignes": [["Mémo sur les comptes 47 et 27 (service comptable)", "Nature de chaque créance : clinique, Hôpital de la Police, Grenier, prêts aux adhérents de 2004 à 2015", ref("Comptes 47 et 27 : ce qu'en dit le service comptable")],
                            ["Rapport d'expertise du siège (19/03/2019)", "Valeur vénale de 2 899 400 000 FCFA, égale au compte 10620000", ref("Réévaluation du siège : le compte 10620000")],
                            ["Balances et grands livres du Grenier, 2021 au 31/08/2026", "Compte de liaison avec le FPPN, fonds de dotation, épargne inscrite", ref("Grenier du Policier : liaison avec le FPPN")],
                            ["Déclaration des salaires d'août 2026 (EDI ITS)", "100 salariés, dont 33 sans numéro CNPS", ref("Personnel : salariés hors provision retraite")],
                            ["État des comptes bancaires, 08/2020 au 31/08/2026", "Origine du fonds commun de placement (épargne PEP)", ref("Calcul pas à pas : couverture de l'épargne")],
                            ["Conventions cartes (Tekmerion, Microcard System)", "Carte payée 4 000 FCFA, dont 400 FCFA puis 1 000 FCFA reversés au FPPN", "Sans effet sur les chiffres"],
                            ["PV d'arrêté de caisse au 31/12/2024", "Caisse : 5 000 850 FCFA comptés pour 5 000 845 FCFA en comptabilité (écart de 5 FCFA)", "Sans effet sur les chiffres"],
                            ["Inventaire physique au 31/08/2026", "586 lignes, 33 biens non vus, 17 hors service, sans valeur", "Limites"],
                            ["Fichiers de provision du personnel 2020 à 2025", "Nouvelle copie, valeurs identiques à celles du 09/11/2026", "Sans effet sur les chiffres"]]},
        disposition="dessous",
        analyse=["**Les fichiers nouveaux confirment les constats sur l'actif et en ajoutent deux** : la réévaluation du siège porte la valeur entière de l'expertise, et une partie du personnel n'entre pas dans la provision retraite."],
        implication="Les constats ACT-1 et GRE-1 et la provision du personnel sont complétés en conséquence.",
        valable="documents de la dataroom au 11/11/2026",
        source="Source : inventaire de la dataroom du 11/11/2026 (FPPN_inventaire_20261111.csv) ; documents listés")
'''

_ACT16 = r'''
# ---- v16 : réévaluation du siège et comptes 47 et 27
contenu(blt({"titre": "Réévaluation du siège : le compte 10620000", "sous_titre": "La réserve de réévaluation est égale à la valeur entière du siège estimée en 2019, et non à une plus-value",
         "synthese": "2 899 400 000 FCFA au compte 10620000 de 2020 au 31/08/2026, soit au franc près la valeur vénale conclue par l'expert le 19/03/2019"}),
        "Rapprochement du rapport d'expertise n° 005/AKKA/19 et des balances 2020 au 31/08/2026.",
        tableau={"titre": tab("Expertise du siège et comptes du FPPN"), "entetes": ["Élément", "Montant ou fait", "Source"], "largeurs": [10.4, 10.8, 10.0], "taille": 8,
                 "lignes": [["Objet de l'expertise", "Garantie d'une demande de ligne de crédit", "Rapport du 19/03/2019, p. 1"],
                            ["Bien", "Terrain de 4 142 m², Marcory zone 4/C, lot 34/A, titre foncier n° 114 462", "Certificat de propriété annexé"],
                            ["Méthode retenue", "Site évalué comme terrain nu : 4 142 m² x 700 000 FCFA, bâtiments non comptés", "Rapport, p. 10"],
                            ["Valeur vénale conclue", "2 899 400 000 FCFA", "Rapport, conclusion"],
                            ["Compte 10620000 (« Réévaluation matériels », puis « Réserve disponible » en 2026)", "2 899 400 000 FCFA, sans variation de 2020 au 31/08/2026", "Balances"],
                            ["Bâtiments administratifs (23130000), nets des amortissements", "424 341 213 FCFA - 273 868 454 FCFA = 150 472 759 FCFA, gardés en plus du terrain", "Balance au 31/08/2026"],
                            ["Terrains nus (22200000)", "3 931 156 624 FCFA en 2020, 5 245 896 750 FCFA au 31/08/2026", "Balances"]]},
        disposition="dessous",
        analyse=["**Une réserve de réévaluation doit égaler l'écart entre la valeur réévaluée et la valeur comptable.** Ici, elle égale la valeur entière : [[si le coût d'origine du terrain du siège figure déjà au compte 22200000, l'actif et les fonds propres comptent ce coût deux fois]].",
                 "**La valeur date de 2019, a été établie pour obtenir un crédit et traite les bâtiments comme sans valeur,** alors que les comptes les gardent aussi à l'actif."],
        implication="Les fonds propres affichés reposent pour 2 899 400 000 FCFA sur une estimation ancienne, non suivie d'une décision de réévaluation connue. Le FPPN doit produire l'écriture d'origine et le coût d'acquisition du terrain.",
        valable="balances 2020 au 31/08/2026 ; rapport d'expertise du 19/03/2019",
        source="Source : rapport d'expertise de l'ensemble immobilier du siège du FPPN, n° 005/AKKA/19 du 19/03/2019 (dataroom du 11/11/2026) ; " + SRC_COMPTA[9:])
contenu(blt({"titre": "Comptes 47 et 27 : ce qu'en dit le service comptable", "sous_titre": "Des créances financent la clinique, l'hôpital et le Grenier, et des prêts anciens ne sont plus suivis",
         "synthese": "Créances sur la clinique, l'hôpital et le Grenier : 11 406 323 425 FCFA au 31/08/2026. Prêts aux adhérents de 2004 à 2015 non suivis : 379 204 261 FCFA"}),
        "Mémo du service comptable sur les comptes 47 et 27, rapproché de la balance au 31/08/2026. Les explications sont celles du mémo.",
        tableau={"titre": tab("Comptes 47 et 27 au 31/08/2026 : solde et explication du service comptable"), "entetes": ["Compte", "Solde (FCFA)", "Explication du mémo", "Lecture Adenka"], "largeurs": [6.0, 3.6, 11.0, 10.6], "taille": 7.5,
                 "lignes": [["47111104 Investissements au CMP", "5 803 025 455", "Investissements faits par le FPPN pour la clinique", "Actif immobilisé de fait, sans remboursement prévu"],
                            ["47111100 Hôpital de la Police", "3 407 519 850", "Charges et investissements faits pour l'hôpital « avant leur différend »", "Recouvrement non établi (ACT-1)"],
                            ["47111105 Charges de fonctionnement au CMP", "1 280 300 860", "Charges de la clinique payées par le FPPN, « déduites des prestations offertes »", "Dépenses de santé portées à l'actif, hors prestations"],
                            ["47111102 Grenier du Policier", "915 477 260", "Fonds mis à disposition, charges et investissements du Grenier", "Le Grenier n'en reconnaît que 845 285 760 FCFA"],
                            ["27110000 et 27120000 Prêts PEP", "379 204 261", "Prêts aux adhérents de 2004 à 2015, non remboursés, suivi perdu après l'incendie du siège en 2010", "Créance à déprécier sauf justification"],
                            ["47111103 Ticket modérateur", "8 357 475", "Suivi des 20 % de ticket modérateur au CMP abandonné, « à régulariser »", "Créance sans contrepartie identifiée"],
                            ["27200000 Prêts au personnel", "2 500 000", "Prêts au personnel civil et policier", "Le mémo indique 2 400 000 FCFA"]]},
        disposition="dessous",
        analyse=["**Le mémo confirme la lecture de la sinistralité :** une partie des soins de la clinique est payée par le FPPN et inscrite à l'actif, nette des prestations facturées, au lieu d'être en charges."],
        implication="Ces créances ne sont pas des placements : elles doivent être justifiées ou dépréciées avant de compter dans la couverture de l'épargne ou dans les fonds propres.",
        valable="balance au 31/08/2026 (provisoire) ; mémo du service comptable reçu le 11/11/2026",
        source="Source : « MEMO SUR LES COMPTES 47 ET 27.xlsx » (dataroom du 11/11/2026) ; " + SRC_COMPTA[9:])
'''

_GRE16 = r'''
# ---- v16 : Grenier, liaison avec le FPPN
contenu(blt({"titre": "Grenier du Policier : liaison avec le FPPN", "sous_titre": "Les deux comptabilités ne concordent plus, et l'épargne inscrite au Grenier n'a pas d'actif réel en face",
         "synthese": "Écart de 70 191 500 FCFA entre les deux comptes de liaison au 31/08/2026. Épargne figée à 8 493 055 501 FCFA au Grenier, contre 13 814 066 964 FCFA au FPPN fin 2025"}),
        "Balances et grands livres du Grenier 2021 au 31/08/2026 (« Cabinet Adenka »), rapprochés des balances et grands livres du FPPN.",
        tableau={"titre": tab("Grenier et FPPN : comptes réciproques"), "entetes": ["Point", "Grenier", "FPPN", "Lecture"], "largeurs": [6.2, 7.6, 7.6, 9.8], "taille": 7.5,
                 "lignes": [["Compte de liaison au 31/08/2026", "37921400 : 845 285 760 FCFA dus au FPPN", "47111102 : 915 477 260 FCFA dus par le Grenier", "Écart de 70 191 500 FCFA, égaux au 31/12/2025"],
                            ["Fonds de dotation", "500 000 000 FCFA reçus le 18/05/2022, reclassés en dette au 31/12/2025", "Compté dans la créance sur le Grenier", "Le Grenier n'a plus de dotation propre"],
                            ["Épargne PEP et EFC", "8 493 055 501 FCFA, sans mouvement depuis 2023", "13 814 066 964 FCFA au 31/12/2025", "Écart de 5 321 011 463 FCFA"],
                            ["Contrepartie de cette épargne", "37911400 « fonds en attente de réception »", "Aucun transfert de fonds", "Actif sans encaissement réel"],
                            ["Crédits aux membres au 31/08/2026", "1 044 302 107 FCFA, aucune dépréciation", "—", "Risque de crédit non mesuré (PROV-3)"],
                            ["Gadgets et dépliants du Grenier", "Refacturés par le FPPN via le compte de liaison", "4 factures « GRENIER » passées en prestations maladie : 8 775 000 FCFA (2024), 12 210 000 FCFA (2025)", "Charges du Grenier dans les soins du FPPN"]]},
        disposition="dessous",
        analyse=["**Les comptes du Grenier 2025 ne sont pas clôturés :** la balance au 31/08/2026 reprend les charges et produits de 2025. Résultat 2025 : 30 543 376 FCFA, conforme aux états financiers.",
                 "**Les salaires du personnel du Grenier sont payés par le FPPN et refacturés en fin d'année :** le résultat de huit mois de 2026 (116 597 438 FCFA) ne les contient pas."],
        implication="Le FPPN et le Grenier doivent rapprocher leur compte de liaison chaque mois et supprimer l'épargne inscrite au Grenier, qui n'y a pas d'actif.",
        valable="balances et grands livres du Grenier du 31/12/2021 au 31/08/2026 ; balances et grands livres du FPPN",
        source="Source : « BALANCE GENERALE » 2021 à 2026, « Grd livre inter grenier-FPPN » 2024 à 2026 (compte 37921400), « Fonds de dotation 2022.xlsx » du Grenier (dataroom du 11/11/2026) ; grands livres du FPPN 2024 au 31/08/2026 ; " + SRC_COMPTA[9:])
'''

_PERS16 = r'''
# ---- v16 : personnel sans numéro CNPS
contenu(blt({"titre": "Personnel : salariés hors provision retraite", "sous_titre": "Un tiers des salariés d'août 2026 sont déclarés sans numéro CNPS et ne figurent pas dans le calcul de la provision",
         "synthese": "33 salariés sur 100 portent le numéro fictif 111111111111 en août 2026, contre 4 en 2024. Ce sont surtout des agents de la clinique"}),
        "Déclarations mensuelles des salaires (EDI ITS, état 301) de septembre 2024 à septembre 2026, rapprochées par numéro CNPS du fichier de provision retraite 2025 du FPPN.",
        tableau={"titre": tab("Salariés déclarés et salariés sans numéro CNPS"), "entetes": ["Mois", "Salariés", "Sans numéro CNPS", "Brut du mois (FCFA)"], "largeurs": [7.0, 6.0, 7.0, 7.0], "taille": 8.5,
                 "lignes": [["Septembre 2024", "61", "4", "17 326 865"], ["Janvier 2025", "58", "2", "16 332 109"], ["Août 2025", "79", "8", "21 866 911"],
                            ["Décembre 2025", "79", "8", "45 545 641"], ["Janvier 2026", "79", "10", "24 649 368"], ["Juin 2026", "83", "16", "25 133 146"],
                            ["Août 2026", "100", "33", "28 171 544"], ["Septembre 2026", "98", "30", "27 940 045"]]},
        disposition="dessous",
        analyse=["**Août 2026 : 67 salariés sont dans le fichier de provision 2025, 33 n'ont pas de numéro CNPS** (18 aides-soignants, 4 auxiliaires en pharmacie, 2 brancardiers, 2 archivistes, agents d'accueil, infirmier, laborantin, ressources humaines), pour 6 340 264 FCFA de salaires bruts par mois.",
                 "**La provision retraite (compte 196) ne couvre que les 78 agents du fichier 2025.** Les salariés sans numéro CNPS n'y sont pas, et leurs cotisations sociales ne sont pas identifiables."],
        implication="[[La provision du personnel est sous-estimée d'un montant qui ne peut pas être calculé sans les dates de naissance et d'embauche de ces 33 salariés.]] Leur situation sociale relève du volet juridique et social.",
        valable="déclarations de salaires de septembre 2024 à septembre 2026 ; fichier de provision retraite 2025",
        source="Source : « edi-annexe-ITS » 2024 à 2026 (dataroom des 09/11 et 11/11/2026) ; « FICHIER PERSONNEL CIVIL FPPN POUR PROVISION CONGES ET RETRAITE EN 2025.xlsx »")
'''

rep('contenu(blt({"titre": "Documents reçus le 09/11/2026"', _DOC16 + 'contenu(blt({"titre": "Documents reçus le 09/11/2026"')
rep('contenu(blt({"titre": "Actifs dont la valeur n\'est pas établie"', _ACT16 + 'contenu(blt({"titre": "Actifs dont la valeur n\'est pas établie"')
rep('contenu(blt({"titre": "Grenier du Policier : ses propres états financiers"', _GRE16 + 'contenu(blt({"titre": "Grenier du Policier : ses propres états financiers"')
rep('contenu(blt({"titre": "Provision du personnel : méthode du FPPN et compte 196"', _PERS16 + 'contenu(blt({"titre": "Provision du personnel : méthode du FPPN et compte 196"')
rep('''("Limite", "Le manuel des produits d'épargne ne fixe pas d'échéancier individuel. Le FCP est compté comme actif réservé, aucun document ne l'affectant à un autre emploi.")''',
    '''("Limite", "Le manuel des produits d'épargne ne fixe pas d'échéancier individuel. Le FCP est compté comme actif réservé, aucun document ne l'affectant à un autre emploi. " + bl("L'état des comptes bancaires reçu le 11/11/2026 le confirme : le dépôt à terme « BOA PEP » (505 383 858 FCFA) est fermé le 30/10/2025 et le FCP de 500 000 000 FCFA ouvert le 04/11/2025 depuis la banque PEP. L'épargne PEP passe ainsi d'un dépôt garanti à un placement dont la valeur peut baisser."))''')
# sinistralité : la nature du compte 47111105 est confirmée par le mémo du service comptable (le chiffre n'est plus à confirmer)
rep('"[[Taux sous-estimé : en 2022, ces charges nettes des remboursements (469 980 456 FCFA) auraient porté le taux maladie de 6,47 % à 32,86 %]]"',
    '"Taux sous-estimé : en 2022, ces charges nettes des remboursements (469 980 456 FCFA) auraient porté le taux maladie de 6,47 % à 32,86 %. " + bl("Le mémo du service comptable confirme qu\'elles sont « déduites des prestations offertes ».")')
# date de couverture du rapport
rep('CJ[k] = {"Google Shape;138;p25": "Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "10 novembre 2026"}',
    'CJ[k] = {"Google Shape;138;p25": "Audit comptable et financier de la gestion du FPPN", "Text Placeholder 4": "11 novembre 2026"}')
