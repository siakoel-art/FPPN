# -*- coding: utf-8 -*-
"""
FPPN - Registre de vérification des constats de la revue indépendante (13/11/2026)
Chaque ligne regroupe les constats en doublon des réviseurs A à E, le résultat de la vérification sur les données
(verif_revue_fppn.py, test_interets_epargne_fppn.py, regle_interets_pep_fppn.py) et la décision proposée.
Utilisation : python registre_revue_fppn.py <sortie.xlsx>
"""
import sys
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SORTIE = sys.argv[1]
C, A, J, D, S = "À corriger", "À ajouter", "À demander au FPPN", "Confirmé, sans changement", "Non revérifié"
CF, PA, NV = "Confirmé", "Confirmé en partie", "Non revérifié"

# (thème, références, diapositives, gravité la plus haute, constat, vérification et chiffre, résultat, décision)
L = [
 ("Adossement et provisions", "A-B1, D-B1", "4, 7, 10, 24, 42, 162, 163, 172, 173, 175", "Bloquant",
  "Le passif à un an contient toutes les provisions maladie et obsèques, jugées excessives par le rapport. « Borne favorable » est faux.",
  "Balance 31/08/2026 : 2 188 438 207 (dettes hors épargne) + 4 829 298 290 (15801, 15802) = 7 017 736 497 FCFA. Couverture 34,42 % devient 66,04 % avec le besoin estimé, 64,46 % avec un mois de prestations maladie (147 636 280 FCFA). Manque à 5 ans -7 786 335 924 FCFA. Au-delà de 5 ans +3 973 884 876 FCFA. Hors épargne : 156,60 %. « Borne favorable » présent diapositives 42, 51, 162, 163, 292.",
  CF, C),
 ("Adossement et provisions", "D-B2, D-m4", "163, 165", "Bloquant",
  "« Même en lecture comptable, l'actif ne couvre pas tous les engagements (96,33 %) » s'inverse avec les provisions au besoin.",
  "22 761 252 888 - 4 809 716 898 + 3 278 913 238 : couverture 103,27 % (122,14 % sans intérêts). L'écart de 43 666 678 FCFA entre -835 832 022 et -879 498 700 FCFA est le compte 196.",
  CF, C),
 ("Adossement et provisions", "A-B3, D-B3, A-M15, A-m10", "9, 10, 16, 18, 168, 177, 178, 204", "Bloquant",
  "Fonds propres « négatifs » et « sous-évalués de 4,8 milliards » sans passage chiffré. Seules les corrections négatives sont retenues.",
  "Pont recalculé : -879 498 700 + 4 809 716 898 - 3 278 913 238 - 14 016 392 = +637 288 568 FCFA (+663 770 796 FCFA avec l'excédent du compte 196). Le signe dépend de l'ordre des corrections.",
  CF, A),
 ("Adossement et provisions", "A-M6", "6, 10, 24", "Majeur",
  "Deux excédents de provisions dans le résumé (4 327 397 229 et 4 809 716 898 FCFA).",
  "4 327 397 229 FCFA en diapositive 6, 4 809 716 898 FCFA en diapositives 10 et 24.", CF, C),
 ("Adossement : scénarios et écoulement", "D-M1", "164, 165, 166", "Majeur",
  "Le scénario 3 retire le terrain du siège alors que son expertise (2 899 400 000 FCFA) a été reçue.",
  "Arithmétique du réviseur reprise : 39,93 % au lieu de 27,19 % au 31/08/2026.", CF, C),
 ("Adossement : scénarios et écoulement", "D-M2", "164, 165, 172", "Majeur",
  "Les scénarios gardent des actifs que leur propre critère devrait retirer.",
  "Balance 31/08/2026 : prêts 27110000 + 27120000 = 379 204 261, 48500000 = 154 758 000, 47111103 = 8 357 475, 22900000 = 574 762 164, 23100000 = 242 441 705 FCFA. Scénario 2 corrigé : 69,42 % au lieu de 75,73 %.",
  CF, C),
 ("Adossement : scénarios et écoulement", "D-M3", "171, 172, 173", "Majeur",
  "Un taux de sortie unique (19,24 %) alors que le PEP sort bien plus vite que l'EFC.",
  "Onglet 24a : PEP 30,30 %, EFC 8,29 % en moyenne. Sorties de l'année 1 : 2 456 094 737 FCFA (15,60 % du stock). PEP 2026 : 1 347 386 639 FCFA sortis en 8 mois, soit 27,14 % du stock d'ouverture.",
  CF, C),
 ("Adossement : scénarios et écoulement", "D-M4", "172, 173, 175", "Majeur",
  "L'écoulement suppose l'arrêt des versements. Le choc sévère mélange stock et flux.",
  "Balance 2025 : versements 4 387 823 635 FCFA contre 1 358 821 103 FCFA de sorties sur les comptes 41195100 et 41195200.",
  CF, C),
 ("Adossement : scénarios et écoulement", "C-M5", "103, 105", "Majeur",
  "Seuil de 21,34 % de retraits calculé sans déduire les dettes exigibles.",
  "Arithmétique confirmée : (3 359 100 618 - 690 234 451 - 489 419 669) / 15 743 516 391 = 13,84 %.", CF, C),
 ("Adossement : scénarios et écoulement", "D-M11, A-V4", "168", "Majeur",
  "Créance sur l'État de 3 932 124 730 FCFA dans les actifs du bilan. SONEB non retiré de la sensibilité.",
  "Aucun compte de la balance au 31/08/2026 ne porte 3 932 124 730 FCFA.", CF, C),
 ("Adossement : scénarios et écoulement", "D-V4", "173", "À vérifier",
  "Choc maladie calibré sur +103,53 % (2024), hausse liée à l'extension de la couverture.",
  "Hausse maladie 2026 annualisée : +39,37 %. La nature du saut de 2024 reste à documenter.", NV, J),
 ("Adossement : scénarios et écoulement", "A-V2, D-V1", "166, 189, 218", "À vérifier",
  "Terrain du siège peut-être compté deux fois (10620000 et 22200000).",
  "10620000 = 2 899 400 000 FCFA depuis 2020. Écriture d'origine non reçue.", NV, J),
 ("Adossement : scénarios et écoulement", "D-m10, D-m11", "161, 175", "Mineur",
  "Renvoi mal écrit en 175. Compte 31200000 débiteur non commenté.",
  "31200000 : 50 000 FCFA au débit en 2026, nul en 2025.", CF, C),
 ("Intérêts de l'épargne", "A-M4, C-B1", "5, 6, 10, 24, 35, 109, 127, 156 à 159", "Bloquant",
  "Les intérêts dus (3 278 913 238 FCFA) sont un scénario présenté comme un recalcul, avec l'EFC à 5,5 % dès 2020 et 2020 compté sur 12 mois.",
  "Recalcul sur les balances : 3 278 913 238 (manuel), 2 954 084 955 (EFC à 3,5 % avant 2024), 2 541 799 542 (3,5 % partout), 2 385 635 461 FCFA (EFC à 0 % avant 2024). 2020 sur 4 mois : 2 974 997 849 FCFA. Capitalisation annuelle non recalculée.",
  CF, C),
 ("Intérêts de l'épargne", "A-V1, C-V2 et test du 13/11/2026", "5, 10, 35, 107, 110", "À vérifier",
  "Les sorties d'épargne contiennent-elles des intérêts payés hors comptabilité ?",
  "PEP : 4 987 remboursements sur 6 482 (09/2020 à 08/2026) sont égaux au versé × 1,035, soit 198 040 675 FCFA payés sur le compte 41195100. Le manuel promet 3,5 % « tous les ans ». Sur 1 190 dossiers complets : 19 715 675 FCFA payés contre 43 468 398 FCFA promis. EFC : règlements en montants ronds, test non concluant.",
  PA, A),
 ("Intérêts de l'épargne", "C-M8", "108, 109", "Majeur",
  "Deux bases pour les intérêts (fin d'année et moyenne). « 16,1 fois » compare un stock à des flux.",
  "EFC 2025 : 486 686 477 FCFA sur l'encours de fin, 434 598 607 FCFA sur l'encours moyen. 2 804 987 325 / 173 881 316 = 16,13.", CF, C),
 ("Intérêts de l'épargne", "A-M3", "9, 10, 15, 39", "Majeur",
  "« Pas d'exploitation » alors que les excédents ignorent les charges de santé à l'actif et les intérêts.",
  "Phrase en diapositive 9. 437 899 911 - 588 520 656 = -150 620 745 FCFA en 2025 au taux du manuel.", CF, C),
 ("Ratio de 16,28 %", "A-M8, C-B2", "9, 10, 13, 17, 19, 21, 100, 105", "Bloquant",
  "16,28 % présenté comme le ratio de solvabilité des risques longs de la Décision n° 05/2022.",
  "Décision n° 05/2022, § 2.1 : « actifs admis rapporté aux engagements réglementés ». 2 562 885 661 / 15 743 516 391 = 16,28 % est un ratio interne.", CF, C),
 ("Ratio de 16,28 %", "C-V1, D-V2", "100, 101, 164, 165, 173, 179", "À vérifier",
  "FCP compté comme réservé à l'épargne et à sa valeur d'achat.",
  "53300001 = 500 000 000 FCFA. Sans le FCP : 13,10 %. Décision d'affectation et valeur liquidative non reçues.", NV, J),
 ("Ratio de 16,28 %", "A-M9", "13, 17", "Majeur",
  "« L'épargne grossit plus vite que les actifs qui la garantissent. »",
  "Phrase en diapositive 17. Couverture 6,49 %, 11,64 %, 16,28 % de 2024 à 2026 : elle remonte.", CF, C),
 ("Sinistralité 2026", "B-B1, A-M2, B-m4", "13, 19, 21, 82, 84, 85, 91, 93, 95, 98", "Bloquant",
  "Trois factures de la clinique d'octobre à décembre 2025 comptées deux fois (provision 2025 puis facture 2026).",
  "Grand livre 2025 : OD 805, 806, 807, 255 000 000 FCFA. Grand livre 2026 : pièces SIN 462, 552, 619, 240 615 119 FCFA (le réviseur cite 240 471 119, écart de 144 000 FCFA à expliquer). Aucun crédit en comptes 60 en 2026. S/C maladie 97,40 % devient 83,16 %, taux réglementaire 75,20 % devient 65,14 %.",
  CF, C),
 ("Sinistralité 2026", "B-V1", "82, 93", "À vérifier",
  "Facture générale de la clinique de février 2026 peut-être absente.",
  "Non tranché sans la liste des factures de janvier à août 2026.", NV, J),
 ("Sinistralité 2026", "B-V2", "85, 86", "À vérifier",
  "Décès : 55 600 000 FCFA dans l'état du FPPN au 30/06/2026 contre 152 816 666 FCFA en comptabilité au 31/08/2026.",
  "Rapprochement à faire sur l'état du FPPN.", NV, J),
 ("Ratio combiné, frais et hors soins", "A-M1, B-B2", "11, 13, 19, 24, 93, 95", "Bloquant",
  "Le ratio des soins seuls retire les dépenses hors soins au lieu de les reclasser en frais.",
  "(1 284 143 358 + 875 420 236 + 70,5 % × 487 492 003) / 2 409 281 030 = 103,90 %. Résultat technique maladie 2025 reclassé : -93 891 850 FCFA.", CF, C),
 ("Ratio combiné, frais et hors soins", "B-M1", "91, 93, 95", "Majeur",
  "Tous les frais du FPPN sont imputés à la maladie et au décès, dotations comprises, rien à l'épargne ni au Grenier.",
  "Dotations 6812 et 6813 en 2025 : 140 643 086 FCFA, nulles en 2026. Formule du rapport 2025 : 32,23 %.", CF, C),
 ("Ratio combiné, frais et hors soins", "E-M3", "192, 193, 245", "Majeur",
  "Le résultat technique 2025 a deux valeurs (351 429 674 et 211 376 588 FCFA).",
  "Frais 61 à 66 hors 6518 : 1 101 350 021 FCFA. Frais 612, 613, 618, 62 à 68 : 1 241 993 107 FCFA.", CF, C),
 ("Ratio combiné, frais et hors soins", "B-M9, E-V4", "91, 93, 96, 245", "Majeur",
  "Les frais de 2026 (44,12 %) portent des charges ponctuelles et sont comparés à des années pleines.",
  "65214001 : 82 300 000 contre 38 700 000 FCFA. 63421000 : 155 804 995 contre 78 486 000 FCFA. Dotations 68 nulles en 2026.", CF, C),
 ("Ratio combiné, frais et hors soins", "B-M2", "92, 93, 95", "Majeur",
  "Le classement par libellé laisse en soins des dépenses qui n'en sont pas.",
  "Classement du script reproduit au franc près (487 492 003 FCFA, 27,52 % en 2025). Règles absentes : GIS PLUS, SACS KABA, HADJ, LOBBING. Avec les règles élargies : +55 067 300 FCFA en 2025, part hors soins 30,62 % (le réviseur trouve 31,18 %).",
  CF, C),
 ("Ratio combiné, frais et hors soins", "B-m6", "92, 94", "Mineur",
  "Les dons passent du compte 6583 aux comptes de soins.",
  "6583 : 52 270 446 (2021), 61 894 670 (2022), 21 853 600 (2023), 12 792 675 (2024), 0 FCFA (2025).", CF, A),
 ("Ratio combiné, frais et hors soins", "B-m1, B-m2, B-m3", "81, 82, 83, 84, 90", "Mineur",
  "Compte 6012 cité avec deux montants. Formule écrite deux fois. 60800008 exclu de la maladie.",
  "6012 en 2025 : 989 729 545 + 4 881 810 + 700 000 = 995 311 355 FCFA. 60800008 : 18 750 000 (2023), 4 800 000 FCFA (2024).", CF, C),
 ("Ratio combiné, frais et hors soins", "B-m5, B-V3", "92", "Mineur",
  "Nature des honoraires trimestriels et des « honoraires médicaux » mensuels.",
  "Libellés « Me … / RECO » relevés par le réviseur. Nature à confirmer par le FPPN.", NV, J),
 ("Sinistralité historique et tendances", "B-B3", "91", "Bloquant",
  "« De 2020 à 2022, moins de 10 % des cotisations rendues en prestations. »",
  "Taux réglementaire : 8,92 % (2020), 20,77 % (2021), 9,06 % (2022).", CF, C),
 ("Sinistralité historique et tendances", "B-M4", "89", "Majeur",
  "TCAM maladie de 104,42 % calculé depuis 2021, année la plus basse.",
  "2020 à 2025 : 53,81 %. 2021 à 2025 : 104,42 %.", CF, C),
 ("Sinistralité historique et tendances", "B-M5", "98", "Majeur",
  "« Hausse statistiquement significative de 12,91 points par an. »",
  "Pente 12,91, t = 3,96 sur 6 points, résidus en U (15,88 ; -3,57 ; -16,01 ; -8,47 ; -0,16 ; 12,33). Pente 2023 à 2025 : 23,31 points par an.", CF, C),
 ("Sinistralité historique et tendances", "B-M6, B-m7, B-m9", "69, 70", "Majeur",
  "Tendance mensuelle des cotisants testée comme une droite alors qu'elle avance par paliers de juillet.",
  "Expression « statistiquement significative » présente diapositives 69 et 98. Série mensuelle non retestée ici.", PA, C),
 ("Sinistralité historique et tendances", "B-M3", "74, 81, 82, 83", "Majeur",
  "Le rapport applique la fourchette de 70 % à 80 % par garantie, ce qu'il reproche au FPPN.",
  "« fourchette réglementaire » présent diapositives 74, 81 et 82.", CF, C),
 ("Sinistralité historique et tendances", "B-M7", "73, 79", "Majeur",
  "« Trois clés de répartition coexistent. »",
  "Part maladie en comptabilité : 70,00 % de 2020 à 2023, 70,22 %, 70,50 %, 70,65 %. La comptabilité suit le barème 70 / 30.", CF, C),
 ("Sinistralité historique et tendances", "B-M8, B-V4", "76", "Majeur",
  "Écart comptabilité / précomptes de 6,89 % attribué à la clé.",
  "Le texte cite 2 587 432 980 FCFA de précomptes 936 en 2025. La base et le classeur donnent 2 529 485 250 FCFA, soit un écart de 4,75 %. Le chiffre du texte n'est pas celui du classeur.",
  CF, C),
 ("Sinistralité historique et tendances", "B-M10", "75, 79", "Majeur",
  "La baisse de la cotisation par adhérent n'est pas montrée.",
  "Arithmétique confirmée : 125 690 FCFA (2020) contre 107 884 FCFA (2025).", CF, A),
 ("Sinistralité historique et tendances", "B-m8, B-m10", "73, 87", "Mineur",
  "Taux de décès rapporté à l'effectif de fin d'année. Clé 3,52 / 1,48 arrondie.",
  "99 / 29 668 = 3,34 ‰ au lieu de 3,13 ‰. L'arrondi 3,52 / 1,48 est acceptable.", CF, C),
 ("Sinistralité historique et tendances", "A-m2, A-m3, E-m10", "13, 15, 84, 86, 193", "Mineur",
  "Cotisations et sinistralité décès 2025 calculées avec ou sans le compte MEDOC.",
  "70220000 (MEDOC) 2025 : 590 000 FCFA. 193 136 666 / 1 008 270 692 = 19,16 %.", CF, C),
 ("Sinistralité historique et tendances", "A-m5, E-m9, C-m2", "16, 132", "Mineur",
  "Multiplicateur des prestations sans période.",
  "3,90 de 2021 à 2025. 9,37 depuis 2020. 8,61 pour la maladie seule depuis 2020.", CF, C),
 ("Sinistralité historique et tendances", "A-m1, A-m4", "13, 39", "Mineur",
  "Résultat 2025 sans les produits HAO. Deux taux de frais (32,23 % et 36,34 %).",
  "82200000 : 10 000 000 FCFA, 84100000 : 6 831 858 FCFA.", CF, C),
 ("Provisions techniques", "A-M5, C-M1, E-M6, E-M7", "5, 6, 10, 24, 49, 131, 132, 156 à 158, 249, 254", "Majeur",
  "Le besoin maladie de 12 481 392 FCFA est une borne basse présentée comme une mesure. Chain Ladder écarté pour une raison discutable.",
  "Onglet 18c : 1 215 600 / 15,57 % × 1,599. Un mois de prestations maladie 2025 : 147 636 280 FCFA. « 230 fois » et « 1 103 fois » en diapositive 132.", CF, C),
 ("Provisions techniques", "C-M2", "136, 156 à 158", "Majeur",
  "Les décès tardifs (7 100 000 FCFA) viennent d'un seul état.",
  "Onglet 19c : les 6 dossiers viennent tous de l'état 2021.", CF, C),
 ("Provisions techniques", "C-M9, C-m4", "156, 158", "Majeur",
  "Total des engagements à dates mélangées, dettes maladie exclues mais dettes décès incluses.",
  "Arithmétique confirmée : avec le 40110000 et le besoin maladie, l'épargne pèse 94,06 % au lieu de 97,46 %.", CF, C),
 ("Provisions techniques", "C-m1, E-M13, A-m6", "6, 20, 129, 130, 199", "Mineur",
  "Délai de paiement 2020 divisé par 12 mois. Deux séries de délais non rapprochées.",
  "Crédits du 40110000 en 2020 : 313 578 298 FCFA sur 4 mois, délai 17,79 mois au lieu de 53,38.", CF, C),
 ("Provisions techniques", "C-V3", "135", "À vérifier",
  "Dette décès de 489 419 669 FCFA dont 257 683 332 FCFA de dossiers antérieurs à 2021, peut-être forclos.",
  "Dates de dépôt non reçues.", NV, J),
 ("Provisions techniques", "C-M11, E-M5", "4, 127, 137, 242, 244, 255, 256", "Majeur",
  "Aucun ordre de grandeur pour la santé des retraités ni pour la mortalité par âge approché.",
  "Données disponibles : 610 retraités, 105 dates de naissance, cotisation de 11 954 FCFA par mois.", CF, A),
 ("Provisions techniques", "C-m8", "127", "Mineur",
  "« Quatre provisions techniques » et « montant hérité d'environ 2004 ».",
  "Texte du règlement d'exécution n° 004/2022 non relu ici.", NV, C),
 ("Provision du personnel", "A-M10, C-B3, C-m3, E-m7", "5, 6, 11, 51, 127, 144, 146, 155, 158, 216", "Bloquant",
  "« Provision suffisante » et « aucune donnée ne peut porter la provision au-dessus du compte 196 », contredit par le tableau de la même diapositive.",
  "Diapositive 155 : 72 102 465 FCFA sans rotation, au-dessus du compte 196 (43 666 678 FCFA). « Méthode non transmise » encore en diapositives 51, 127, 158, 216.", CF, C),
 ("Provision du personnel", "C-M6", "143, 155", "Majeur",
  "Croissance des salaires de 10,36 % par an jusqu'au départ, actualisation à 7,19 %.",
  "Onglet 14a : croissance 10,357 %, rotation 9,79 %. Variantes à 4 % et 5 % non recalculées ici.", PA, C),
 ("Provision du personnel", "C-M7", "142 à 155", "Majeur",
  "Charges patronales listées mais absentes du calcul.",
  "Onglet 14b : formule sans charges sociales (79 agents, 17 184 450 FCFA).", CF, C),
 ("Provision du personnel", "C-V5, A-m7, C-m10", "12, 145, 147, 148", "Mineur",
  "Dates des effectifs non précisées, renvoi erroné à la recommandation 9, sous-titre répété.",
  "Constats de forme confirmés à la lecture.", CF, C),
 ("Grenier", "A-M14, C-m7", "8, 15, 112, 118", "Majeur",
  "« Financement des prêts du Grenier : 915 477 260 FCFA » alors que le compte comprend la dotation de 500 000 000 FCFA.",
  "Libellé à corriger. Montant au 31/08/2026 non rapproché ici.", PA, C),
 ("Grenier", "C-M4", "139", "Majeur",
  "Le compte de liaison n'est égal qu'en 2025.",
  "Balances du Grenier et du FPPN : écarts de 67 280 924 (2022), 87 274 688 (2023), 53 178 248 (2024) et 0 FCFA (2025).", CF, C),
 ("Grenier", "C-M3", "140", "Majeur",
  "« Le résultat du Grenier disparaîtrait avec la dépréciation. »",
  "4 746 129 / 30 543 376 = 15,54 % du résultat 2025 : il reste positif.", CF, C),
 ("Grenier", "C-M10", "114, 116, 118", "Majeur",
  "« Retards maîtrisés » alors que 15 à 24 % des échéances ne sont pas précomptées.",
  "Mot présent en diapositive 118. Rééchelonnements non testés.", PA, J),
 ("Grenier", "D-V3", "167, 172", "À vérifier",
  "Épargne inscrite au Grenier (8 493 055 501 FCFA) peut-être comptée deux fois avec celle du FPPN.",
  "Compte 37911400 du Grenier : 8 493 055 501 FCFA au 31/12/2025.", NV, J),
 ("Épargne", "C-m9", "102", "Mineur",
  "« Les épargnants PEP passent de 1 327 en 2020 à 17 199 en 2026 » sans dire que 2020 vient de fichiers partiels.",
  "Base des précomptes : poste 998, 1 327 matricules en 2020 et 17 199 en 2026, pour un PEP de 7 366 229 747 FCFA fin 2020.", CF, C),
 ("Épargne", "C-V4", "141", "À vérifier",
  "Produits annexes 2025 (43 668 069 FCFA) très supérieurs au débit GNA (837 513 FCFA).",
  "Compte 71800000 non analysé au grand livre 2025.", NV, J),
 ("Grenier", "C-m5, C-m6, C-m11", "115, 138, 139", "Mineur",
  "Risque de crédit dit non mesuré alors qu'il l'est. Mécano du 4e prêt. Prêts de plus de 24 mois.",
  "Constats de forme confirmés à la lecture.", CF, C),
 ("Ratios prudentiels", "A-M7, D-M5", "7, 9, 10, 16, 19, 20, 24, 202, 207, 208, 229", "Majeur",
  "La lecture corrigée met toute l'épargne en ressource longue, en contradiction avec le taux de sortie de l'adossement.",
  "Arithmétique du réviseur reprise (liquidité 0,45, équilibre stable 1,11 en 2025). Entrées non relues une à une.", PA, C),
 ("Ratios prudentiels", "D-M6", "195, 204, 215, 221", "Majeur",
  "Statut de la marge changeant d'une diapositive à l'autre. Explication par le compte 106 inexacte.",
  "Décision n° 05/2022 : marge = réserve obligatoire, fonds d'établissement, réserves libres, droits d'adhésion, dons et legs. Sans le 106 : 2,64, conforme.", CF, C),
 ("Ratios prudentiels", "D-M7", "31, 190, 195, 220, 221, 223, 269", "Majeur",
  "Le risque de retrait d'agrément est rattaché à des ratios de gestion (tableau 2).",
  "Décision n° 05/2022 : le tableau 1 regroupe les ratios « de contrôle » dont le non-respect « entraîne des sanctions ». Le tableau 2 sert au suivi de la gouvernance et de la gestion.", CF, C),
 ("Ratios prudentiels", "D-M8", "214, 217", "Majeur",
  "Affectation du résultat 2025 non examinée.",
  "Au 31/08/2026 : 11100000 inchangé à 820 232 819 FCFA, résultat 2025 de 437 899 911 FCFA au 12100000. Minimum décès selon le réviseur : 224 575 578 FCFA.", CF, J),
 ("Ratios prudentiels", "D-M9", "225, 227, 230, 235", "Majeur",
  "« 16 002 055 145 FCFA de fonds propres manquent » : objectif inatteignable pour une mutuelle d'épargne.",
  "Montant présent diapositives 225, 227, 235.", CF, C),
 ("Ratios prudentiels", "D-M10, E-M12, D-m5", "182, 197, 234, 270, 273", "Majeur",
  "Rendements 2022 à 2024 faibles parce que les intérêts du DAT PEP sont tous passés en 2025.",
  "Classe 77 : 32 409 354 (2021), 7 258 416 (2022), 7 124 778 (2023), 9 176 280 (2024), 117 912 488 FCFA (2025). DAT PEP : 638 047 119 - 505 383 858 = 132 663 261 FCFA.", CF, C),
 ("Ratios prudentiels", "E-M11", "268, 271", "Majeur",
  "Ratio FEP : provisions techniques au numérateur et au dénominateur.",
  "Onglet 11_formules, F-10-04 : « 10 + 11 + 12 + 13 + 14 + 190 + 15 / classe 4 créditrice + 16 + 190 + 15 + 5 créditeur ». Ratio non recalculé.", CF, C),
 ("Ratios prudentiels", "D-m1, D-m2, D-m3, D-m6, D-m7, D-m8, D-m9", "189, 194, 195, 201, 205, 208, 218, 223, 225, 230, 234", "Mineur",
  "Réserve en mois non annualisée, montants et sources à corriger, norme ambiguë (6 à 9 mois).",
  "3,65 mois annualisé contre 5,47. 11100000 = 820 232 819 FCFA. 47111103 = 8 357 475 FCFA. La Décision donne [6 mois ; 9 mois] au tableau 1 et 9 mois au § 1.2.2.5.4.", CF, C),
 ("Ratios prudentiels", "A-V3", "20, 33", "À vérifier",
  "Recouvrement des retraités au moins égal à 90 % avec un dénominateur incomplet.",
  "Liste des souscripteurs retraités non reçue.", NV, J),
 ("Projection", "E-B1", "245, 246, 50", "Bloquant",
  "La projection ignore les huit mois de 2026, déjà déficitaires.",
  "2026 : cotisations 2 392 197 346, prestations 1 798 855 783, frais 1 055 460 837, résultat technique -462 119 274 FCFA (-693 178 911 FCFA sur 12 mois). Frais annualisés 1 583 191 256 FCFA. « deux bornes » et « la cible » en diapositives 173 et 246.",
  CF, C),
 ("Projection", "E-M2", "245", "Majeur",
  "Le scénario A prolonge 5 ans un TCAM de 72,21 % calculé sur deux variations.",
  "Hausse maladie 2026 annualisée : +39,37 %.", CF, C),
 ("Projection", "E-M4", "243, 245", "Majeur",
  "L'épargne, premier engagement, est exclue de la projection.",
  "Constat de méthode confirmé à la lecture.", CF, A),
 ("Tables de mortalité", "E-M1", "238 à 242, 296", "Majeur",
  "Les tables CIMA sont décrites comme des tables de population générale.",
  "Table ronde CIMA FANAF 2011 : tables d'une « population assurée », mortalité « inférieure d'environ 70% » à la population générale, âges exploitables de 30 à 55 ans. « Population générale de la zone CIMA » en diapositives 239, 240, 296.", CF, C),
 ("Tables de mortalité", "E-m1, E-m2, E-m3, E-m4, E-V3", "239, 240, 241, 242", "Mineur",
  "Rente exprimée en années, comparaison par e0, écarts « sur l'ensemble des âges », taux de 3,5 % sans sensibilité.",
  "Table ronde : « minorent de 17% » (CIMA H) et « inférieure de 4% » (CIMA F), sans plage d'âges.", CF, C),
 ("Tables de mortalité", "E-m15, E-m16, E-V1, E-V2", "155, 222, 256, 269, 277, 278, 299", "Mineur",
  "IAS 19 § 145 et Code CIMA art. 338 cités hors de leur champ.",
  "Citations présentes : IAS 19 § 145 en diapositives 155, 222, 277. Articles du Code CIMA non vérifiés (texte absent du dossier).", PA, C),
 ("Résumé et cohérence documentaire", "A-B2, A-M11", "3 à 27, 169", "Bloquant",
  "Aucun fait de novembre 2026 dans le résumé exécutif, rien sur Azaguié dans les risques.",
  "Recherche de 12 chiffres et mots de novembre (53,30, 119,07, 487 492 003, 1 280 300 860, 379 204 261, 70 191 500, 8 493 055 501, 96,33, 27,19, Azaguié, SOMAVIE, 263 000 000) dans les diapositives 3 à 27 : aucune occurrence. Azaguié en diapositives 169 et 251 seulement.",
  CF, A),
 ("Résumé et cohérence documentaire", "A-M12, E-M14", "4, 9, 42, 43, 47, 53, 261 à 263, 299", "Majeur",
  "Dates de validité périmées.",
  "« 22/10/2026 » dans 48 diapositives, « au 23/10/2026 » dans 8. Classeur « FPPN_Calculs_rapport_20261022.xlsx » cité en diapositive 299.", CF, C),
 ("Résumé et cohérence documentaire", "A-M13", "42", "Majeur",
  "Verdicts de la diapositive 42 différents du résumé.", "« recalcul en attente » présent en diapositive 42.", CF, C),
 ("Résumé et cohérence documentaire", "A-M2", "13, 19, 21", "Majeur",
  "Taux réglementaire 2026 de 75,20 % affiché conforme.", "Voir sinistralité 2026 : 65,14 % après correction.", CF, C),
 ("Résumé et cohérence documentaire", "A-m8, A-m9, A, diapositives 52 et 54", "9, 32, 43, 45, 49, 52, 54", "Mineur",
  "« Quatre apportent » pour cinq documents, périodes des précomptes divergentes, scripts v9 cités, limite du contrôle de la diapositive 54.",
  "« Quatre apportent » en diapositive 45. rapport_v9 et donnees_v9 cités en diapositive 52.", CF, C),
 ("Résumé et cohérence documentaire", "E-M8, E-M9", "248, 249, 254", "Majeur",
  "Liste des analyses en attente incomplète, statuts des données inversés.",
  "« Quatre analyses » en diapositive 249.", CF, C),
 ("Résumé et cohérence documentaire", "E-M10, E-M15", "264 à 279", "Majeur",
  "Légende du tableau des formules (« norme respectée ») et formules manquantes.",
  "« norme respectée » présent dans les diapositives 265 à 279.", CF, C),
 ("Résumé et cohérence documentaire", "E-m5, E-m6, E-m7, E-m8, E-m11, E-m12, E-m13, E-m14, E-m17", "7, 120, 122, 200, 216, 242, 249, 251, 256, 261, 262, 270 à 279", "Mineur",
  "Renvois PBC, périodes, formules et statuts du tableau des formules.",
  "Constats de forme confirmés à la lecture.", CF, C),
]

wb = Workbook()
ws = wb.active
ws.title = "Registre"
noir, gris, orange = "060606", "BBBCBC", "EFAC62"
f = Font(name="Helvetica", size=10)
fb = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
bd = Border(*(Side(style="thin", color=gris),) * 4)
ws.append(["Registre de vérification de la revue indépendante du rapport du 11/11/2026 (13/11/2026)"])
ws["A1"].font = Font(name="Helvetica", size=13, bold=True)
ws.append(["Vérification sur les balances, les grands livres 2023 au 31/08/2026, le classeur du 11/11/2026 recalculé et le texte du rapport (scripts verif_revue_fppn.py, regle_interets_pep_fppn.py). Aucune modification du rapport avant validation."])
ws["A2"].font = Font(name="Helvetica", size=9, italic=True)
ws.append([])
H = ["N°", "Thème", "Constats (réviseur)", "Diapositives", "Gravité", "Constat", "Vérification et chiffre recalculé", "Résultat", "Décision proposée"]
ws.append(H)
for j, _ in enumerate(H, 1):
    c = ws.cell(4, j)
    c.font, c.fill, c.alignment, c.border = fb, PatternFill("solid", fgColor=noir), Alignment(wrap_text=True, vertical="center"), bd
for i, l in enumerate(L, 1):
    ws.append([i, *l])
    for j in range(1, len(H) + 1):
        c = ws.cell(4 + i, j)
        c.font, c.alignment, c.border = f, Alignment(wrap_text=True, vertical="top"), bd
    if l[3] == "Bloquant":
        ws.cell(4 + i, 5).fill = PatternFill("solid", fgColor=orange)
for col, w in zip("ABCDEFGHI", (5, 20, 18, 18, 11, 48, 70, 14, 18)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:I{4 + len(L)}"

s = wb.create_sheet("Synthèse")
s.append(["Décision proposée", "Lignes"])
for k, v in Counter(l[7] for l in L).most_common():
    s.append([k, v])
s.append([])
s.append(["Résultat de la vérification", "Lignes"])
for k, v in Counter(l[6] for l in L).most_common():
    s.append([k, v])
for r in s.iter_rows():
    for c in r:
        c.font = f
for r in (1, 2 + len(Counter(l[7] for l in L)) + 1):
    for c in s[r]:
        c.font, c.fill = fb, PatternFill("solid", fgColor=noir)
s.column_dimensions["A"].width = 30
wb.save(SORTIE)
refs = sum(len([x for x in l[1].replace(" et test du 13/11/2026", "").split(",") if x.strip()]) for l in L)
print(len(L), "lignes,", refs, "références de constats")
print(Counter(l[7] for l in L), Counter(l[6] for l in L))
